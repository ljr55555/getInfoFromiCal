import os.path

# Requirements for Google iCal download
import requests
from icalendar import Calendar as iCalendar, Event

# Requirements for ExchangeLib
from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, FaultTolerance, \
	Configuration, NTLM, GSSAPI, SSPI, Build, Version, CalendarItem, EWSDateTime, EWSTimeZone

from exchangelib.folders import Calendar
from exchangelib.items import MeetingRequest, MeetingCancellation, SEND_TO_ALL_AND_SAVE_COPY
from exchangelib.protocol import BaseProtocol

from urllib.parse import urlparse
import requests.adapters

# Used to decrypt password from config file
from cryptography.fernet import Fernet
from base64 import b64encode, b64decode

# Misc imports
from datetime import datetime, timedelta
import dateutil.parser as parser
import pytz

from config import strUsernameCrypted, strPasswordCrypted, strEWSHost, strPrimarySMTP, dictCalendars, iMaxExchangeResults, strKey

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

# Set our time zone
local_tz = pytz.timezone('US/Eastern')

# This needs to be the key used to stash the username and password values stored in config.py
f = Fernet(strKey)
strUsername = f.decrypt(strUsernameCrypted)
strUsername = strUsername.decode("utf-8")

strPassword = f.decrypt(strPasswordCrypted)
strPassword = strPassword.decode("utf-8")

def calculateDaysUntil(d1):
	try:
		d1 = datetime.strptime(d1, "%Y%m%d")
	except:
		try:
			d1 = datetime.strptime(d1, "%Y%m%dT%H%M%SZ")
		except:
			d1 = datetime.strptime(d1, "%Y%m%dT%H%M%S")

	return ((  d1 - datetime.now() ).days)

def createExchangeItem(objExchangeAccount, strTitle, strLocn, strStartDate, strEndDate, strInviteeSMTP=None):
	iDaysUntilItem = calculateDaysUntil(strStartDate)
	if iDaysUntilItem > 0:
		print("Creating item {} which starts on {} and ends at {}".format(strTitle, strStartDate,strEndDate))
		objStartDate = parser.parse(strStartDate)
		objEndDate = parser.parse(strEndDate)
		if strInviteeSMTP is None:
			item = CalendarItem(
				account=objExchangeAccount,
				folder=objExchangeAccount.calendar,
				start=objExchangeAccount.default_timezone.localize(EWSDateTime(objStartDate.year, objStartDate.month, objStartDate.day, objStartDate.hour, objStartDate.minute)),
				#start=EWSDateTime.from_string(strStartDate),
				end=objExchangeAccount.default_timezone.localize(EWSDateTime(objEndDate.year,objEndDate.month,objEndDate.day,objEndDate.hour,objEndDate.minute)),
				#end=(EWSDateTime(objEndDate.year,objEndDate.month,objEndDate.day,objEndDate.hour,objEndDate.minute,0,0,EWSTimeZone.timezone('America/New_York'))),
				#end=EWSDateTime.from_string(strEndDate),
				subject=strTitle,
				reminder_minutes_before_start=60,
				reminder_is_set=True,
				location=strLocn,
				body=""
			)
		else:
			item = CalendarItem(
				account=objExchangeAccount,
				folder=objExchangeAccount.calendar,
				start=objExchangeAccount.default_timezone.localize(EWSDateTime(objStartDate.year, objStartDate.month, objStartDate.day, objStartDate.hour, objStartDate.minute)),
				end=objExchangeAccount.default_timezone.localize(EWSDateTime(objEndDate.year,objEndDate.month,objEndDate.day,objEndDate.hour,objEndDate.minute)),
				subject=strTitle,
				reminder_minutes_before_start=60,
				reminder_is_set=True,
				location=strLocn,
				body="",
				required_attendees=[strInviteeSMTP]
			)
		item.save(send_meeting_invitations=SEND_TO_ALL_AND_SAVE_COPY )
	#else:
	#	print(f"Event {strTitle} is in the past, so not being created")


def utc_to_local(utc_dt):
	interim_dt = utc_dt.replace(tzinfo=None)
	local_dt = local_tz.localize(interim_dt)
	return local_tz.normalize(local_dt) 

def main():
	class RootCAAdapter(requests.adapters.HTTPAdapter):
		# An HTTP adapter that uses a custom root CA certificate at a hard coded location
		def cert_verify(self, conn, url, verify, cert):
			cert_file = {
				'exchange01.rushworth.us': './ca.crt'
			}[urlparse(url).hostname]
			super(RootCAAdapter, self).cert_verify(conn=conn, url=url, verify=cert_file, cert=cert)

	#Use this SSL adapter class instead of the default
	BaseProtocol.HTTP_ADAPTER_CLS = RootCAAdapter

	# Get Exchange calendar events and save to dictEvents 
	dictEvents = {}

	credentials = Credentials(username=strUsername, password=strPassword)
	config = Configuration(server=strEWSHost, credentials=credentials)
	account = Account(primary_smtp_address=strPrimarySMTP, config=config, autodiscover=False, access_type=DELEGATE)
	
	print("Starting to check Exchange")
	for item in account.calendar.all().order_by('-start')[:iMaxExchangeResults]:
		if item.start:
			objEventStartTime = parser.parse(str(item.start))
			objEventStartTime = utc_to_local(objEventStartTime)
	
			strEventKey = "{}{:02d}-{:02d}-{:02d}".format(str(item.subject), int(objEventStartTime.year), int(objEventStartTime.month), int(objEventStartTime.day))
			strEventKey = strEventKey.replace(" ","")
			dictEvents[strEventKey]=1

	print("Starting to check Google calendar ...")

	for strCalendarName, strCalURI in dictCalendars.items():
		print(f"I am getting the calendar {strCalendarName} from {strCalURI}")
		# Grab the iCal file
		objIcalData = requests.get(strCalURI)
		gcal = iCalendar.from_ical(objIcalData.text)
		for component in gcal.walk():
			strSummary = f"{strCalendarName}: {(component.get('summary'))}"
			strDesc = component.get('description')
			strLocation = component.get('location')
			dateStart = component.get('dtstart')
			if dateStart is not None:
				dateStart = dateStart.to_ical().decode()
				
			dateEnd = (component.get('dtend'))
			if dateEnd is not None:
				 dateEnd = dateEnd.to_ical().decode()
			else:
				dateEnd = dateStart

			strSummary = strSummary.replace("BZA ","Board of Zoning Appeals ")
			strSummary = strSummary.replace("ZC Meeting","Zoning Commission Meeting")
			
			strThisEventKey = strSummary + (str(dateStart).split('T'))[0]
			strThisEventKey = strThisEventKey.replace(" ","")
			if (strThisEventKey not in dictEvents) and (strSummary is not None) and (dateStart is not None) and (dateStart[0:4] != "1970"):
				createExchangeItem(account, strSummary, strLocation, dateStart, dateEnd)
			else:
				print(f"The event {strThisEventKey} on {dateStart} already exists in the calendar or has no valid start date.")

if __name__ == '__main__':
	main()
