Vacation Responder 
by
Waqar Ahamd (support.waqar@gmail.com)


________________


Table of Contents




SBC Technology
Document Revison History
Idea Source
API References
Introduction
Installation Notes
Updating the Vacation Responder setting for domain users
Demo:




________________
Document Revison History


Rev. #
	Comments
	Updated by
	Date
	1
	Initial draft
	Waqar Ahmad
	Nov 4, 2012
	2
	Added Introduction section
	Waqar Ahmad
	Nov 5, 2012
	

________________
Idea Source
http://googleappsdeveloper.blogspot.in/2012/01/automatically-setting-vacation.html


API References
1. Gmail Setting API - Manage Vaction Responder Settting 
2. Provisioning and E-mail Migration APIs - Retreive USer accounts
3. Google Calendar API - Event List


Vacation responder script will work only with Google Apps Administrator account. Administrator will be able to run the script through time driven trigger on behalf of domain users.


Script can be published to the script gallery.


Introduction
This application will be used by Google Apps Administrator or by a Domain User who has admin access. 
Sole purpose of the application is to check domain user’s calendars and find if there are any vacation events. (Event must have word ‘vacation’ either in event title or event description and it should be an all day event). If the application finds a vacation event on users calendar, it will set the domain user’s vacation responder setting for that event period and notifies the user..


Installation Notes
Until the script is published in script gallery one needs to follow these steps to install the application.
1. Open your spreadsheet, and go to Script Editor.
Path : Your Spreadsheet > Tools > Script Editor
2. Now copy the Vacation Responder Application Code from the source.
3. Now Save the script, It will ask you for a name. Give it any name.
4. In script editor go to Run > onInstall, This will ask you to authorize the application. Authorize it. This will ask you to authorize more than once as it will ask for authorization for different APIs one by one. Run onInstall multiple time until it runs successfully.
5. once onInstall Function is run successfully, a sheet with name “Vacation Responder” in your spreadsheet will appear, and open a Popup panel for Configurations.
6. Now come to your spreadsheet and enter configuration parameters and Save.
7. In you spreadsheet, Go to SBC Tech > Update user list
This will list all the domain users in the Vacation Responder sheet.


Configuration Parameters:


OAuth consumer key
	if you are Google Apps admin user, leave it as “anonymous”, if you are not then ask your administrator for oAuth consumer key. usually, it is the name of the domain
	oAuth comsumer secret
	if you are Google Apps admin user, leave it as “anonymous”, if you are not then ask your administrator for oAuth consumer secret. 
	API Key
	Here are the steps to get the API Key
Go to Google API console > Create an API project
In the API project, go to services > Turn on the Calendar API
Now go to API access > Look for Simple API access, API key is given there
	Subject
	Subject for the Vacation Responder mail
	Message
	Message for the Vacation Responder mail
	

Updating the Vacation Responder setting for domain users
   1. Now go to SBC Tech > Update Vacation Responder setting to set the vacation Responders setting for all users manually.
   2. if you want the Vacation Responder setting to be update automatically, Go to Script Editor > Resources > Current Script Triggers > Add a new Trigger with following details
updateVacationResponder - Time-diven - Day timer - Midnight to 1am
Demo:
      1. Demo video link
	http://youtu.be/bHl1rQ6K_bw