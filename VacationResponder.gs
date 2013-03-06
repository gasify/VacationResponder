//    Script developed by Waqar Ahmad (support.waqar@gmail.com)
//    Website : https://sites.google.com/site/appsscripttutorial/


//*****************************************************************************************************
//Global Variables
var domain = UserManager.getDomain();
var vrSheetName = 'SBC VacationResponder';
var timeZone = Session.getTimeZone()
//*****************************************************************************************************

function updateVacationResponder(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(vrSheetName);
  var lastRow = sheet.getLastRow();
  //get the old data
  var oldData = sheet.getRange(2, 1, lastRow-1, 6).getValues();
  var newData = [];
  for(var i in oldData){
    //If not in exclusion list, update Vacation responder
    if(oldData[i][5] != 'yes'){
      var vacationTime = getVacationTime_(oldData[i][0]);
      if(vacationTime.startDate != undefined){
        newData.push([oldData[i][0], oldData[i][1], vacationTime.startDate, vacationTime.endDate, new Date(), '']);
        
        var oldSartDate = typeof oldData[i][2] != 'string'? Utilities.formatDate(oldData[i][2], timeZone, 'yyyy-MM-dd'): oldData[i][2];
        var oldEndDate = typeof oldData[i][3] != 'string'? Utilities.formatDate(oldData[i][3], timeZone, 'yyyy-MM-dd') : oldData[i][3];
        
        //if the date has changed, send notification to user
        
        if(oldSartDate != vacationTime.startDate || oldEndDate != vacationTime.endDate){
          //set the vacation responder
          //get these values from ScriptProperties
          var properties = ScriptProperties.getProperties();
          var vacationResponder = {
            contactsOnly : properties.contactsOnly,
            domainOnly : properties.domainOnly,
            enable : 'true',
            subject : properties.vrSubject,
            message : properties.vrMessage,
            startDate : vacationTime.startDate,
            endDate : vacationTime.endDate
          };
          
          var updated = updateVacationResponder_(oldData[i][0].split('@')[0], vacationResponder);
          
          //Send notification to user
          GmailApp.sendEmail(oldData[i][0] , '[Administrator] : Vacation responder set', 
                             'Vacation responder has been set for the days between '+vacationTime.startDate+' and '+vacationTime.endDate,
                             {noReply : true});
        }
        
      }
      else{
        newData.push([oldData[i][0], oldData[i][1], 'NA', 'NA', new Date(), '']);
      }
      
    }
    else {
      newData.push([oldData[i][0], oldData[i][1], '', '', '', 'yes']);
    }
  }
  sheet.getRange(2, 1, newData.length, 6).setValues(newData);
}


//*****************************************************************************************************
// This function will insert a sheet in active spreadsheet and will apply some formatting
function onInstall(){
  //Calling ALL the APIs so that it will prompt for authorization.
  try{
    getVacationTime_('user');//Uses Calendar API
    updateVacationResponder_('user', 'vacationResponder');//Uses Email setting API
    getUsers_(); //Uses Provisioning API
  }catch(e){};
  
  //Insert a sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Preparing the sheet");
  insertVacationResponderSheet_(ss);
  
  //some code to prompt the user for apiKey, consumerKey and consumerSecret
  openConfigPanel(ss);

  //Show custom menus
  onOpen();
  
  ss.toast("Setup complete");
}

function insertVacationResponderSheet_(ss){
  var sheet = ss.insertSheet(vrSheetName);
  sheet.setFrozenRows(1);
  var headerData = [['Username', 'Full Name', 'VacationResponder Start Date', 'VacationResponder End Date', 'Last Updated', 'Ignore this user (\'yes\'/<blank>)']];
  var headerRange = sheet.getRange(1, 1, 1, headerData[0].length);
  headerRange.setValues(headerData).setFontStyle('oblique')
    .setBackgroundColor('#92c47d').setBorder(true, true, true, true, true, true);
}
//*****************************************************************************************************


//*****************************************************************************************************
//Custom menus
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [ {name: "Update Vacation Responder setting", functionName: "updateVacationResponder"},                     
                    {name: "Update user list", functionName: "listUsers"},
                    {name: "Initial Configuration", functionName: "openConfigPanel"}];
  ss.addMenu("SBC Tech", menuEntries);
}
//*****************************************************************************************************




//*****************************************************************************************************
//This function will return start and end date of icoming all-day calendar event which is a result of
//free text search for any field for term 'vacation', except for extended properties
// e.g {startDate=2012-11-02, endDate=2012-11-03}
function getVacationTime_(user){
  var vacation = {};
  var base = 'https://www.googleapis.com/auth/calendar';
  //Get Google oAuth Arguments
  var fetchArgs = googleOAuthDomain_('calendar', base);
  fetchArgs.method = 'GET';
  var apiKey = ScriptProperties.getProperty('apiKey'); //Need to be generated from API Console
  //Feed URL
  var cur_time = Utilities.formatDate(new Date(), Session.getTimeZone() , "yyyy-MM-dd'T'HH:mm:ssZ");
  
  //url to get only those calendars which have word 'vacation'
  var url = 'https://www.googleapis.com/calendar/v3/calendars/'+encodeURIComponent(user)+'/events?key='+apiKey
      +'&singleEvents=true'
      +'&q=vacation'
      +'&timeMin='+encodeURIComponent(cur_time)
      +'&orderBy=startTime';
  try{
    var urlFetch = UrlFetchApp.fetch(url, fetchArgs);
    var vacationEvents = Utilities.jsonParse(urlFetch.getContentText()).items;
    //If no vacation found, make it zero length array
    if(vacationEvents == undefined){
      vacationEvents = [];
    }
    
    for(var i in vacationEvents){
      //if it is all day calendar event, return it;
      if(vacationEvents[i].start.date != undefined){
        vacation.startDate = vacationEvents[i].start.date;
        vacation.endDate = vacationEvents[i].end.date;
        break;
      }
    }
    return vacation;
  }
  catch(e){
    return 'error occured: '+ e;
  }
}
//*****************************************************************************************************





//*****************************************************************************************************
//This function will update the vacation responder setting for a user
//updateVacationResponder
/*parameters
user type string
vacationResponder type object
*/
function updateVacationResponder_(user, vacationResponder) {
  //Make the following three parameters false if not set true by the calling function
  vacationResponder.contactsOnly = vacationResponder.contactsOnly || 'false';
  vacationResponder.domainOnly = vacationResponder.domainOnly || 'false';
  vacationResponder.enable = vacationResponder.enable || 'false';
  
  var base = 'https://apps-apis.google.com/a/feeds/emailsettings/2.0/';
  
  //Build the xml
  var xmlRaw = '<?xml version="1.0" encoding="utf-8"?>'+
      '<atom:entry xmlns:atom="http://www.w3.org/2005/Atom" xmlns:apps="http://schemas.google.com/apps/2006">'+
      '<apps:property name="enable" value="'+vacationResponder.enable+'" />'+
      '<apps:property name="subject" value="'+vacationResponder.subject+'" />'+
      '<apps:property name="message" value="'+vacationResponder.message+'" />'+
      '<apps:property name="contactsOnly" value="'+vacationResponder.contactsOnly+'" />'+
      '<apps:property name="domainOnly" value="'+vacationResponder.domainOnly+'" />'+
      '<apps:property name="startDate" value="'+vacationResponder.startDate+'" />'+
      '<apps:property name="endDate" value="'+vacationResponder.endDate+'" />'+
      '</atom:entry>';
  
  try{
    var fetchArgs = googleOAuthDomain_('emailSetting',base);
    fetchArgs.method = 'PUT';
    fetchArgs.payload = xmlRaw;
    fetchArgs.contentType = 'application/atom+xml';
    var url = base+domain+'/'+user+'/vacation';
    var urlFetch = UrlFetchApp.fetch(url, fetchArgs);
    return vacationResponder;
  }
  catch(e){
     return 'error occured: '+ e;
  }
  
}

//*****************************************************************************************************


////*****************************************************************************************************
//This function will list the domain users in Vacation Responder sheet
function listUsers(sheet){
  var ss;
  if(sheet == undefined){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    try {sheet = ss.getSheetByName(vrSheetName);}
    catch(e){ insertVacationResponderSheet_(ss);
             sheet = ss.getSheetByName(vrSheetName);}
  }
  else ss = sheet.getParent();
  ss.toast("Fetching user list");
  //Get the new list of users
  var users = getUsers_(); //returned data is a 2D array
  
  var newUserList = [];
  //get the users from spreadsheet if any
  var oldUserList;
  try{
    oldUserList = sheet.getRange(2, 1, sheet.getLastRow()-1, 6).getValues();
  } catch(e){oldUserList = []};
  
  if (oldUserList.length > 0){
    //Compare users and push them in new userList with user data
    for(var i in users){
      var userFound = false;
      for(var j in oldUserList){
        if(users[i][0] == oldUserList[j][0]){
          newUserList.push(oldUserList[j]);
          userFound = true;
          break;
        }
      }
      if(!userFound){
        newUserList.push([users[i][0], users[i][1], '', '', '', '']);
      }
    }
    sheet.deleteRows(2, sheet.getLastRow()-1);
    sheet.getRange(2, 1, newUserList.length, 6).setValues(newUserList);
  }
  else {
    sheet.getRange(2, 1, users.length, users[0].length).setValues(users);
  }
  SpreadsheetApp.flush();
  ss.toast("Fetching user list complete");
}
//*****************************************************************************************************



//*****************************************************************************************************
//This function will return an 2D array of users 
/*  [['userName1', 'fullName1'],['userName2', 'fullName2'], ... ]  */
function getUsers_() {  
  try {
    //Get the domian name from
    var base = "https://apps-apis.google.com/a/feeds/";
    var fetchArgs = googleOAuthDomain_("provisioning", base);
    var url = base + domain+ "/user/2.0?alt=json";
    var result = UrlFetchApp.fetch(url, fetchArgs);
    var text = result.getContentText();
    var jsonResponse = Utilities.jsonParse(text);
    var userEntries = jsonResponse.feed.entry;
    var users = [];
    for(var i in userEntries){
      var temp = [];
      //userName
      temp.push(userEntries[i].apps$login.userName+'@'+domain);
      //fullName, i.e givenName+familyName
      temp.push(userEntries[i].apps$name.givenName + ' ' + userEntries[i].apps$name.familyName);
      users.push(temp);
    }
  return users;
  }
  catch(e){
    return 'error occured: '+ e;
  }
}
//*********************************************************************************************************



//*********************************************************************************************************
//Google OAuth, Called by functions which need oAuth for domainAPIs
function googleOAuthDomain_(name,scope) {
  var oAuthConfig = UrlFetchApp.addOAuthService(name);
  oAuthConfig.setRequestTokenUrl("https://www.google.com/accounts/OAuthGetRequestToken?scope="+scope);
  oAuthConfig.setAuthorizationUrl("https://www.google.com/accounts/OAuthAuthorizeToken");
  oAuthConfig.setAccessTokenUrl("https://www.google.com/accounts/OAuthGetAccessToken");
  var consumerKey = ScriptProperties.getProperty("consumerKey");
  var consumerSecret = ScriptProperties.getProperty("consumerSecret");
  oAuthConfig.setConsumerKey(consumerKey);
  oAuthConfig.setConsumerSecret(consumerSecret);
  return {oAuthServiceName:name, oAuthUseToken:"always"};
}
//*********************************************************************************************************


//*********************************************************************************************************
//This function will open a UI panel in spreadsheet and ask for configuration Parameters
function openConfigPanel(ss){
  if(ss == undefined){
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  var app = UiApp.createApplication().setTitle("Configurations");
  var scrollPanel = app.createScrollPanel().setSize('100%', '99%').setStyleAttribute('border', '1px solid grey');
      //.setStyleAttribute('borderTop', '1px solid grey');
  var mainPanel = app.createVerticalPanel().setId('mainPanel').setSize('98%', '100%')
      .setStyleAttribute('padding', '0 2%');
  scrollPanel.add(mainPanel);
  app.add(scrollPanel);
  
  ////Get the stored values and set them in UI elements
  var values = getProjectProperties_();
  
  var table1 = app.createFlexTable().setId('oAuthTable').setWidth('100%');
  mainPanel.add(table1);
  table1.setWidget(0, 0, app.createLabel('OAuth consumer key:'))
    .setWidget(0, 1, app.createTextBox().setId('consumerKey').setName('consumerKey').setWidth('200').setValue(values.consumerKey))
    .setWidget(0, 2, app.createAnchor('?', 'http://www.google.com').setTitle('Help').setStyleAttribute('fontWeight', 'bold'))
    .setWidget(1, 0, app.createLabel('OAuth consumer secret:'))
    .setWidget(1, 1, app.createTextBox().setId('consumerSecret').setName('consumerSecret').setWidth('200').setValue(values.consumerSecret))
    .setWidget(1, 2, app.createAnchor('?', 'http://www.google.com').setTitle('Help').setStyleAttribute('fontWeight', 'bold'))
    .setWidget(2, 0, app.createLabel('API Key:'))
    .setWidget(2, 1, app.createTextBox().setId('apiKey').setName('apiKey').setWidth('200').setValue(values.apiKey))
    .setWidget(2, 2, app.createAnchor('?', 'http://www.google.com').setTitle('Help').setStyleAttribute('fontWeight', 'bold'));
  
  var lineSeparator = app.createFlexTable().setId('vrSettingTable').setWidth('100%')
      .setStyleAttribute('borderTop' , '1px solid grey').setStyleAttribute('margin', '5 0');
  mainPanel.add(lineSeparator);
  
  mainPanel.add(app.createHorizontalPanel().setWidth('100%')
                .add(app.createLabel('Vacation Responder').setStyleAttribute('fontWeight', 'bold'))
                .add(app.createAnchor('?', 'http://www.google.com').setTitle('Help').setStyleAttribute('fontWeight', 'bold').setStyleAttribute('float', 'right'))
               )
  
  var table2 = app.createFlexTable().setId('vrSettingTable').setWidth('100%');
  mainPanel.add(table2);
  
  table2.setWidget(0, 0, app.createLabel('Subject:'))
    .setWidget(0, 1, app.createTextBox().setWidth('100%').setId('subject').setName('subject').setValue(values.vrSubject))
    .setWidget(1, 0, app.createLabel('Message:'))
    .setWidget(1, 1, app.createTextArea().setSize('100%', '100').setId('message').setName('message').setValue(values.vrMessage));
  
  var table3 = app.createFlexTable().setStyleAttribute('margin', '0'); 
  table2.setWidget(2, 1, table3);
  table2.setStyleAttribute(1, 0, 'verticalAlign', 'top');
  
  table3.setWidget(0, 0, app.createCheckBox('Only send a response to people in user\'s Contacts').setName('contactsOnly').setValue((values.contactsOnly == 'true')))
    .setWidget(1, 0, app.createCheckBox('Only send a response to people in Google Apps domain').setName('domainOnly').setValue((values.domainOnly == 'true')));
  
  var lineSeparator2 = app.createFlexTable().setId('vrSettingTable').setWidth('100%')
      .setStyleAttribute('borderTop' , '1px solid grey').setStyleAttribute('margin', '5 0');
  mainPanel.add(lineSeparator2);
  
  var validationInfo = app.createHTML('', true).setId('validationInfo').setStyleAttribute('color', 'red')
      .setStyleAttribute('fontSize', '11');
  mainPanel.add(validationInfo);
  
  var btnPanel = app.createHorizontalPanel().setStyleAttribute('margin', '10 auto');
  mainPanel.add(btnPanel);
  var btnActionHandler = app.createServerHandler('buttonAction_').addCallbackElement(mainPanel);
  var saveBtn = app.createButton('Save').setId('saveBtn').addClickHandler(btnActionHandler);
  btnPanel.add(saveBtn);
  var cancelBtn = app.createButton('Cancel').setId('cancelBtn').addClickHandler(btnActionHandler);
  btnPanel.add(cancelBtn);
  
  ss.show(app);
}

function buttonAction_(e){
  var source = e.parameter.source;
  var app = UiApp.getActiveApplication();
  var infoMessageLabel = app.getElementById('validationInfo').setHTML('');
  if(source == 'cancelBtn'){
    app.close();
  }
  else if (source == 'saveBtn'){
    var validationMessage = validateInputs_(e.parameter);
    if(validationMessage == 'OK'){
      saveInputs_(e.parameter);
      app.close();
    }  else{
      infoMessageLabel.setHTML(validationMessage);
    }
  }
  return app;
}

function validateInputs_(parameter){
  var validationMessage = '';
  if(parameter.consumerKey == ''){ validationMessage += 'Enetr a valid Consumer key.<br>';  }
  if(parameter.consumerSecret == ''){ validationMessage += 'Enetr a valid Consumer secret.<br>';  }
  if(parameter.apiKey == ''){ validationMessage += 'Enetr a valid API Key.<br>';  }
  if(parameter.subject == ''){ validationMessage += 'Subject is mandatory.<br>';  }
  if(parameter.message == ''){ validationMessage += 'Message is mandatory.<br>';  }
  
  if(validationMessage == ''){ return 'OK'; }
  else {return validationMessage + '<b>Refer to help (?) links for details<b>';}
}

function saveInputs_(parameter){
  //New Values
  var properties = {};
  properties.consumerKey = parameter.consumerKey;
  properties.consumerSecret = parameter.consumerSecret;
  properties.apiKey = parameter.apiKey;
  properties.vrSubject = parameter.subject;
  properties.vrMessage = parameter.message;
  properties.domainOnly = parameter.domainOnly;
  properties.contactsOnly = parameter.contactsOnly;
  //Save the new properties
  ScriptProperties.setProperties(properties);
}

function getProjectProperties_(){
  //Get older script Properties
  //If the properties are not defined, define them
  var temp = ScriptProperties.getProperties();
  if(temp.consumerKey == undefined){ ScriptProperties.setProperty('consumerKey', 'anonymous') }
  if(temp.consumerSecret == undefined){ ScriptProperties.setProperty('consumerSecret', 'anonymous') }
  if(temp.apiKey == undefined){ ScriptProperties.setProperty('apiKey', '') }
  if(temp.vrSubject == undefined){ ScriptProperties.setProperty('vrSubject', 'Out of office') }
  if(temp.vrMessage == undefined){ ScriptProperties.setProperty('vrMessage', 'I am out of office. Please expect delay in reply.') }
  if(temp.domainOnly == undefined){ ScriptProperties.setProperty('domainOnly', 'true') }
  if(temp.contactsOnly == undefined){ ScriptProperties.setProperty('contactsOnly', 'true') }
  var properties = ScriptProperties.getProperties();
  return properties;
}
//*********************************************************************************************************
