// ======================================================================
// Instructions 
// ======================================================================
//
// 1. I would recommend making a copy of this script so that you can edit your settings, and maintain ownership of the code.
//    This is done using the menu above: File->Make A Copy
//
// 2. Edit these user configurable settings - or leave them if you think they're ok.

var MARK_UNREAD = true;                           // When true, marks thread as unread when moving it back to the inbox
var MARK_MESSAGE_UNREAD_ONLY = true;             // When true, only marks the most recent message as unread
var ADD_UNSNOOZED_LABEL = true;                   // Adds the unsnoozed label when moving back to the inbox
var ADD_SNOOZE_ERROR_LABEL = true;                // If an invalid snooze label is used, labels thread as invalid when returning to inbox.
var SNOOZE_LABEL = "* Snooze";                    // Prefixing it with * puts it at the top of the list.
var UNSNOOZE_LABEL = "* Unsnoozed";               // Prefixing it with * puts it at the top of the list.  
var SNOOZE_ERROR_LABEL = "Snooze Error!";

var TIME_A_DAY_STARTS = "0800";

// 3. Select Run->Setup from the menu above.
//
// 4. If prompted to authorise this script to access Gmail and Calendar (which is used to determine your preferred timezone) 
//    then authorise the script until you see the text "Now you can run the script." - THEN REPEAT STEP 3 to complete the setup.
//
// 5. You're now good to go - when you go to Gmail, you should now have Snooze labels that you can use to snooze in various ways.
//    
//
// ------------------------------------------------------------------------------------------------------------------------
//
//    For more detailed instructions, and examples, visit: http://www.littlebluemonkey.com/gmail-snooze-without-mailboxapp/
//
// ------------------------------------------------------------------------------------------------------------------------
//
//    Release Notes: 03/07/2013 1.2 - Fixed buy where certain mmddyy dates were resulting in 'Snooze Error!' labels being applied
//
// ------------------------------------------------------------------------------------------------------------------------










// ======================================================================
// Code - Don't proceed any further unless you're a trained professional.
// ======================================================================


// --------------------------------------------------------------------------
// Stuff that you really shouldn't be touching unless you want to break stuff.
// --------------------------------------------------------------------------

var KEY_TOMORROW_DATE = "tomorrowDate";
var KEY_CURRENT_DAY = "currentDay";

var MONDAY_LABEL = "1 - Monday";
var TUESDAY_LABEL = "2 - Tuesday";
var WEDNESDAY_LABEL = "3 - Wednesday";
var THURSDAY_LABEL = "4 - Thursday";
var FRIDAY_LABEL = "5 - Friday";
var SATURDAY_LABEL = "6 - Saturday";
var SUNDAY_LABEL = "7 - Sunday";

var snoozeErrorLabel = null;
var unsnoozeLabel = null;


// --------------------------------------------------------------------------
// Setup function to add labels, initialise properties and set script trigger.
// --------------------------------------------------------------------------

function setup() {
   
  GmailApp.createLabel(SNOOZE_LABEL);
  GmailApp.createLabel(SNOOZE_LABEL + "/0 - Tomorrow");
  GmailApp.createLabel(SNOOZE_LABEL + "/" + MONDAY_LABEL);
  GmailApp.createLabel(SNOOZE_LABEL + "/" + TUESDAY_LABEL);
  GmailApp.createLabel(SNOOZE_LABEL + "/" + WEDNESDAY_LABEL);
  GmailApp.createLabel(SNOOZE_LABEL + "/" + THURSDAY_LABEL);
  GmailApp.createLabel(SNOOZE_LABEL + "/" + FRIDAY_LABEL);
  GmailApp.createLabel(SNOOZE_LABEL + "/" + SATURDAY_LABEL);
  GmailApp.createLabel(SNOOZE_LABEL + "/" + SUNDAY_LABEL);
  GmailApp.createLabel(SNOOZE_LABEL + "/8 - time_hhmm");
  GmailApp.createLabel(SNOOZE_LABEL + "/8 - time_hhmm/1400");
  GmailApp.createLabel(SNOOZE_LABEL + "/8 - time_hhmm/1800");
  GmailApp.createLabel(SNOOZE_LABEL + "/8 - time_hhmm/2200");
  GmailApp.createLabel(SNOOZE_LABEL + "/9 - date_ddmmyy");
  GmailApp.createLabel(SNOOZE_LABEL + "/9 - date_mmddyy");
  GmailApp.createLabel(SNOOZE_LABEL + "/9 - date_ddmmyy/280913");
  if (ADD_UNSNOOZED_LABEL) {
    GmailApp.createLabel(UNSNOOZE_LABEL);
  }
  
  initialiseProperty(KEY_TOMORROW_DATE, getCurrentIsoDate());
  initialiseProperty(KEY_CURRENT_DAY, getCurrentDay());
  
  ScriptApp.newTrigger("moveSnoozes").timeBased().everyMinutes(15).create();
}




// ------------------------------------------------------------------------
// The main script that does all the good stuff - triggered every 1 minute:
// ------------------------------------------------------------------------

function moveSnoozes() {
  var oldLabel, newLabel, page;
  
  var today = getCurrentDay();
  var todayIso = getCurrentIsoDate();
  var now = getCurrentIsoDateAndTime();
  
  var snoozeLength = SNOOZE_LABEL.length;
  var labelName = "";
  
  var labels = GmailApp.getUserLabels();
  
  for (var i = 0; i < labels.length; i++) {
    
    labelName = labels[i].getName();
    
    if (labelName.length >= snoozeLength) {
      
      if (labelName.substr(0,snoozeLength) == SNOOZE_LABEL) {
        
        if (isDayLabel(labelName)) {
          if (labelName.substr(snoozeLength + 5) === today && todayCanBeProcessed(today) && timeCanBeProcessed(TIME_A_DAY_STARTS,now)) {
            processLabel(labels[i]);
            UserProperties.setProperty(KEY_CURRENT_DAY, today);
          }  
        }
        
        else if(labelName.substr(snoozeLength) == "/0 - Tomorrow") {
          if (tomorrowCanBeProcessed(todayIso) && timeCanBeProcessed(TIME_A_DAY_STARTS,now)) {
            processLabel(labels[i]);
            UserProperties.setProperty(KEY_TOMORROW_DATE, todayIso);
          }  
        }  
     
        else if(isValidTimeLabel(labelName,snoozeLength)) {
          var time = getTimeFromLabel(labelName,snoozeLength);
      
          if (timeCanBeProcessed(time,now)) {
            processLabel(labels[i]);
            UserProperties.setProperty("Time:"+time,now.substr(0,10));
          }
        }
        
        else if(isValidDDMMYYLabel(labelName,snoozeLength)) {
          if (dateCanBeProcessed(getDateFromDDMMYYLabel(labelName,snoozeLength),todayIso) && timeCanBeProcessed(TIME_A_DAY_STARTS,now)) {
            processLabel(labels[i]);
            GmailApp.deleteLabel(labels[i]);
          }
        }
        
        else if(isValidMMDDYYLabel(labelName,snoozeLength) && timeCanBeProcessed(TIME_A_DAY_STARTS,now)) {
          if (dateCanBeProcessed(getDateFromMMDDYYLabel(labelName,snoozeLength),todayIso)) {
            processLabel(labels[i]);
            GmailApp.deleteLabel(labels[i]);
          }
        }
      
        // If there are any "Snooze" type labels that aren't what we're expecting
        // then flag them in error and move back to the inbox:
        
        else {
          processInvalidLabel(labels[i]);
        }
      }  
    }  
  }
  
  return;
}

function processLabel(label) { 
  
  var page = null;

  
  // Get threads in "pages" of 100 at a time
  while(!page || page.length == 100) {
    page = label.getThreads(0, 100);
    if (page.length > 0) {
      for (var i =0;i<page.length;i++) {
        if (MARK_UNREAD) {
          if (MARK_MESSAGE_UNREAD_ONLY) {
            var msgs = page[i].getMessages();
            msgs[msgs.length-1].markUnread();
          }
          else {
            page[i].markUnread();
          }  
        }  
        page[i].moveToInbox();
        page[i].removeLabel(label);
        if (ADD_UNSNOOZED_LABEL) {
          var unsnooze = getUnsnoozeLabel();
          if (unsnooze != null) {
            page[i].addLabel(unsnooze);
          }  
        }
      }  
    }
  }
}

function processInvalidLabel(label) { 
 
  var page = null;
  
    // Get threads in "pages" of 100 at a time
  while(!page || page.length == 100) {
    page = label.getThreads(0, 100);
    if (page.length > 0) {
      for (var i =0;i<page.length;i++) {
        page[i].markUnread();
        page[i].moveToInbox();
        page[i].removeLabel(label);
        if (ADD_SNOOZE_ERROR_LABEL) {
          var snoozeError = getSnoozeErrorLabel();
          if (snoozeError != null) {
            page[i].addLabel(snoozeError);
          }
        }
      }  
    }
  }
  
  return;
 
}

function todayCanBeProcessed(newDay) {
  
  var oldDay = UserProperties.getProperty(KEY_CURRENT_DAY);
  
  if (oldDay !== newDay) {
    // Check that it's after a configured time.
    return true;
  }  
  else {
    return false;   
  }  
  
}

function tomorrowCanBeProcessed(newIsoDate) {
  
  var oldDate = UserProperties.getProperty(KEY_TOMORROW_DATE);
  
  if (oldDate !== newIsoDate) {
    return true;
  }  
  else {
    return false;   
  }  
  
}

function timeCanBeProcessed(time,now) {
  
  if (time !== "") {
    var triggerDateAndTime = now.substr(0,10) + "-" + time;
    var processedDate = UserProperties.getProperty("Time:"+time);
    if (processedDate === null) { 
      if (now > triggerDateAndTime) {
        processedDate = now.substr(0,10);
      }  
      else {
        processedDate = "0001-01-01";  
      }  
      UserProperties.setProperty("Time:"+time, processedDate);  
    }  
    
    if (now >= triggerDateAndTime) {
        
      if (processedDate === null || processedDate === "" || processedDate < getCurrentIsoDate()) {
        return true;
      }  
      
    }  
  }  
  
  return false;
  
}  

function dateCanBeProcessed(theDate,today) {
  if (today >= theDate) {
    return true;
  }
  return false;
}


function getTimeFromLabel(labelName,snoozeLength) {
  
  var labelLength = labelName.length;
  var timeLength = labelLength - (snoozeLength+15);
  var time = "";
  
  if ( labelLength == snoozeLength + 18 || labelLength == snoozeLength + 19) { 
    if (labelName.substr(snoozeLength,15) == "/8 - time_hhmm/") {
      time = labelName.substr(snoozeLength + 15,timeLength);
      if (timeLength == 3) {
        time = "0" + time;
      }
      var isValid = true;
      for (var i = 0; i<4; i++) {
        if (isNaN(time.substr(i,1))) {
          isValid = false;
        }    
      }  
      if (isValid) {
        var hours = parseInt(time.substr(0,2),10);
        var minutes = parseInt(time.substr(2,2),10);  
        if (hours < 0 || hours > 23) {
          isValid = false;
        }  
        else if (minutes < 0 || minutes > 59) {
          isValid = false;
        }  
      }
      if (!isValid) {
        time = "";
      }    
    }
  }  
  
  return time;
} 

function isValidTimeLabel(labelName,snoozeLength) {
  var time = getTimeFromLabel(labelName,snoozeLength);
  if (time !== "") {
    return true;
  }  
  return false;
}  


function getDateFromDDMMYYLabel(labelName,snoozeLength) {
  
  var labelLength = labelName.length;
  var ddmmyyLength = labelLength - (snoozeLength+17);
  var ddmmyy = "";
  
  if ( labelLength == snoozeLength + 17 + 6) { 
    if (labelName.substr(snoozeLength,17) == "/9 - date_ddmmyy/") {
      ddmmyy = labelName.substr(snoozeLength + 17,6);
      var isValid = true;
      for (var i = 0; i<6; i++) {
        if (isNaN(ddmmyy.substr(i,1))) {
          isValid = false;
        }    
      }  
      if (isValid) {
        var dd = parseInt(ddmmyy.substr(0,2),10);
        var mm = parseInt(ddmmyy.substr(2,2),10);  
        var yyyy = 2000 + parseInt(ddmmyy.substr(4,2),10);  
        isValid = isValidDateValues(dd,mm,yyyy);
      }  
    }  
  }      
  if (!isValid) {
    ddmmyy = "";
  }
  else {
    ddmmyy = '20' + ddmmyy.substr(4,2) + "-" + ddmmyy.substr(2,2) + '-' + ddmmyy.substr(0,2);
  }
    
  return ddmmyy;
} 

function getDateFromMMDDYYLabel(labelName,snoozeLength) {
  
  var labelLength = labelName.length;
  var mmddyyLength = labelLength - (snoozeLength+17);
  var mmddyy = "";
  
  if ( labelLength == snoozeLength + 17 + 6) { 
    if (labelName.substr(snoozeLength,17) == "/9 - date_mmddyy/") {
      mmddyy = labelName.substr(snoozeLength + 17,6);
      var isValid = true;
      for (var i = 0; i<6; i++) {
        if (isNaN(mmddyy.substr(i,1))) {
          isValid = false;
        }    
      }  
      if (isValid) {
        var dd = parseInt(mmddyy.substr(2,2),10);
        var mm = parseInt(mmddyy.substr(0,2),10);  
        var yyyy = 2000 + parseInt(mmddyy.substr(4,2),10);  
        isValid = isValidDateValues(dd,mm,yyyy);
      }  
    }  
  }      
  if (!isValid) {
    mmddyy = "";
  }
  else {
    mmddyy = '20' + mmddyy.substr(4,2) + "-" + mmddyy.substr(0,2) + '-' + mmddyy.substr(2,2);
  }
    
  return mmddyy;
} 


function isValidDateValues(dd,mm,yyyy) {
  
  if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31 && (yyyy >= 2000 && yyyy <= 9999)) {
    if (((mm == 4 || mm == 6 || mm == 9 || mm == 11) && dd <= 30 ) 
              || (mm == 2 && dd <= 28 || dd <= 29 && isLeapYear(yyyy))
              || (mm != 4 && mm != 6 && mm != 9 && mm != 11 && mm != 2)) {
      return true;
    }
  }  
  
}

function isLeapYear(years) {
  
  if (years % 400 === 0  ) {return true;}
  else if (years % 100 === 0  ) {return false;}
  else if (years % 4 === 0  ) {return true;}
  else {return false;}
  
}

function isValidDDMMYYLabel(labelName,snoozeLength) {
  var ddmmyy = getDateFromDDMMYYLabel(labelName,snoozeLength);
  if (ddmmyy !== "") {
    return true;
  }
  return false;
}  


function isValidMMDDYYLabel(labelName,snoozeLength) {
  var mmddyy = getDateFromMMDDYYLabel(labelName,snoozeLength);
  if (mmddyy !== "") {
    return true;
  }
  return false;
}  


function getUsersTimeZone() {
  return CalendarApp.getTimeZone();
}  



function getDayOfWeek(aDate) {
 
   return Utilities.formatDate(aDate, getUsersTimeZone(), "EEEE");
  
} 

function getCurrentDay() {
  return getDayOfWeek(new Date());
}  


function getCurrentIsoDateAndTime() {
   return Utilities.formatDate(new Date(), getUsersTimeZone(), "yyyy-MM-dd-HHmm");
} 

function getCurrentIsoDate() {
   return Utilities.formatDate(new Date(), getUsersTimeZone(), "yyyy-MM-dd");
} 

function isDayLabel(labelName) {
  
  
  if (labelName === SNOOZE_LABEL + "/" + MONDAY_LABEL
    || labelName === SNOOZE_LABEL + "/" + TUESDAY_LABEL 
    || labelName === SNOOZE_LABEL + "/" + WEDNESDAY_LABEL 
    || labelName === SNOOZE_LABEL + "/" + THURSDAY_LABEL 
    || labelName === SNOOZE_LABEL + "/" + FRIDAY_LABEL 
    || labelName === SNOOZE_LABEL + "/" + SATURDAY_LABEL 
    || labelName === SNOOZE_LABEL + "/" + SUNDAY_LABEL) {
    return true;
  }  
  
  return false;
  
}  

function getUnsnoozeLabel() {
  
  if (unsnoozeLabel === null) {
    Logger.log("Load unsnooze label from global");
    unsnoozeLabel = GmailApp.getUserLabelByName(UNSNOOZE_LABEL);
  }  

  if (unsnoozeLabel === null) {
    unsnoozeLabel = GmailApp.createLabel(UNSNOOZE_LABEL);
  }  
  else {
    Logger.log("unsnooze found");
  }  

  return unsnoozeLabel;
  
}  

function getSnoozeErrorLabel() {
  
  if (snoozeErrorLabel === null) {
    Logger.log("Load snoozeError label from global");
    snoozeErrorLabel = GmailApp.getUserLabelByName(SNOOZE_ERROR_LABEL);
  }  

  if (snoozeErrorLabel === null) {
    snoozeErrorLabel = GmailApp.createLabel(SNOOZE_ERROR_LABEL);
  }  
  else {
    Logger.log("SnoozeError found");
  }  

  return snoozeErrorLabel;
  
}  

function initialiseProperty(key,initialValue) {
  
  var value = UserProperties.getProperty(key);
  if (value === null || value === "") {
    UserProperties.setProperty(key, initialValue);
  }  
  
}
