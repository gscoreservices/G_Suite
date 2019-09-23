var refSS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Backend Reference');
var completeSS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Group Data');
var emailSS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email Listing');


/*
** End Global Scope Variables
** ---------------------------------------------------------------------------------------------
** Begin UI Script Element for Displaying User's Groups
*/

function onOpen(){
  var ui = SpreadsheetApp.getUi().createMenu('Sheet Interactions')
    .addItem('My Groups', 'myGroups')
    .addItem('List Members of Groups','processGroups')
    .addToUi();
}

/*
** End UI Script Element for Displaying User's Groups
** ---------------------------------------------------------------------------------------------
** Begin Script to Process all Group Email Address Data
*/

function processGroups(){
  var addtEmails = refSS.getRange(2, 3,19,1).getValues();
  var addtEmailArray = [];
  
  for(var ea = 0; ea<addtEmails.length;cea++){
    if(addtEmails[ea] != ''){
      addtEmailArray.push(addtEmails[ea].toString())
    };
  };
  
  updateGroupData();
  SpreadsheetApp.flush();
  
  emailReport(addtEmailArray);      

}  // Closes function processGroups


/*
** End UI Script Element for Displaying User's Groups
** ---------------------------------------------------------------------------------------------
** Begin Script to Process all Group Email Address Data
*/


function updateGroupData(){
  var groupData = refSS.getRange(2, 1,18,1).getValues();
  var completeSSLastRow = completeSS.getLastRow()-1;
  var completeRange;
  var completeArray = [];

  if(completeSSLastRow>0){
    completeRange = completeSS.getRange(2, 1,completeSSLastRow,2);
    completeRange.clear();
  }     // Closes if statement to clear email listing range
  
  for(var gd=0;gd<groupData.length;gd++){
    var groupEmail = groupData[gd];
    var groupEmailStr = groupEmail.toString();
    
    if(groupEmailStr !=''){
      var group = GroupsApp.getGroupByEmail(groupEmailStr);
      var users = group.getUsers();
      
      for(var ud=0;ud<users.length;ud++ ){
        var user = users[ud];
        
        completeArray.push([groupEmailStr,user]);                      
      }     // Closes for loop for users in group    
    }     // Closes if statement for group email string not blank
  }     // Closes for loop for group data email addresses
  
  completeSS.getRange(2, 1,completeArray.length,2).setValues(completeArray);
  SpreadsheetApp.flush();
  updateEmailListing();
}     // Closes function update listing


/*
** End Script to Process all Group Email Address Data
** ---------------------------------------------------------------------------------------------
** Begin Script to Create Non-Duplicated List of Users
*/


function updateEmailListing() {
  var emailArray = [];  
  var emailSSLastRow = emailSS.getLastRow()-1;
  var emailSSLastCol = emailSS.getLastColumn();
  
  var completeSSLastRow = completeSS.getLastRow()-1;
  var emailData = completeSS.getRange(2, 2,completeSSLastRow,1).getValues();
  
  if(emailSSLastRow>0){
    emailSS.getRange(2, 1,emailSSLastRow,emailSSLastCol).clear();
  }
  
  for(var ed=0;ed<emailData.length;ed++){
    var email = emailData[ed];
    var emailStr = email.toString();
    
    if(emailArray.indexOf(emailStr)== -1){
      emailArray.push(emailStr)
      emailSS.getRange((1+emailArray.length), 1).setValue(emailStr)
    }

  
  }     // Closes for loop for email data
  SpreadsheetApp.flush();

}     // Closes function update listing

/*
** End Script to Create Non-Duplicated List of Users
** ---------------------------------------------------------------------------------------------
** Begin Script to Create CSV File and Send Email to Specified Users
*/

function emailReport(addtEmailArray){
  var timeStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy hh:mm a");  
  var emailSSLastRow = emailSS.getLastRow();
  var data = emailSS.getRange(1, 1,emailSSLastRow).getValues();
  
  
  
  for(var ea = 0; ea<addtEmailArray.length;ea++){
    var email = addtEmailArray[ea];
      data.push([email]);
  }

// Following Script Creates CSV Data Object
    
  var csvFile = undefined;  
  try {
    if (data.length > 1) {
      var csv = "";
      
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
      
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }

        if (row < data.length-1) {
          csv += data[row].join(",") + "\r\n";
        }
        
        else {
          csv += data[row];
        }
      }
      csvFile = csv;
    }
  }
  catch(err) {
    Logger.log(err);
  }  
  
// Following Script Creates CSV and Sends Email
  
  var reportName = "Email_Listing_Report_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMddyyyyhhmm")+ ".csv";
  var reportFile = Utilities.newBlob(csvFile, "application/vnd.ms-excel", reportName)
  
  var emailRecipients = refSS.getRange(2, 2).getValue();  
  var emailSubject = "Email Listing Report";
  var emailBody = "The Email Listing Report was ran at "+ timeStamp +" and is attached. Should you have any questions, please contact email@emailAddress!";
  
  var email = {
    to: emailRecipients,
    noReply: true,
    subject: emailSubject,
    htmlBody: emailBody,
    attachments: reportFile
  }
  
  try {
    MailApp.sendEmail(email);
  } catch (error) {
    Logger.log(error)
  }
  
  
}     // Closes function email listing

/*
** End Script to Create CSV File and Send Email to Specified Users
** ---------------------------------------------------------------------------------------------
** Begin Script to Display User at Clients Groups
*/

function myGroups(){
  var groups = GroupsApp.getGroups();
  var groupMessage = 'You are a direct member of ' + groups.length + ' groups: \n\n';
  
  for (var i = 0; i < groups.length; i++) {
    var group = groups[i];
    var groupBody = group.getEmail()
      groupMessage = groupMessage + groupBody + ' \n';
  }
      
  var ui = SpreadsheetApp.getUi();
    ui.alert(groupMessage)

   
}



