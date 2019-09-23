/*
** Begin Functions that Control the Project Triggers
*/

function createTriggers(){
  var timeTriggerId = sProperties.getProperty(timeTriggerKey);
  
  if(timeTriggerId == null) {
    var triggerId = ScriptApp.newTrigger("processRootFolder")
                              .timeBased()
                              .everyMinutes(30)
                              .create()
                              .getUniqueId();
    
    sProperties.setProperty(timeTriggerKey, triggerId);
  }
}

function deleteTriggers(){
  var timeTriggerId = sProperties.getProperty(timeTriggerKey);
  var allTriggers = ScriptApp.getProjectTriggers();
  
  for(var i=0; i<allTriggers.length; i++){
    if(allTriggers[i].getUniqueId() == timeTriggerId) {
      ScriptApp.deleteTrigger(allTriggers[i])
    }
  }
  
  sProperties.deleteProperty(timeTriggerKey);
  dProperties.deleteProperty(userDriveIdKey);
  dProperties.deleteProperty(updateSheetKey);

  systemLogsSS.getRange(2,1,systemLogsSS.getLastRow(), systemLogsSS.getLastColumn()).clear();

}

/*
** End Functions that Control the Project Triggers
** --------------------------------------------------------------------------------------------------------------
** Begin User Interface Menu Functions
*/

// Function Prompts the User for a Google Drive Folder and Handles the User's Response
function promptLargeUrl(){
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt("Process Google Drive Folder",'Please enter the url of the Google Drive folder that you would like to index:', ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  var rootFolderUrl = result.getResponseText();
  var rootFolderArray = rootFolderUrl.split("/");

  if((button == ui.Button.OK) && (rootFolderUrl == "")){
    SpreadsheetApp.getUi().alert('A Google Drive Folder URL is required for this system to work properly!');

  } else if((button == ui.Button.OK) && (rootFolderUrl != "")){
    if(rootFolderArray[4] == "folders"){
      var folderId = rootFolderArray[5];
      dProperties.setProperty(userDriveIdKey,folderId);
      checkAndClearSheet(folderId)
      
      createTriggers();
      // processRootFolder();

      SpreadsheetApp.getUi().alert("Your request has been scheduled. You will recieve an email when the Google Drive App has finished processing your request.")

    } else if(rootFolderArray[6] == "folders"){
      var folderId = rootFolderArray[7];
      dProperties.setProperty(userDriveIdKey,folderId);
      checkAndClearSheet(folderId)
      
      createTriggers();
      // processRootFolder();

      SpreadsheetApp.getUi().alert("Your request has been scheduled. You will recieve an email when the Google Drive App has finished processing your request.")

    } else {
      SpreadsheetApp.getUi().alert('Please provide a valid Google Drive Folder Url!');
    }
  }
}


// Function Prompts the User for a Google Drive Folder and Handles the User's Response
function promptSmallUrl(){
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt("Process Google Drive Folder",'Please enter the url of the Google Drive folder that you would like to index in the text field below. \n\n NOTE: If your Google Drive Folder has more than 1000 items, this function may not successfully complete within the allowable execution time. If your folder has more than 1000 folders and files please use the "Harvest Large Folder Metadata" option above.', ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  var rootFolderUrl = result.getResponseText();
  var rootFolderArray = rootFolderUrl.split("/");

  if((button == ui.Button.OK) && (rootFolderUrl == "")){
    SpreadsheetApp.getUi().alert('A Google Drive Folder URL is required for this system to work properly!');

  } else if((button == ui.Button.OK) && (rootFolderUrl != "")){
    if(rootFolderArray[4] == "folders"){
      var folderId = rootFolderArray[5];
      dProperties.setProperty(userDriveIdKey,folderId);
      checkAndClearSheet(folderId)
      
      processRootFolder(rootFolderArray[5]);
      

    } else if(rootFolderArray[6] == "folders"){
      var folderId = rootFolderArray[7];
      dProperties.setProperty(userDriveIdKey,folderId);
      checkAndClearSheet(folderId)
      
      
      processRootFolder(rootFolderArray[7]);
      

    } else {
      SpreadsheetApp.getUi().alert('Please provide a valid Google Drive Folder Url!');
    }
  }
}

function checkAndClearSheet(folderId){
  var folderName = DriveApp.getFolderById(folderId).getName();
  var allSheetArray = getSheetNames();
  
  if(allSheetArray.indexOf(folderName) == -1){
  
  } else {
    var sheet = activeSheet.getSheetByName(folderName);
    var sheetLastRow = sheet.getLastRow()-1
    var sheetLastCol = sheet.getLastColumn();
    
    sheet.getRange(2,1,sheetLastRow,sheetLastCol).clear();
  }
}


function resetApp(){
  dProperties.deleteProperty(recursiveKey);
  dProperties.deleteProperty(userDriveIdKey);
  dProperties.deleteProperty(updateSheetKey);

  deleteTriggers();

  systemLogsSS.getRange(2,1,systemLogsSS.getLastRow(), systemLogsSS.getLastColumn()).clear();
}


/*
** End Global Scope Variables and Helper Functions
** -------------------------------------------------------------------------------------------------------------------------------------
**
*/


function processRootFolder(folderId) {
  var rootFolderId = "";

  if(typeof folderId == "string"){
    rootFolderId = folderId;
  } else {
    rootFolderId = dProperties.getProperty(userDriveIdKey);
  }
  
  var rootFolder = DriveApp.getFolderById(rootFolderId);
  var rootFolderName = rootFolder.getName();

  var allSheetArray = getSheetNames();

  if(allSheetArray.indexOf(rootFolderName) == -1){
    activeSheet.insertSheet(0,{template: tempSheet}).setName(rootFolderName);
  } 
  
  dProperties.setProperty(updateSheetKey,rootFolderName);

  var MAX_RUNNING_TIME_MS = 27 * 60 * 1000;  // (Minutes * Seconds * Milliseconds)

  var startTime = (new Date()).getTime();

  // [{folderName: String, fileIteratorContinuationToken: String?, folderIteratorContinuationToken: String}]
  var recursiveIterator = JSON.parse(dProperties.getProperty(recursiveKey));
  if (recursiveIterator !== null) {
    // verify that it's actually for the same folder
    if (rootFolder.getName() !== recursiveIterator[0].folderName) {
      recursiveIterator = null;
      systemLogsSS.appendRow([new Date(),"WARNING", "Looks like this is a new folder. Clearing out the old iterator."])
      
    } else {
      systemLogsSS.appendRow([new Date(),"INFO", "Resuming session."])
    }
  }
  
  if (recursiveIterator === null) {
    systemLogsSS.appendRow([new Date(),"INFO", "Starting new session."])    
    recursiveIterator = [];
    recursiveIterator.push(makeIterationFromFolder(rootFolder));
  }

  while (recursiveIterator.length > 0) {
    recursiveIterator = nextIteration(recursiveIterator, startTime);

    var currTime = (new Date()).getTime();
    var elapsedTimeInMS = currTime - startTime;
    var timeLimitExceeded = elapsedTimeInMS >= MAX_RUNNING_TIME_MS;
    if (timeLimitExceeded) {
      dProperties.setProperty(recursiveKey, JSON.stringify(recursiveIterator));

      systemLogsSS.appendRow([new Date(),"INFO", "Stopping loop after "+ elapsedTimeInMS +" milliseconds. Please continue running."]);
      
      return;
    }
  }
  dProperties.deleteProperty(recursiveKey);
  systemLogsSS.appendRow([new Date(),"INFO", "Done running."]);

  var emailSubject = "Google Drive Indexing Complete for " + rootFolderName;
  var emailBody = 'Hi, <br/><br/> We have finished indexing your Google Drive Folder (' + rootFolderName + '), please click the link <a href="https://docs.google.com/spreadsheets/d/1leAIN7xBI-VpKVTLkRwKcDjkdw5rMS6Qzn561Nhyhm8/edit">here</a>, to view the data. If you experience any technical difficulties, please contact ' + devEmailAccount + '.<br/><br/> Thanks!';

  var emailObj = {
    name: "Google Drive App",
    to: Session.getEffectiveUser().getEmail(),
    subject: emailSubject,
    htmlBody: emailBody,
    noReply: true
  }

  MailApp.sendEmail(emailObj);
  setOccuranceFormula(rootFolderName);
  
  deleteTriggers();

}

function setOccuranceFormula(sheetName){
  var ss = activeSheet.getSheetByName(sheetName);
  var ssLastRow = ss.getLastRow()-1;
  
  var occuranceFormula = "=COUNTIF($D$2:$D,D2)";
    ss.getRange(2,16,ssLastRow,1).setFormula(occuranceFormula);
}


function nextIteration(recursiveIterator) {
  var currentIteration = recursiveIterator[recursiveIterator.length-1];

  // Script processes the next files if any
  if (currentIteration.fileIteratorContinuationToken !== null) {
    var fileIterator = DriveApp.continueFileIterator(currentIteration.fileIteratorContinuationToken);    
    if (fileIterator.hasNext()) {
      var path = recursiveIterator.map(function(iteration){return iteration.folderName;}).join("/");
      
      processDriveObj(fileIterator.next(),path)
      
      currentIteration.fileIteratorContinuationToken = fileIterator.getContinuationToken();
      recursiveIterator[recursiveIterator.length-1] = currentIteration;
      return recursiveIterator;
      
    } else {
      // No More Files to Process
      currentIteration.fileIteratorContinuationToken = null;
      recursiveIterator[recursiveIterator.length-1] = currentIteration;
      return recursiveIterator;
    }
  }

  // Script processes the next folders if any
  if (currentIteration.folderIteratorContinuationToken !== null) {
    var folderIterator = DriveApp.continueFolderIterator(currentIteration.folderIteratorContinuationToken);
    if (folderIterator.hasNext()) {
      var folder = folderIterator.next();
      var path = recursiveIterator.map(function(iteration){return iteration.folderName}).join("/");
      
      processDriveObj(folder,path)
      
      recursiveIterator[recursiveIterator.length-1].folderIteratorContinuationToken = folderIterator.getContinuationToken();
      recursiveIterator.push(makeIterationFromFolder(folder));
      return recursiveIterator;

    } else {
      // No More Folders to Process
      recursiveIterator.pop();
      return recursiveIterator;
    }
  }
  
  cLogSS.appendRow([new Date(),"ERROR", "Should Never Have Gotten Here"]);
}

function makeIterationFromFolder(folder) {
  return {
    folderName: folder.getName(), 
    fileIteratorContinuationToken: folder.getFiles().getContinuationToken(),
    folderIteratorContinuationToken: folder.getFolders().getContinuationToken()
  };
}

/*
** 
** -------------------------------------------------------------------------------------------------------------------------------------
** Begin Functions to Update Complete Listing Sheet with Drive Object Attributes
*/

function processDriveObj(driveObj,path){
  var id = driveObj.getId();
  var url = driveObj.getUrl();
  var name = driveObj.getName();
  
  var lastUpdated = driveObj.getLastUpdated();
  var numFolders = path.split("/").length;
  
  var owner = "";
  try {
    owner = driveObj.getOwner().getEmail();
  } catch (error){
    owner = error;
  }

  var access = "";
  try {
    access = driveObj.getSharingAccess();
  } catch (error) {
    access = error;
  }

  var permission = "";
  try {
    permission = driveObj.getSharingPermission();
  } catch (error) {
    permission = error; 
  }
  
  var type = "";
  var fileSize = 0
  if(url.indexOf("/drive/folders/") != -1){
    type = "Folder"
  } else {
    type = driveObj.getMimeType();
    fileSize = (driveObj.getSize())/1000000;
  }
  
  var editorEmailArray = [];
  var editors = driveObj.getEditors();
  for(var i=0; i<editors.length; i++){
    editorEmailArray.push(editors[i].getEmail());
  }

  var viewerEmailArray = [];
  var viewers = driveObj.getViewers();
  for(var i=0; i<viewers.length; i++){
    viewerEmailArray.push(viewers[i].getEmail());
  }

  var sheetName = dProperties.getProperty(updateSheetKey)
  var updateSheet = activeSheet.getSheetByName(sheetName);

  var parents = driveObj.getParents();
  while(parents.hasNext()){
    var parent = parents.next();
    var parentName = parent.getName();
    var parentId = "";
    
    if(parentName != sheetName){
      parentId = parent.getId();
    }

    updateSheet.appendRow([type, parentId, parentName, id, name, url, owner, editorEmailArray.toString(), viewerEmailArray.toString(), access, permission, lastUpdated, path, numFolders,fileSize]);

  }
}
