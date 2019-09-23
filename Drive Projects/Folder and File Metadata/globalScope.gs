var activeSheet = SpreadsheetApp.getActiveSpreadsheet();

var querySheetName = "Custom Queries";

var systemLogSSName = "System Log Sheet";
var systemLogsSS = activeSheet.getSheetByName(systemLogSSName);

var tempSheetName = "Template Sheet";
var tempSheet = activeSheet.getSheetByName(tempSheetName);

var referenceSheet = "Reference Sheet";

var testSheetName = "test";
var testSS = activeSheet.getSheetByName(testSheetName);

var multLocationSSName = "Items in More than One Location";


var sProperties = PropertiesService.getScriptProperties();
var dProperties = PropertiesService.getDocumentProperties();
var uProperties = PropertiesService.getUserProperties();

var timeTriggerKey = "timeTrigger";
var recursiveKey = "recursiveKey";
var userDriveIdKey = "userDriveIdKey";
var updateSheetKey = "updateSheet";

var devEmailAccount = "dnr_dwrappdev@state.co.us";


var a = ScriptApp.getScriptId();
var checkOwner = '';

/*
** End Global Scope Variables
** ------------------------------------------------------------------------------------------------------------------------
** Begin Event and Helper Functions
*/

// Function Displays the Menu Items in the User Interface/Google Sheet
function onOpen() {
    var ui = SpreadsheetApp.getUi().createMenu('Google Drive App')
      
      .addItem('Harvest Large Folder Metadata','promptLargeUrl')
      .addItem('Harvest Small Folder Metadata','promptSmallUrl')
      .addSeparator()
      .addItem('Reset Application', 'resetApp') 
      .addToUi(); 
  }

function getSheetNames(){
    var sheetArray = [];

    activeSheet.getSheets().forEach(function(sheet){
        sheetArray.push(sheet.getName());
    })

    return sheetArray
}

function doGet(e){
  var page = HtmlService.createTemplateFromFile('hierachicalTreePage')
                        .evaluate()
                        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
                        .addMetaTag("viewport", "initial-scale=1.0, maximum-scale=1.0")
                        .setTitle("Google Drive Graphical View");
  
  return page
  
}

function loadStyleSheet() {
  return HtmlService.createHtmlOutputFromFile('styleSheet').getContent();
}     // Closes function loadStyleSheet()
