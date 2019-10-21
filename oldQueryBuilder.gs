
  function main() {
  if(messageType == "") {
    var mData = {
      nickName: messageName,
      messageContent: messageData
    };
  }
  else {
    var mData = {
      nickName: messageName,
      messageContent:messageData
    };
  }
  
  else if((e.range.getColumn() == 8 || e.range.getColumn() == 9) && (e.range.getRow() == 3 || e.range.getRow() == 4)) {
    var searchType = 2;
    Logger.log(searchType)
    var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(searchToolSheet);
    var mySearchTerms = ss.getRange("H4").getValues();
    var checks = ss.getRange(3, 8, 1, 2).getValues();
    
    // set P checkboxes blank
    ss.getRange(3, 3, 1, 4).setValues([[false,false,false,false]]);
    ss.getRange(4, 3, 1, 1).setValue("");
    ss.getRange("H7:H").clear();
    ss.getRange("I7:I").clear();
    
    //Reads C3,D3,E3,F3 then decides which query to build. Separates words in F by "," and turns "," into query words
    var checkResult =  parseInt(checkTool(checks));
    var cols = "Col1,Col2,Col3,Col4,Col6,Col7,Col8,Col9"; // took out ,Col5
    
    //Build custom where part of query
    var myTerms = buildSearchTerms(mySearchTerms,searchType);
    var myDataSourceString = buildDataSources2(checks);

    var formula = "if(H4<>\"\"\,iferror(arrayformula(query({"+myDataSourceString+"},\"\select "+cols+" where "+myTerms+"\"\)), \"\No Results!\"\), \"\"\)";
    ss.getRange("A7").setFormula(formula);

    var header= ["SENDER",  "JOBID",  "DATE", "LEAD STATUS",  "CLIENT NAME",  "DEST. E-MAIL/PHONE", "MESSAGE NICKNAME", "NOTES", "LINK"];
    var headerRange = ss.getRange("A6:I6").setValues([header]);
    SpreadsheetApp.flush();
    jobLinkHandlerMessageLogIncText();
    return;
  }
}
// Create copy of data souces into Search Storagen Hidden Sheet. This reformats mixed data from Message Log into text for query matching
function copyMLtoHidden() {
  var hiddenSS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Search Storage Hidden");
  hiddenSS.getRange(2,1,hiddenSS.getLastRow() ,26).clear();
  var messageLogSS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(messageLog);
  var mlData = messageLogSS.getRange(4, 1, messageLogSS.getLastRow()-4, messageLogSS.getMaxColumns()).copyValuesToRange(hiddenSS, 1, messageLogSS.getMaxColumns(), 4, messageLogSS.getLastRow()-4);
  SpreadsheetApp.flush();
  var numRows = hiddenSS.getRange("C4:C").getValues().filter(String).length;

  // Convert column G (phone/email) to string
  var contactInfo = hiddenSS.getRange(4, 7, numRows, 1).getValues();

  for(var i=0; i<contactInfo.length; i++) {
    contactInfo[i][0] = ""+contactInfo[i][0];
  }
  Logger.log(contactInfo)
  hiddenSS.getRange(4, 7, contactInfo.length, 1).setValues(contactInfo)
}

function copyITtoHidden() {
  var hiddenSS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Search Storage Hidden");
  var textSS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(incomingTexts);
  var txtData = textSS.getRange(15, 1, textSS.getLastRow()-15, textSS.getMaxColumns()).copyValuesToRange(hiddenSS, 11, 20, 15, textSS.getLastRow()-14);
  SpreadsheetApp.flush();
  var numRows = hiddenSS.getRange("M4:M").getValues().filter(String).length;

  // Convert column G (phone/email) to string
  var contactInfo = hiddenSS.getRange(15, 17, numRows, 1).getValues();

  for(var i=0; i<contactInfo.length; i++) {
    contactInfo[i][0] = ""+contactInfo[i][0];
  }
  Logger.log(contactInfo)
  hiddenSS.getRange(15, 17, contactInfo.length, 1).setValues(contactInfo)
}


// Message Log and Text Message Links: Search from the bottom of the data to top
// Search by matching contact
 Creates Links to a Cell based on the query result
function jobLinkHandlerMessageLogIncText() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var searchToolSS = ss.getSheetByName(searchToolSheet);
  var messageLogSS = ss.getSheetByName(messageLog);
  var incomingTextSS = ss.getSheetByName(incomingTexts);
  var messageLogData = messageLogSS.getRange(4, 1, messageLogSS.getLastRow()-3, messageLogSS.getLastColumn()).getValues();
  var incomingTextData = incomingTextSS.getRange(15, 1, incomingTextSS.getLastRow()-14, incomingTextSS.getLastColumn()).getValues();
  var cellLinks = [];

  //Get gid for each sheet. This can change if the sheet is moved so must do this
  var MLGID = ss.getSheetByName(messageLog).getSheetId();
  var ITGID = ss.getSheetByName(incomingTexts).getSheetId();
  var queryData = searchToolSS.getRange(7, 1, searchToolSS.getLastRow()-6, 9).getValues();

  for(var k=0; k<queryData.length; k++) {
    if(queryData[k][5].toString().match(/^\(\d\d\d\) \d\d\d-\d\d\d\d$/) !== null) {
      queryData[k][5] = ""+queryData[k][5].replace(/[() -]/g,"");
    }
  }

  //Loop through sheet depending on nickname
  for(var i=0; i<queryData.length; i++) {
    // Found no Message Nickname, therefore it is from Incoming Text  
    if(queryData[i][6] == "") { 
      Logger.log("INCOMING TEXT")
      var j = incomingTextData.length;
      while(j--) {
          var row = 15+j;
          Logger.log("FOUND A INCOMING TEXT ROW")
          cellLinks.push(['https://docs.google.com/spreadsheets/d/'+spreadsheetId+'/edit#gid='+ITGID+'&range='+row+':'+row]);
          j=0;
        }
        else if(j==0) {
        cellLinks.push(['']);
        }
      }
    }
    // Found Message Nickname, therefore it is from Message Log
    else { 
      var j = messageLogData.length;

      while(j--) {      
          var row = 4+j;
          Logger.log("FOUND A MESSAGE LOG ROW")
          cellLinks.push(['https://docs.google.com/spreadsheets/d/'+spreadsheetId+'/edit#gid='+MLGID+'&range='+row+':'+row]);
          j=0;
        }
        else if(j==0) {       
          cellLinks.push(['']);
          j=0;
        }
      }  
    searchToolSS.getRange("I7:I").clear();

    if(cellLinks.length >0) {
      searchToolSS.getRange(7, 9, cellLinks.length, 1).setValues(cellLinks);
    }

    SpreadsheetApp.flush();
}

// Build sources if both Message Log and Incoming are checked
function buildDataSources2(checks) {
  var sourceStringStart = "{";
  var sourceStringEnd = "}";
  var hasMessageLog = false;

  //if both checks are checked use this
  if(checks[0][0] === true && checks[0][1] === true) {
    try{
      copyMLtoHidden();
    } catch(e) {
      Logger.log("Empty Message Log")
    }
    try{
      copyITtoHidden();
    } catch(e) {
      Logger.log("Empty Text Log")
    }
    var sourceStringStart = "'Search Storage Hidden'!A4:I" + ";" + "'Search Storage Hidden'!K15:S";

    return sourceStringStart;
  }
  else {
    for(var i=0; i<checks[0].length; i++) {
      if(checks[0][i] === true) {
        if(i == 0){
          copyMLtoHidden();
          var x = "'Search Storage Hidden'!A4:I;";

          sourceStringStart = sourceStringStart+x;
          hasMessageLog = true;
        }
        else if(i == 1){
          var x = "'Incoming'!A15:I;";

          sourceStringStart = sourceStringStart+x;
        }
      }
    }
    var finalsourceString = sourceStringStart.replace(/\;$/,"");
    
    finalsourceString = finalsourceString+sourceStringEnd;

    return finalsourceString;
  }
}


function buildDataSources3(checks) {
  var sourceStringStart = "{";
  var sourceStringEnd = "}";

  for(var i=0; i<checks[0].length; i++) {
    if(checks[0][i] === true) {
      if(i == 0){
        var x = "'Message Log'!$G$4:$G;";

        sourceStringStart = sourceStringStart+x;
      }
      else if(i == 1){
        var x = "'Incoming'!$G$15:$G;";

        sourceStringStart = sourceStringStart+x;
      }
    }
  }

  var finalsourceString = sourceStringStart.replace(/\;$/,"");

  finalsourceString = finalsourceString+sourceStringEnd;

  return finalsourceString;
}
