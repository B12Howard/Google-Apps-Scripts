/**
* @OnlyCurrentDoc
*/

/*
Pull Project data from Insightly. Including CustomFields.
*/
var sheetName = "Sheet1";

function connectToInsightly(endp) {
  
  var key = "YOUR-API-KEY";
  var api = "https://api.insightly.com/v3.1/" + endp;
  
  var response = UrlFetchApp.fetch(api, {
    muteHTTPexceptions: true,
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(key),
      'Content-Type': 'application/json'
    }
  });
  
  return JSON.parse(response.getContentText());
  
}

function getProjectDetails() {
  var skip = 0;
  var top = 500;
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  var projName = [];
  var dateStarted = [];
  var stage = [];
  var customField1Array = [];
  var customField2Array = [];
  var customField3Array = [];
  
  // Return nothing if there is an error getting data from API
  try {
    var pipelineStages = connectToInsightly("PipelineStages");
  } catch(e) {
    console.log(e)
    return;
  }
  
  var pipelineDict = {};
  
  // Stages are stored as an id number in Projects. Get the English name of the stage from this endpoint
  // Use dict for faster lookup
  for(var i=0; i<pipelineStages.length; i++) {
    pipelineDict[pipelineStages[i].STAGE_ID] = pipelineStages[i].STAGE_NAME
  }

  // API pagination. Skip is where to start. Top is the number of items to get (500 max)
  while (true) {
    var endp = "Projects?skip=" + skip + "&top=500";
    
    // Return nothing if there is an error getting data from API
    try {
      var projects = connectToInsightly(endp);
    } catch(e) {
      console.log(e)
      return;
    }
    
    if(!projects.length) break;
    
    // parse json response from api
    for (var p in projects) {
      projName.push(projects[p].PROJECT_NAME ? [projects[p].PROJECT_NAME]: [""]);     
      dateStarted.push(projects[p].STARTED_DATE ? [projects[p].STARTED_DATE]: [""]);

      // Handle Stage
      if(projects[p].STAGE_ID) {
        stage.push([pipelineDict[""+projects[p].STAGE_ID]]);
      }
      else {
        stage.push([""])
      }
      
      var customField1 = "";
      var customField2= "";
      var customField3 = "";
      
      for(var i=0; i<projects[p].CUSTOMFIELDS.length; i++) {
        // Handle Addresses
        if(projects[p].CUSTOMFIELDS[i].FIELD_NAME == "CUSTOM_FIELD_NAME_1") {
          customField1 = projects[p].CUSTOMFIELDS[i].FIELD_VALUE
          customField1Array.push([customField1]);
        }
        
        // Handle Jurisdiction
        if(projects[p].CUSTOMFIELDS[i].FIELD_NAME == "CUSTOM_FIELD_NAME_2") {
          customField2 = projects[p].CUSTOMFIELDS[i].FIELD_VALUE;
          customField2Array.push([customField2]);
        }

        // Handle Latest Update
        if(projects[p].CUSTOMFIELDS[i].FIELD_NAME == "CUSTOM_FIELD_NAME_3") {
          customField3 = projects[p].CUSTOMFIELDS[i].FIELD_VALUE;
          customField3Array.push([customField3]);
        }
      }
      
      // If in CUSTOMFIELDS there isn't an address, jurisdiction or latest update
      if(!customFieldValue1) { customField1Array.push([""]) }
      if(!customFieldValue2) { customField2Array.push([""]) }
      if(!customFieldValue3) { customField3Array.push([""]) }
    }
    skip += top;
  }
  
  // In case of error and the data is blank
  if(projName.length > 0) { 
  
    // Catch error clearing if the sheet is blank from row 2 onwards
    try {
      ss.getRange(2, 1, ss.getLastRow()-1, ss.getLastColumn()).clearContent();
    } catch(e) {
      console.log("No lines to clear")
    }
    
    ss.getRange(2, 1, projName.length, 1).setValues(projName);
    ss.getRange(2, 2, customField1Array.length, 1).setValues(customField1Array);  
    ss.getRange(2, 3, customField2Array.length, 1).setValues(customField2Array);
    ss.getRange(2, 4, dateStarted.length, 1).setValues(dateStarted);
    ss.getRange(2, 5, stage.length, 1).setValues(stage);
    ss.getRange(2, 6, customField3Array.length, 1).setValues(customField3Array);
  }
}
