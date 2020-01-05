/* Script that goes through all the tabs in a Sheet and sends out emails depending on variables in the Sheet. */

/*Global Variables*/

/* Tab names to ignore*/
var ignore1 = "";
var ignore2 =  "";
var ignore3 = "";

var ss = SpreadsheetApp.getActiveSpreadsheet();

/* Used to designate what status/ email to send */
var OS = "OS"; 
var PS = "PS";
/* Used to mark on the Sheet the email was sent */
var osEmailSent = "Order Status Email Sent";
var psEmailSent = "Project Status Email Sent"

/* Where to send errors*/
var errorEmailAddresses = ["a@fake.com", "b@fake.com"];

/* A criteria to match to send the email*/
var needFeedback = "";
var needOrderNum = "";


/* OBJECTS*/
/* For keeping track of errored emails*/
function errorObject(e,t,p) {
  this.email = e;
  this.type = t;
  this.product = p;
}


/* Start here*/
function main() {
  getTabNames();
}

function getTabNames() {
  var errorEmail = [];
  var tabNames = ss.getSheets();
  
  /* Iterate over the tab names */
  for(var i in tabNames) {
    var tempSheetName = tabNames[i].getSheetName();

    if(tempSheetName != ignore1 && tempSheetName != ignore2 && tempSheetName != ignore3) {
      var activeSheet = ss.getSheetByName(tempSheetName);
      var data = activeSheet.getRange(1,1,activeSheet.getLastRow(), activeSheet.getMaxColumns()).getValues();
      var productName = data[2][1];

      /* Iterate over data in the tab 
      	In this case it starts at row 14, it can be changed.
      */
      for(var i=13; i<data.length; i++) {
        var emailSuccess = false;
        var tempType = "";

        /* Case of Order Status */
        /* Put your criterias to send an email here! */	
        if(data[i][2].toString() == needOrderNum && data[i][21].toString() == "" && data[i][8] != "") {
          emailSuccess = email(data[i], tempSheetName, OS, productName);
          tempType = OS;

          if(emailSuccess) {
            var row = i+1;
            activeSheet.getRange(i+1, 22).setValue(osEmailSent);
          }
          else {
            errorEmail.push(new errorObject(data[i][8],tempType,productName));
          }
        }
        /* Case of Project Status */
        /* Put your criterias to send an email here! */
        else if(data[i][3].toString() == needFeedback && data[i][22].toString() == "" && data[i][8] != "") {
          emailSuccess = email(data[i], tempSheetName, PS, productName);
          tempType = PS;

          if(emailSuccess) {
            var row = i+1;
            activeSheet.getRange(i+1, 23).setValue(psEmailSent);
          }
          else {
            errorEmail.push(new errorObject(data[i][8],tempType,productName));
          }
        }
        else continue;
      }
    }
  }
  
  if(errorEmail.length > 0) {
    sendErrorEmail(errorEmail)
  }
}

function email(rowData, tabName, type, productName) {
 var firstName = rowData[4].toString().ucfirst();
 var to = rowData[8].toString().trim();
  
  switch (type) {
    case 'OS':
      var message = '<p>Hi ' +  + ',</p>'
      + '<p>' +  + '</p>'
      + '<p>'
      + '</p>'

      
      break;
      
    case 'PS':
      var message = '<p>Hi ' +  + ',</p>'
      + '<p>' +  + "</p>"
      + '<p></p>'
      + '<p></p>'
         
      break;
  }
  
  var signoff = '<br>'
  + '<p></p>'
  + '<p></p>';
  
  message = message + signoff;
  
  var subject = '';
  
  try {
    MailApp.sendEmail(to, subject, message, {
      htmlBody: message
    }); 
    return type;
  } catch(e) {
    Logger.log(e)
    return 0;
  }
}

function sendErrorEmail(errorEmailAddressArray) {
  var errString = "";
  
  for(var i=0; i<errorEmailAddressArray.length; i++) {
    var status = errorEmailAddressArray[i].type == OS ? "Order Status" : "Project Status";
    errString = errString + errorEmailAddressArray[i].email + "\t" + status + "\t" + errorEmailAddressArray[i].product + "\n";
  }

  MailApp.sendEmail(errorEmailAddresses, "Error Sending Email", "There was an error sending an email to these addresses:\n\n" + "Email Address  \t\t" + "Status  \t\t" + "Product Name\n\n" + errString); 
  
}
