function sendButton() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName();
  var range = sheet.getDataRange();
  var method = range.getCell(2,2).getValue();
  var recipient = range.getCell(3,2).getValue(); //To
  var subject = range.getCell(4,2).getValue(); //Heading
  var mailFROM =  GmailApp.getAliases()[0];
  var ui = SpreadsheetApp.getUi();
  
  
  if(method == "E-Mail") {
    var body = '<p>' + range.getCell(7,1).getValue() <'/p'>;

    try{
      GmailApp.sendEmail(recipient, subject, body,{
        htmlBody: body,
        from:mailFROM
      })
      
      var result = ui.alert(
        'Success',
        'Message Sent',
        ui.ButtonSet.OK);
      
    }
    catch(e) {
      var result = ui.alert(
        'Message Failed to Send',
        'Click Send Again',
        ui.ButtonSet.OK);
    }
  }
  if(method == "Text") {
    var body = range.getCell(8,1).getValue();
    var key = '';
    var YOURACCOUNTSID = "";
    var YOURAUTHTOKEN = "";
    var messages_url = "https://api.twilio.com/2010-04-01/Accounts/"+YOURACCOUNTSID+"/Messages.json";

    var payload = {
      "To": "+1" + recipient,
      "Body" : body,
      "From" : "+"
    };

    var options = {
      "method" : "post",
      "payload" : payload
    };
    options.headers = { 
      "Authorization" : "Basic " + Utilities.base64Encode(YOURACCOUNTSID+":"+YOURAUTHTOKEN)
    };

    try{
     UrlFetchApp.fetch(messages_url, options);
      logActivity();
      var result = ui.alert(
        'Success',
        'Message Sent',
        ui.ButtonSet.OK);
    }
    catch(e){
      Logger.log(e)
      var result = ui.alert(
        'Message Failed to Send',
        'Click Send Again',
        ui.ButtonSet.OK);
    }
  }
  
}
