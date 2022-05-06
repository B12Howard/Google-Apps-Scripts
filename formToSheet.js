// Run this first before making a trigger to get the Form permissions
function getForm() {
  FormApp.getActiveForm()
}

function getFormData(e) {
  const formResponses = e.response.getItemResponses(); 
  const dict = {};
  // Optional email
  // let email = e.response.getRespondentEmail();
  // get responses 
  formResponses.forEach(formResponse => {
    dict[formResponses.getItem().getTitle()] = formResponse.getResponse();
  })
  
  // Clean the data
  dict["First Name"] = dict["First Name"] ? dict["First Name"].trim(): dict["First Name"];
  dict["Last Name"] = dict["Last Name"] ? dict["Last Name"].trim() : dict["Last Name"];
  dict["Phone"] = dict["Phone"] ? dict["Phone"].trim() : dict["Phone"];

  addRowToDb(dict);
}

function openSS() {
  return SpreadsheetApp.openByUrl(databaseSSURL);
}

function findTabName(param) {
  // Optional find tab by param
  // const re = new RegExp(tabName);
  // if(re.test( sheets[j].getName())) {
  //   // Found tab add data
  //   let sheet = ss.getSheetByName(sheets[j].getName());
  //   try {
  //     let insertAtRow = sheet.getLastRow() + 1;
  //     let data = [dict["Last Name"], dict["First Name"], dict["Phone"], "", "", "", "", dict["Points"], dict["Phase"]];
  //     sheet.getRange(insertAtRow, 1, 1, data.length).setValues([data]);
  //     SpreadsheetApp.flush();

  //   }
  //   catch(error) {
  //     console.error(error)
  //   }
  // }
  // console.error("Could not find matching tab for " + tabName)
  return false;
}

function addRowToDb(dict) {
  const ss = SpreadsheetApp.openByUrl(databaseSSURL);
  const sheets = ss.getSheets();
  const tabName = findTabName(param) || "defaultTab";

  const sheet = ss.getSheetByName(tabName);
  
  try {
    let insertAtRow = sheet.getLastRow() + 1;
    let data = [dict["Last Name"], dict["First Name"], dict["Phone"]];
    sheet.getRange(insertAtRow, 1, 1, data.length).setValues([data]);
    SpreadsheetApp.flush();
  }
  catch(error) {
    console.error(error)
  }
}
