const simplegooglesheets = require("./index.js");


let spreadsheet = new simplegooglesheets(); //Initialize the spreadsheet object

//Next, authorize the spreadsheet access
spreadsheet.authorizeAPIKey("MY_API_KEY");
//Or, authorize with a service account
spreadsheet.authorizeServiceAccount("CLIENT_EMAIL", "PRIVATE_KEY");
//Or, authorize with a JSON key file (containing fields client_email and private_key)
spreadsheet.authorizeServiceAccount("/path/to/key/file.json");


//Use an existing spreadsheet
const spreadsheetId = "MY_SHEET_ID";
spreadsheet.setSpreadsheet(spreadsheetId).then(() => {
    spreadsheet.setSheet("MY_SHEET_NAME").then(() => {
        //Run functions, such as findAndReplace
    });
});

//Or create a new one
spreadsheet.createAndSetSpreadsheet("My new spreadsheet");

