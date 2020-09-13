/**
 * SimpleGoogleSheets module
 * @module simplegooglesheets
 */

const {google} = require("googleapis");

/**
 * 
 * @class
 * @classdesc Simplified wrapper for the Google Sheets API
 */
class SimpleGoogleSheets {
    /**
     * Default constructor
     * 
     * @constructor
     */
    constructor() {
        this.spreadsheetId = null;
        this.sheets = null; //No sheets object yet - initialized during authentication
        this.sheetId = null;
        this.sheetName = null; //No current sheet
        this.metadata = {}; //No meta data yet
    }

    /* ========== Authorize Functions ========== */

    /**
     * Authorize with a service account from an email and a private key
     * 
     * @function authorizeServiceAccount
     * @param {string} client_email The service account email
     * @param {string} private_key The service account private key
     */
    authorizeServiceAccount = (client_email, private_key) => {
        let auth = new google.auth.GoogleAuth({client_email: client_email, private_key: private_key, scopes: ['https://www.googleapis.com/auth/spreadsheets']}); //Google Authentication for API
        this.sheets = google.sheets({version: 'v4', auth}).spreadsheets; //Sheets API
    }

    /**
     * Authorize with a service account from a file
     * 
     * @function authorizeServiceAccount
     * @param {string} keyFile Name of the file to get the client_email and private_key from
     */
    authorizeServiceAccount = (keyFile) => {
        let auth = new google.auth.GoogleAuth({keyFile: keyFile, scopes: ['https://www.googleapis.com/auth/spreadsheets']}); //Google Authentication for API
        this.sheets = google.sheets({version: 'v4', auth}).spreadsheets; //Sheets API
    }

    /**
     * Authorize with an API Key
     * 
     * @function authorizeAPIKey
     * @param {string} key API Key
     */
    authorizeAPIKey = (key) => {
        this.sheets = google.sheets({version: 'v4', key}).spreadsheets; //Sheets API
    }



    /* ========== Private Helper Functions ========== */

    /**
     * Update the meta data object with current meta data
     * 
     * @private
     */
    #updateMetaData = () => {
        return new Promise((resolve, reject) => {
           this.sheets.get({
                spreadsheetId: this.spreadsheetId
            }, (err, res) => {
                if (err) {
                    reject(err);
                } else {
                    this.metadata = res.data;
                    resolve();
                }
            });
        })
    }

        /**
     * Get the sheet name from the meta data and id
     * 
     * @param {number} id Integer id of the sheet
     * @returns {string} Name of sheet (or undefined if not found)
     */
    #getSheetName = (id) => {
        return this.metadata.sheets[id].properties.name;
    }

    /**
     * Get the sheet id from the meta data and name
     * 
     * @param {string} name Name of the sheet
     * @returns {number} Id of sheet (or undefined if not found)
     */
    #getSheetId = (name) => {
        let sheet = this.metadata.sheets.filter((s) => { return s.properties.title == name })[0];
        return (sheet ? sheet.properties.index : undefined); //Return the sheet index (id) if it exists or undefined if it doesn't
    }

    /**
     * Run the updated request on the current spreadsheet
     * 
     * @function updateRequest
     * @private
     * @param {Object} req Request object
     */
    #updateRequest = (req) => {
        return new Promise((resolve, reject) => {
            let reqs = [];
            reqs.push(req);
            const batchReq = {reqs};
            this.sheets.batchUpdate({
                spreadsheetId: this.spreadsheetId,
                resource: JSON.stringify({
                    requests: reqs
                })
            }, (err, res) => {
                if (err) {
                    reject(err);
                } else {
                    resolve(res);
                }
            });
        });
    }



    /* ========== Create Functions ========== */

    /**
     * Create a new spreadsheet and set it as current
     * 
     * @function createAndSetSpreadsheet
     * @param {string} name 
     */
    createAndSetSpreadsheet = (name) => {
        return new Promise((resolve, reject) => {
            this.sheets.create({
                resource: JSON.stringify({
                    properties: {
                        title: name
                    }
                }),
                fields: "spreadsheetId"
            }, (err, res) => {
                if (err) {
                    reject(err);
                } else {
                    this.spreadsheetId = res.spreadsheetId;
                    this.#updateMetaData(request)
                        .then(() => resolve())
                        .catch((err) => reject(err));
                }
            });
        });
    }

    /**
     * Create and set as current sheet
     * 
     * @param {string} name String name of sheet
     */
    createSheet = (name) => {
        this.sheetId = this.metadata.sheets.length;
        this.sheetName = nameOrId;
        this.#updateMetaData();
    }




    /* ========== Delete Functions ========== */

    /**
     * Delete the sheet referenced by the name or the ID
     * 
     * @param {*} nameOrId Integer ID or string name of sheet
     */
    deleteSheet = (nameOrId) => {
        if (typeof nameOrId == 'number') { //ID
            let id = nameOrId;
        } else { //Name
            nameOrId = (this.metadata.sheets.map((sheet) => {
                console.log(sheet);
            }));
            
        }
    }

    /**
     * Delete the row in the spreadsheet
     * 
     * @function deleteRow
     * @param {number} rowIndex Index of row to delete
     * @return {Promise} Promise object representing the status of the function
     */
    deleteRow = (rowIndex) => {
        return new Promise((resolve, reject) => {
            let request = {
                deleteDimension: {
                    range: {
                        sheetId: this.currSheet,
                        startIndex: rowIndex,
                        endIndex: rowIndex+1,
                        dimension: "ROWS"
                    }
                }
            };
            this.#updateRequest(request)
                .then(() => resolve())
                .catch((err) => reject(err));
        });
    }

    /**
     * Delete the row in the spreadsheet
     * 
     * @function deleteCol
     * @param {number} colIndex Index of column to delete
     * @return {Promise} Promise object representing the status of the function
     */
    deleteCol = (colIndex) => {
        return new Promise((resolve, reject) => {
            let request = {
                deleteDimension: {
                    range: {
                        sheetId: this.sheetId,
                        startIndex: colIndex,
                        endIndex: colIndex+1,
                        dimension: "COLUMNS"
                    }
                }
            };
            this.#updateRequest(request)
                .then(() => resolve())
                .catch((err) => reject(err));
        });
    }



    /* ========== Read Functions ========== */
    
    /**
     * Read the range in the spreadsheet
     * 
     * @function readRange
     * @param {string} range Range (ex. "A1:J4") to read from
     * @return {Promise} Promise object representing the status of the function
     */
    readRange = (range) => {
        return new Promise((resolve, reject) => {
            this.sheets.values.get({
                spreadsheetId: this.spreadsheetId,
                range: this.sheetName + "!" + range
            }, (err, result) => {
                if (err)
                    reject(err);
                else
                    resolve(result.data.values);
            });
        });
    }

    /* ========== Write Functions ========== */

    /**
     * Write the range in the spreadsheet
     * 
     * @function writeRange
     * @param {string} range Range (ex. "A1:J4") of data to update
     * @param {string} inputType How to parse the input data ("RAW" for storage as string, "USER_ENTERED" for parsing into date, numbers, formulas, currencies, etc.)
     * @param {Array<Array<string>>} data Data to write into the sheet (must be same dimensions as range)
     * @return {Promise} Promise object representing the status of the function
     */
    writeRange = (range, inputType, data) => {
        return new Promise((resolve, reject) => {
            this.sheets.values.update({
                spreadsheetId: spreadsheetId,
                range: this.sheetName + "!" + range,
                valueInputOption: inputType, //Default to raw, so Google Sheets will take data as text rather than parsing a value from it
                resource: data
            }, (err, _) => {
                if (err)
                    reject(err);
                else
                    resolve();
            });
        });
    }



    /* ========== Setters ========== */

    /**
     * Set the current sheet, referenced by the name or the id
     * 
     * @function setSheet
     * @param {string|number} nameOrId Integer id or string name of sheet
     */
    setSheet = (nameOrId) => {
        return new Promise((resolve, reject) => {
            if (typeof nameOrId == 'number') { //ID
                if (this.#getSheetName(nameOrId) == undefined) {
                   reject("no sheet exists for the id ", nameOrId);
                }
                this.sheetId = nameOrId;
                this.sheetName = this.#getSheetName(nameOrId);
                resolve();
            } else { //Name
                if (this.#getSheetId(nameOrId) == undefined) {
                    reject("no sheet exists for the name ", nameOrId);
                }
                this.sheetId = this.#getSheetId(nameOrId); //Get the sheet that matched the name
                this.sheetName = nameOrId;
                resolve();
            }
        });
    }

    /**
     * Set the current spreadsheet
     * 
     * @function setSpreadsheet
     * @param {string} spreadsheetId ID of spreadsheet to use
     * @returns {Promise} Promise of successful setting of id
     */
    setSpreadsheet = (spreadsheetId) => {
        return new Promise((resolve, reject) => {
            this.spreadsheetId = spreadsheetId;
            this.#updateMetaData()
                .then(() => resolve())
                .catch((err) => reject(err));
        });
    }



    /* ========== Getters ========== */

    /**
     * Get the current spreadsheet meta data
     * 
     * @function getSpreadsheetMetaData
     * @returns {Object} Metadata
     */
    get getSpreadsheetMetaData() {
        return this.metadata;
    }

    /**
     * Find and replace the data in the spreadsheet
     * 
     * @function findAndReplace
     * @param {string} find The string to search for
     * @param {string} replacement Replacement string for find and replace
     * @param {boolean} allSheets Search all sheets or just the current one
     * @param {boolean} matchCase Match the case
     * @param {boolean} entireCell Match the entire cell
     * @param {boolean} useRegex Use regex to search the cell (find must be a regex string)
     * @returns {Promise} Promise of number of values replaced
     */
    findAndReplace = (find, replacement, allSheets, matchCase, entireCell, useRegex) => {
        return new Promise((resolve, reject) => {
            let request = {
                findReplace: {
                    find: find,
                    replacement: replacement,
                    matchCase: matchCase,
                    matchEntireCell: entireCell,
                    searchByRegex: useRegex,
                    allSheets: allSheets ? allSheets : undefined,
                    sheetId: allSheets ? undefined : this.sheetId
                }
            };
            this.#updateRequest(request).then((res) => {
                let vals = res.data.replies[0].findReplace.valuesChanged;
                resolve(vals ? vals : 0); //If it is undefined, none were replaced
            }).catch((err) => {
                reject(err);
            });
        });
    }
}

module.exports = SimpleGoogleSheets;