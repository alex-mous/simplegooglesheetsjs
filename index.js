/**
 * @module simplegooglesheetsjs
 * @requires googleapis
 */
const {google} = require("googleapis");
const Headers = require("./Headers.js");

/**
* SimpleGoogleSheets default constructor
* @constructor
* @memberof module:simplegooglesheetsjs
* @class SimpleGoogleSheets
* @classdesc  SimpleGoogleSheets is a simplified wrapper for the Google Sheets API
*/
class SimpleGoogleSheets {
    #headers;
    #metadata;
    #spreadsheetId;
    #sheets;
    #sheetId;
    #sheetName;

    constructor() {
        this.#spreadsheetId = null;
        this.#sheets = null; //No sheets object yet - initialized during authentication
        this.#sheetId = null;
        this.#sheetName = null; //No current sheet
        this.#metadata = [];
        this.#headers = null; //Headers of sheet
    }

    /* ========== Authorize Functions ========== */

    /**
     * Authorize with a service account from an email and a private key
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~authorizeServiceAccount
     * @param {string} client_email The service account email
     * @param {string} private_key The service account private key
     */
    authorizeServiceAccount = (client_email, private_key) => {
        const auth = new google.auth.JWT({email: client_email, key: private_key, scopes: ['https://www.googleapis.com/auth/spreadsheets']})
        this.#setSheetAuth(auth);
    }

    /**
     * Authorize with a service account from a file
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~authorizeServiceAccountFile
     * @param {string} keyFile Name of the file to get the client_email and private_key from
     */
    authorizeServiceAccountFile = (keyFile) => {
        let auth = new google.auth.GoogleAuth({keyFile: keyFile, scopes: ['https://www.googleapis.com/auth/spreadsheets']}); //Google Authentication for API
        this.#setSheetAuth(auth);
    }

    /**
     * Authorize with an API Key
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~authorizeAPIKey
     * @param {string} key API Key
     */
    authorizeAPIKey = (key) => {
        this.#setSheetAuth(key);
    }

    


    /* ========== Private Helper Functions ========== */

    /**
     * Set the sheet from the auth
     * 
     * @function setSheets
     * @private
     * @param {Object} auth Google Auth object
     */
    #setSheetAuth = (auth) => {
        this.#sheets = google.sheets({version: 'v4', auth}).spreadsheets; //Sheets API
    }

    /**
     * Update the meta data object with current meta data
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~updateMetaData
     * @private
     */
    #updateMetaData = () => {
        return new Promise((resolve, reject) => {
           this.#sheets.get({
                spreadsheetId: this.#spreadsheetId
            }, (err, res) => {
                if (err) {
                    reject(err);
                } else {
                    this.#metadata = res.data;
                    resolve();
                }
            });
        })
    }

    /**
     * Get the sheet name from the meta data and id
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~getSheetName
     * @param {number} id Integer id of the sheet
     * @private
     * @returns {string} Name of sheet (or undefined if not found)
     */
    #getSheetName = (id) => {
        return this.#metadata.sheets[id].properties.name;
    }

    /**
     * Get the sheet id from the meta data and name
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~getSheetId
     * @param {string} name Name of the sheet
     * @private
     * @returns {number} Id of sheet (or undefined if not found)
     */
    #getSheetId = (name) => {
        let sheet = this.#metadata.sheets.filter((s) => { return s.properties.title == name })[0];
        return (sheet ? sheet.properties.index : undefined); //Return the sheet index (id) if it exists or undefined if it doesn't
    }

    /**
     * Run the updated request on the current spreadsheet
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~updateRequest
     * @private
     * @param {Object} req Request object
     */
    #updateRequest = (req) => {
        return new Promise((resolve, reject) => {
            let reqs = [];
            reqs.push(req);
            this.#sheets.batchUpdate({
                spreadsheetId: this.#spreadsheetId,
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

    /**
     * Read the range in the spreadsheet
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~readRange
     * @private
     * @param {string} range Range (ex. "A1:J4") to read from
     * @return {Promise} Array of rows (array of cells)
     */
    #readRange = (range) => {
        return new Promise((resolve, reject) => {
            this.#sheets.values.get({
                spreadsheetId: this.#spreadsheetId,
                range: this.#sheetName + "!" + range
            }, (err, result) => {
                if (err)
                    reject(err);
                else
                    resolve(result.data.values);
            });
        });
    }


    /**
     * Write the range in the spreadsheet
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~writeRange
     * @private
     * @param {string} range Range (ex. "A1:J4") of data to update
     * @param {string} inputType How to parse the input data ("RAW" for storage as string, "USER_ENTERED" for parsing into date, numbers, formulas, currencies, etc.)
     * @param {Array<Array<string>>} data Data to write into the sheet (must be same dimensions as range)
     * @returns {Promise} Undefined on success
     */
    #writeRange = (range, inputType, data) => {
        return new Promise((resolve, reject) => {
            this.#sheets.values.update({
                spreadsheetId: this.#spreadsheetId,
                range: this.#sheetName + "!" + range,
                valueInputOption: inputType, //Default to raw, so Google Sheets will take data as text rather than parsing a value from it
                resource: {
                    values: data
                }
            }, (err, _) => {
                if (err)
                    reject(err);
                else
                    resolve();
            });
        });
    }

    /**
     * Refresh the current spreadsheet headers
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~refreshHeaders
     * @returns {Promise} Promise of status
     */
    refreshHeaders = () => {
        return new Promise((resolve, reject) => {
            this.#readRange("1:1").then((res) => {
                let headers = new Headers();
                res[0].forEach((val, i) => {
                    if (val && val.length > 0) {
                        if (!headers.addHeader(val, i)) {
                            headers.addHeader(val+i, i); //Add index if the header exists
                        }
                    }
                });
                this.#headers = headers;
                resolve();
            }).catch((err) => reject(err));
        });
    }

    /**
     * Set the row with the data in the array. Warning: will overwrite any headers (on row 0)
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~setRowArray
     * @param {number} rowIndex Index of row
     * @param {Array<string>} dataArray Data to set in cells (index=column number)
     * @returns {Promise} Status of function
     */
    setRowArray = (rowIndex, dataArray) => {
        if (rowIndex < 1) {
            throw new Error("Invalid row index. Must be >= 1");
        } else {
            if (rowIndex == 1) { //Overwrite the headers
                this.#headers = null;
            }
            let range = this.getRangeName(rowIndex, 0, rowIndex, dataArray.length-1);
            return this.#writeRange(range, "RAW", [dataArray]);
        }
    }

    /**
     * Set the row with the data in the format {"HEADER_1":"data1", "HEADER_N":"datan"}
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~setRow
     * @param {number} rowIndex Index of row
     * @param {HeaderData} data Data to set in cells (in HEADER_NAME:VALUE pairs)
     * @returns {Promise} Status of function
     */
    setRow = (rowIndex, data) => {
        let dataArray = new Array(Object.keys(data).length);
        for (let key in data) {
            dataArray[this.#headers.getHeaders()[key]] = data[key];
        }
        return this.setRowArray(rowIndex, dataArray);
    }

    /**
     * Get the row and return the data in the array
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~getRowArray
     * @param {number} rowIndex Index of row
     * @returns {Promise<Array>} Array of data
     */
    getRowArray = (rowIndex) => {
        return new Promise((resolve, reject) => {
            this.#readRange(this.getRangeName(rowIndex, undefined, rowIndex, undefined))
                .then((res) => resolve(res[0]))
                .catch((e) => reject(e));
        });
    }

    /**
     * Get the row and return the data in the format {"HEADER_1":"data1", "HEADER_N":"datan"}
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~getRow
     * @param {number} rowIndex Index of row
     * @returns {Promise<Object>} Data
     */
    getRow = (rowIndex) => {
        return new Promise((resolve, reject) => {
            this.getRowArray(rowIndex).then((res) => {
                let data = {};
                res.forEach((val, i) => {
                    if (val && val.length > 0) {
                        data[this.#headers.getHeadersByColumn()[i]] = val;
                    }
                });
                resolve(data);
            }).catch((e) => reject(e));
        });
    }

    /**
     * Set the spreadsheet headers
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~setHeaders
     * @param {Headers} headers Headers to set
     * @returns {Promise} Promise of update status
     */
    setHeaders(headers) {
        return new Promise((resolve, reject) => {
            if (headers) {
                this.setRowArray(1, headers.getHeadersByColumn()).then(() => {
                    resolve();
                    this.#headers = headers;
                }).catch((err) => reject(err));
            } else {
                reject("No headers supplied to set");
            }
        });
    }


     
    /* ========== Create Functions ========== */

    /**
     * Create a new spreadsheet and set it as current
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~createAndSetSpreadsheet
     * @param {string} name 
     * @returns {Promise} Undefined on success
     */
    createSpreadsheet = (name) => {
        return new Promise((resolve, reject) => {
            this.#sheets.create({
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
                    this.#spreadsheetId = res.spreadsheetId;
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
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~createSheet
     * @param {string} name String name of sheet
     * @returns {Promise} Undefined on success
     */
    createSheet = (name) => {
        //TODO: add create function
        this.#sheetId = this.#metadata.sheets.length;
        this.#sheetName = name;
        this.#updateMetaData();
    }



    /* ========== Delete Functions ========== */

    /**
     * Delete the sheet referenced by the name or the ID
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~deleteSheet
     * @param {string|number} nameOrId Integer ID or string name of sheet
     * @returns {Promise} Undefined on success
     */
    deleteSheet = (nameOrId) => {
        //TODO: add delete function
        if (typeof nameOrId == 'number') { //ID
            let id = nameOrId;
        } else { //Name
            nameOrId = (this.#metadata.sheets.map((sheet) => {
                console.log(sheet);
            }));
            
        }
    }

    /**
     * Delete the row in the spreadsheet. Warning: deleting row 0 will remove the spreadsheet headers (if any)
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~deleteRow
     * @param {number} rowIndex Index of row to delete
     * @returns {Promise} Undefined on success
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
     * Delete the column in the spreadsheet
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~deleteCol
     * @param {string} colIndex Index of column to delete
     * @returns {Promise} Undefined on success
     */
    deleteColumn = (colIndex) => {
        return new Promise((resolve, reject) => {
            let request = {
                deleteDimension: {
                    range: {
                        sheetId: this.#sheetId,
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



    

    /**
     * Find and replace the data in the spreadsheet
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~findAndReplace
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
                    sheetId: allSheets ? undefined : this.#sheetId
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


    /**
     * Set the current sheet, referenced by the name or the id. Also updates the headers
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~setSheet
     * @param {string|number} nameOrId Integer id or string name of sheet
     * @returns {Promise} Undefined on success
     */
    setSheet = (nameOrId) => {
        if (typeof nameOrId == 'number') { //ID
            if (this.#getSheetName(nameOrId) == undefined) {
                throw new Error("no sheet exists for the id ", nameOrId);
            }
            this.#sheetId = nameOrId;
            this.#sheetName = this.#getSheetName(nameOrId);
        } else { //Name
            if (this.#getSheetId(nameOrId) == undefined) {
                throw new Error("no sheet exists for the name ", nameOrId);
            }
            this.#sheetId = this.#getSheetId(nameOrId); //Get the sheet that matched the name
            this.#sheetName = nameOrId;
        }
        return this.refreshHeaders();
    }

    /**
     * Set the current spreadsheet
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~setSpreadsheet
     * @param {string} spreadsheetId ID of spreadsheet to use
     * @returns {Promise} Undefined on success
     */
    setSpreadsheet = (spreadsheetId) => {
        return new Promise((resolve, reject) => {
            this.#spreadsheetId = spreadsheetId;
            this.#updateMetaData()
                .then(() => resolve())
                .catch((err) => reject(err));
        });
    }


    /**
     * Get the current spreadsheet meta data
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~metadata
     * @returns {Object} Metadata of current spreadsheet
     */
    getMetadata() {
        return this.#metadata;
    }

    /**
     * Get the name of the column (such as A, ABZ, or ZAD) from the index
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~getColumnName
     * @throws Error if invalid index provided
     * @param {number} columnIndex Get the column name from the index
     * @returns Column name (such as ABZ)
     */
    getColumnName = (columnIndex) => {
        let rem = columnIndex; //Remainder 
        let res = "";
        if (rem<0 || rem>18277) { //Check for valid index
            throw new Error("Column index must be between 0 and 18277, inclusive");
        }
        if (rem>701) { //Third value
            let c3 = Math.min(26, Math.floor(rem/676));
            res += String.fromCharCode('A'.charCodeAt(0)+c3-1);
            rem -= c3*676;
        }
        if (rem>25) { //Second value
            let c2 = Math.min(26, Math.floor(rem/26));
            res += String.fromCharCode('A'.charCodeAt(0)+c2-1);
            rem -= c2*26;
        } else if (columnIndex>701) { //Edge case for large r but no second value
            res += "A";
        }
        res += String.fromCharCode('A'.charCodeAt(0)+rem); //First value
        return res;
    }

    /**
     * Get the name of the range (in sheets format) from indexes
     * 
     * @function module:simplegooglesheetsjs.SimpleGoogleSheets~getRangeName
     * @throws Error if invalid values provided
     * @param {number} rowStart Starting row (and/or columnStart)
     * @param {number} columnStart Starting column
     * @param {number} rowEnd Ending row (and/or columnEnd)
     * @param {number} columnEnd Ending column
     * @returns {string} Range
     */
    getRangeName = (rowStart, columnStart, rowEnd, columnEnd) => {
        if ((!rowStart || rowStart >= 1) && (!columnStart || columnStart >= 0) && (!rowEnd || rowEnd >= 1) && (!columnEnd || columnEnd >= 0)) { //Ensure that either the values are set
            let res = "";
            res += (columnStart >= 0) ? this.getColumnName(columnStart) : "";
            res += (rowStart >= 1) ? rowStart : "";
            res += !(rowEnd >= 1 || columnEnd >= 0) ? "" : (
                ":" +
                ((columnEnd >= 0) ? this.getColumnName(columnEnd) : "") +
                ((rowEnd >= 1) ? rowEnd : "")
            );
            return res;
        } else {
            throw new Error("Please ensure that the rows and columns are greater than or equal to 1 and 0, respectively");
        }
    }

    /**
     * Get the current spreadsheet headers (run refreshHeaders to update to latest headers)
     */
    getHeaders() {
        return this.#headers;
    }
}

module.exports = SimpleGoogleSheets;