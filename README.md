# SimpleGoogleSheetsJS
Create, read, and write Google Sheets through a simplified interface for the Google Sheets API

## Installation
`npm install simplegooglesheetsjs`

## Usage
```
const simplegooglesheets = require("./index.js");


let spreadsheet = new simplegooglesheets(); //Initialize the spreadsheet object

//Next, authorize the spreadsheet access
spreadsheet.authorizeAPIKey("MY_API_KEY");
//Or, authorize with a service account
spreadsheet.authorizeServiceAccount("CLIENT_EMAIL", "PRIVATE_KEY");
//Or, authorize with a JSON key file (containing fields client_email and private_key)
spreadsheet.authorizeServiceAccountFile("/path/to/key/file.json");


//Use an existing spreadsheet
const spreadsheetId = "MY_SHEET_ID";
spreadsheet.setSpreadsheet(spreadsheetId).then(() => {
    spreadsheet.setSheet("MY_SHEET_NAME").then(() => {
        //Print out the headers
        console.log(spreadsheet.getHeaders().Headers);
        
        //Get and print a row by header names
        spreadsheet.getRow(5).then((row) => console.log(row));
    });
});

//Or create a new one
spreadsheet.createSpreadsheet("My new spreadsheet");
```

## Documentation
<a name="module_simplegooglesheetsjs"></a>

## simplegooglesheetsjs
**Requires**: <code>module:googleapis</code>  

* [simplegooglesheetsjs](#module_simplegooglesheetsjs)
    * [.SimpleGoogleSheets](#module_simplegooglesheetsjs.SimpleGoogleSheets)
        * [new SimpleGoogleSheets()](#new_module_simplegooglesheetsjs.SimpleGoogleSheets_new)
        * [~authorizeServiceAccount(client_email, private_key)](#module_simplegooglesheetsjs.SimpleGoogleSheets..authorizeServiceAccount)
        * [~authorizeServiceAccountFile(keyFile)](#module_simplegooglesheetsjs.SimpleGoogleSheets..authorizeServiceAccountFile)
        * [~authorizeAPIKey(key)](#module_simplegooglesheetsjs.SimpleGoogleSheets..authorizeAPIKey)
        * [~refreshHeaders()](#module_simplegooglesheetsjs.SimpleGoogleSheets..refreshHeaders) ⇒ <code>Promise</code>
        * [~setRowArray(rowIndex, dataArray)](#module_simplegooglesheetsjs.SimpleGoogleSheets..setRowArray) ⇒ <code>Promise</code>
        * [~setRow(rowIndex, data)](#module_simplegooglesheetsjs.SimpleGoogleSheets..setRow) ⇒ <code>Promise</code>
        * [~getRowArray(rowIndex)](#module_simplegooglesheetsjs.SimpleGoogleSheets..getRowArray) ⇒ <code>Promise.&lt;Array&gt;</code>
        * [~getRow(rowIndex)](#module_simplegooglesheetsjs.SimpleGoogleSheets..getRow) ⇒ <code>Promise.&lt;Object&gt;</code>
        * [~setHeaders(headers)](#module_simplegooglesheetsjs.SimpleGoogleSheets..setHeaders) ⇒ <code>Promise</code>
        * [~createAndSetSpreadsheet(name)](#module_simplegooglesheetsjs.SimpleGoogleSheets..createAndSetSpreadsheet) ⇒ <code>Promise</code>
        * [~createSheet(name)](#module_simplegooglesheetsjs.SimpleGoogleSheets..createSheet) ⇒ <code>Promise</code>
        * [~deleteSheet(nameOrId)](#module_simplegooglesheetsjs.SimpleGoogleSheets..deleteSheet) ⇒ <code>Promise</code>
        * [~deleteRow(rowIndex)](#module_simplegooglesheetsjs.SimpleGoogleSheets..deleteRow) ⇒ <code>Promise</code>
        * [~deleteCol(colIndex)](#module_simplegooglesheetsjs.SimpleGoogleSheets..deleteCol) ⇒ <code>Promise</code>
        * [~findAndReplace(find, replacement, allSheets, matchCase, entireCell, useRegex)](#module_simplegooglesheetsjs.SimpleGoogleSheets..findAndReplace) ⇒ <code>Promise</code>
        * [~setSheet(nameOrId)](#module_simplegooglesheetsjs.SimpleGoogleSheets..setSheet) ⇒ <code>Promise</code>
        * [~setSpreadsheet(spreadsheetId)](#module_simplegooglesheetsjs.SimpleGoogleSheets..setSpreadsheet) ⇒ <code>Promise</code>
        * [~metadata()](#module_simplegooglesheetsjs.SimpleGoogleSheets..metadata) ⇒ <code>Object</code>
        * [~getColumnName(columnIndex)](#module_simplegooglesheetsjs.SimpleGoogleSheets..getColumnName) ⇒
        * [~getRangeName(rowStart, columnStart, rowEnd, columnEnd)](#module_simplegooglesheetsjs.SimpleGoogleSheets..getRangeName) ⇒ <code>string</code>

<a name="module_simplegooglesheetsjs.SimpleGoogleSheets"></a>

### simplegooglesheetsjs.SimpleGoogleSheets
SimpleGoogleSheets is a simplified wrapper for the Google Sheets API

**Kind**: static class of [<code>simplegooglesheetsjs</code>](#module_simplegooglesheetsjs)  

* [.SimpleGoogleSheets](#module_simplegooglesheetsjs.SimpleGoogleSheets)
    * [new SimpleGoogleSheets()](#new_module_simplegooglesheetsjs.SimpleGoogleSheets_new)
    * [~authorizeServiceAccount(client_email, private_key)](#module_simplegooglesheetsjs.SimpleGoogleSheets..authorizeServiceAccount)
    * [~authorizeServiceAccountFile(keyFile)](#module_simplegooglesheetsjs.SimpleGoogleSheets..authorizeServiceAccountFile)
    * [~authorizeAPIKey(key)](#module_simplegooglesheetsjs.SimpleGoogleSheets..authorizeAPIKey)
    * [~refreshHeaders()](#module_simplegooglesheetsjs.SimpleGoogleSheets..refreshHeaders) ⇒ <code>Promise</code>
    * [~setRowArray(rowIndex, dataArray)](#module_simplegooglesheetsjs.SimpleGoogleSheets..setRowArray) ⇒ <code>Promise</code>
    * [~setRow(rowIndex, data)](#module_simplegooglesheetsjs.SimpleGoogleSheets..setRow) ⇒ <code>Promise</code>
    * [~getRowArray(rowIndex)](#module_simplegooglesheetsjs.SimpleGoogleSheets..getRowArray) ⇒ <code>Promise.&lt;Array&gt;</code>
    * [~getRow(rowIndex)](#module_simplegooglesheetsjs.SimpleGoogleSheets..getRow) ⇒ <code>Promise.&lt;Object&gt;</code>
    * [~setHeaders(headers)](#module_simplegooglesheetsjs.SimpleGoogleSheets..setHeaders) ⇒ <code>Promise</code>
    * [~createAndSetSpreadsheet(name)](#module_simplegooglesheetsjs.SimpleGoogleSheets..createAndSetSpreadsheet) ⇒ <code>Promise</code>
    * [~createSheet(name)](#module_simplegooglesheetsjs.SimpleGoogleSheets..createSheet) ⇒ <code>Promise</code>
    * [~deleteSheet(nameOrId)](#module_simplegooglesheetsjs.SimpleGoogleSheets..deleteSheet) ⇒ <code>Promise</code>
    * [~deleteRow(rowIndex)](#module_simplegooglesheetsjs.SimpleGoogleSheets..deleteRow) ⇒ <code>Promise</code>
    * [~deleteCol(colIndex)](#module_simplegooglesheetsjs.SimpleGoogleSheets..deleteCol) ⇒ <code>Promise</code>
    * [~findAndReplace(find, replacement, allSheets, matchCase, entireCell, useRegex)](#module_simplegooglesheetsjs.SimpleGoogleSheets..findAndReplace) ⇒ <code>Promise</code>
    * [~setSheet(nameOrId)](#module_simplegooglesheetsjs.SimpleGoogleSheets..setSheet) ⇒ <code>Promise</code>
    * [~setSpreadsheet(spreadsheetId)](#module_simplegooglesheetsjs.SimpleGoogleSheets..setSpreadsheet) ⇒ <code>Promise</code>
    * [~metadata()](#module_simplegooglesheetsjs.SimpleGoogleSheets..metadata) ⇒ <code>Object</code>
    * [~getColumnName(columnIndex)](#module_simplegooglesheetsjs.SimpleGoogleSheets..getColumnName) ⇒
    * [~getRangeName(rowStart, columnStart, rowEnd, columnEnd)](#module_simplegooglesheetsjs.SimpleGoogleSheets..getRangeName) ⇒ <code>string</code>

<a name="new_module_simplegooglesheetsjs.SimpleGoogleSheets_new"></a>

#### new SimpleGoogleSheets()
SimpleGoogleSheets default constructor

<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..authorizeServiceAccount"></a>

#### SimpleGoogleSheets~authorizeServiceAccount(client_email, private_key)
Authorize with a service account from an email and a private key

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  

| Param | Type | Description |
| --- | --- | --- |
| client_email | <code>string</code> | The service account email |
| private_key | <code>string</code> | The service account private key |

<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..authorizeServiceAccountFile"></a>

#### SimpleGoogleSheets~authorizeServiceAccountFile(keyFile)
Authorize with a service account from a file

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  

| Param | Type | Description |
| --- | --- | --- |
| keyFile | <code>string</code> | Name of the file to get the client_email and private_key from |

<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..authorizeAPIKey"></a>

#### SimpleGoogleSheets~authorizeAPIKey(key)
Authorize with an API Key

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  

| Param | Type | Description |
| --- | --- | --- |
| key | <code>string</code> | API Key |

<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..refreshHeaders"></a>

#### SimpleGoogleSheets~refreshHeaders() ⇒ <code>Promise</code>
Refresh the current spreadsheet headers

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  
**Returns**: <code>Promise</code> - Promise of status  
<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..setRowArray"></a>

#### SimpleGoogleSheets~setRowArray(rowIndex, dataArray) ⇒ <code>Promise</code>
Set the row with the data in the array. Warning: will overwrite any headers (on row 0)

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  
**Returns**: <code>Promise</code> - Status of function  

| Param | Type | Description |
| --- | --- | --- |
| rowIndex | <code>number</code> | Index of row |
| dataArray | <code>Array.&lt;string&gt;</code> | Data to set in cells (index=column number) |

<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..setRow"></a>

#### SimpleGoogleSheets~setRow(rowIndex, data) ⇒ <code>Promise</code>
Set the row with the data in the format {"HEADER_1":"data1", "HEADER_N":"datan"}

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  
**Returns**: <code>Promise</code> - Status of function  

| Param | Type | Description |
| --- | --- | --- |
| rowIndex | <code>number</code> | Index of row |
| data | <code>HeaderData</code> | Data to set in cells (in HEADER_NAME:VALUE pairs) |

<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..getRowArray"></a>

#### SimpleGoogleSheets~getRowArray(rowIndex) ⇒ <code>Promise.&lt;Array&gt;</code>
Get the row and return the data in the array

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  
**Returns**: <code>Promise.&lt;Array&gt;</code> - Array of data  

| Param | Type | Description |
| --- | --- | --- |
| rowIndex | <code>number</code> | Index of row |

<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..getRow"></a>

#### SimpleGoogleSheets~getRow(rowIndex) ⇒ <code>Promise.&lt;Object&gt;</code>
Get the row and return the data in the format {"HEADER_1":"data1", "HEADER_N":"datan"}

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  
**Returns**: <code>Promise.&lt;Object&gt;</code> - Data  

| Param | Type | Description |
| --- | --- | --- |
| rowIndex | <code>number</code> | Index of row |

<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..setHeaders"></a>

#### SimpleGoogleSheets~setHeaders(headers) ⇒ <code>Promise</code>
Set the spreadsheet headers

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  
**Returns**: <code>Promise</code> - Promise of update status  

| Param | Type | Description |
| --- | --- | --- |
| headers | <code>Headers</code> | Headers to set |

<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..createAndSetSpreadsheet"></a>

#### SimpleGoogleSheets~createAndSetSpreadsheet(name) ⇒ <code>Promise</code>
Create a new spreadsheet and set it as current

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  
**Returns**: <code>Promise</code> - Undefined on success  

| Param | Type |
| --- | --- |
| name | <code>string</code> | 

<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..createSheet"></a>

#### SimpleGoogleSheets~createSheet(name) ⇒ <code>Promise</code>
Create and set as current sheet

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  
**Returns**: <code>Promise</code> - Undefined on success  

| Param | Type | Description |
| --- | --- | --- |
| name | <code>string</code> | String name of sheet |

<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..deleteSheet"></a>

#### SimpleGoogleSheets~deleteSheet(nameOrId) ⇒ <code>Promise</code>
Delete the sheet referenced by the name or the ID

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  
**Returns**: <code>Promise</code> - Undefined on success  

| Param | Type | Description |
| --- | --- | --- |
| nameOrId | <code>string</code> \| <code>number</code> | Integer ID or string name of sheet |

<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..deleteRow"></a>

#### SimpleGoogleSheets~deleteRow(rowIndex) ⇒ <code>Promise</code>
Delete the row in the spreadsheet. Warning: deleting row 0 will remove the spreadsheet headers (if any)

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  
**Returns**: <code>Promise</code> - Undefined on success  

| Param | Type | Description |
| --- | --- | --- |
| rowIndex | <code>number</code> | Index of row to delete |

<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..deleteCol"></a>

#### SimpleGoogleSheets~deleteCol(colIndex) ⇒ <code>Promise</code>
Delete the column in the spreadsheet

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  
**Returns**: <code>Promise</code> - Undefined on success  

| Param | Type | Description |
| --- | --- | --- |
| colIndex | <code>string</code> | Index of column to delete |

<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..findAndReplace"></a>

#### SimpleGoogleSheets~findAndReplace(find, replacement, allSheets, matchCase, entireCell, useRegex) ⇒ <code>Promise</code>
Find and replace the data in the spreadsheet

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  
**Returns**: <code>Promise</code> - Promise of number of values replaced  

| Param | Type | Description |
| --- | --- | --- |
| find | <code>string</code> | The string to search for |
| replacement | <code>string</code> | Replacement string for find and replace |
| allSheets | <code>boolean</code> | Search all sheets or just the current one |
| matchCase | <code>boolean</code> | Match the case |
| entireCell | <code>boolean</code> | Match the entire cell |
| useRegex | <code>boolean</code> | Use regex to search the cell (find must be a regex string) |

<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..setSheet"></a>

#### SimpleGoogleSheets~setSheet(nameOrId) ⇒ <code>Promise</code>
Set the current sheet, referenced by the name or the id. Also updates the headers

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  
**Returns**: <code>Promise</code> - Undefined on success  

| Param | Type | Description |
| --- | --- | --- |
| nameOrId | <code>string</code> \| <code>number</code> | Integer id or string name of sheet |

<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..setSpreadsheet"></a>

#### SimpleGoogleSheets~setSpreadsheet(spreadsheetId) ⇒ <code>Promise</code>
Set the current spreadsheet

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  
**Returns**: <code>Promise</code> - Undefined on success  

| Param | Type | Description |
| --- | --- | --- |
| spreadsheetId | <code>string</code> | ID of spreadsheet to use |

<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..metadata"></a>

#### SimpleGoogleSheets~metadata() ⇒ <code>Object</code>
Get the current spreadsheet meta data

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  
**Returns**: <code>Object</code> - Metadata of current spreadsheet  
<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..getColumnName"></a>

#### SimpleGoogleSheets~getColumnName(columnIndex) ⇒
Get the name of the column (such as A, ABZ, or ZAD) from the index

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  
**Returns**: Column name (such as ABZ)  
**Throws**:

- Error if invalid index provided


| Param | Type | Description |
| --- | --- | --- |
| columnIndex | <code>number</code> | Get the column name from the index |

<a name="module_simplegooglesheetsjs.SimpleGoogleSheets..getRangeName"></a>

#### SimpleGoogleSheets~getRangeName(rowStart, columnStart, rowEnd, columnEnd) ⇒ <code>string</code>
Get the name of the range (in sheets format) from indexes

**Kind**: inner method of [<code>SimpleGoogleSheets</code>](#module_simplegooglesheetsjs.SimpleGoogleSheets)  
**Returns**: <code>string</code> - Range  
**Throws**:

- Error if invalid values provided


| Param | Type | Description |
| --- | --- | --- |
| rowStart | <code>number</code> | Starting row (and/or columnStart) |
| columnStart | <code>number</code> | Starting column |
| rowEnd | <code>number</code> | Ending row (and/or columnEnd) |
| columnEnd | <code>number</code> | Ending column |


## License
Copyright 2020 Alex Mous

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.