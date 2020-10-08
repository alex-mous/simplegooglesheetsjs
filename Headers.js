/**
* Headers
* @constructor
* @memberof module:simplegooglesheets
* @class Headers
* @classdesc Headers for SimpleGoogleSheets
*/
module.exports = class Headers {
    #headers; //(name -> column number)
    #invHeaders; //Inverted headers (column number -> name)

    constructor() {
        this.#headers =  {};
        this.#invHeaders = {};
    }

    /**
     * Add the header
     * 
     * @function Headers~setHeader
     * @param {string} name Name of header to add
     * @param {number} col Column number
     * @returns {boolean} Status (false if header name exists)
     */
    addHeader = (name, col) => {
        if (!this.#headers[name]) {
            this.#headers[name] = col;
            this.#invHeaders[col] = name;
            return true;
        }
        return false;
    }

    /**
     * Get the column number given by the name
     * 
     * @function Headers~getColumnNumber
     * @param {string} name Name of header
     * @returns {number} Column number, or undefined if none
     */
    getColumnNumber = (name) => {
        return this.#headers[name];
    }

    getHeaderName = (columnNumber) => {
        return this.#invHeaders[columnNumber];
    }

    /**
     * Get the headers as a key-value object of name-column number pairs
     * 
     * @function Headers~getHeaders
     * @returns {Object} Headers object
     */
    getHeaders = () => {
        return this.#headers;
    }

    /**
     * Get the headers as a key-value object of name-column number pairs
     * 
     * @function Headers~getHeadersByColumn
     * @returns {Array<string>} Array of names (index is column number)
     */
    getHeadersByColumn = () => {
        let arr = [];
        let i = 0;
        for (let index of Object.keys(this.#invHeaders).sort()) {
            while (i++ < index) {
                arr.push(undefined);
            }
            arr.push(this.#invHeaders[index]);
        }
        return arr;
    }

    get Headers() {
        return this.#headers;
    }
}