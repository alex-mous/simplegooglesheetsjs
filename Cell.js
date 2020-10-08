class Cell {
    #value = "";

    constructor(value) {
        this.#value = value;
    }

    get getValue() {
        if (this.#value !== undefined) {
            return this.#value;
        } else {
            throw new Error("Cell value not set");
        }
    }

    setValue(value) {
        this.#value = value;
    }
}

module.exports = Cell;