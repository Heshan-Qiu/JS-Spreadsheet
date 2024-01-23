const SPREADSHEET_ROWS = 101;
const SPREADSHEET_COLUMNS = 101;

document.addEventListener("DOMContentLoaded", function () {
    const spreadsheetContainer = document.getElementById(
        "spreadsheetContainer"
    );
    const grid = new Grid(SPREADSHEET_ROWS, SPREADSHEET_COLUMNS);

    renderSpreadsheet(grid, spreadsheetContainer);

    initRefreshButton(grid, spreadsheetContainer);
    initCheckboxes();
});

function renderSpreadsheet(grid, container) {
    if (!grid instanceof Grid || !container instanceof HTMLElement) {
        throw new Error("Invalid arguments");
    }

    for (let i = 0; i < grid.cells.length; i++) {
        spreadsheetContainer.appendChild(renderRow(grid.cells[i]));
    }
}

function renderRow(rowCells) {
    if (!Array.isArray(rowCells)) {
        throw new Error("Invalid arguments");
    }

    const rowElement = document.createElement("div");
    rowElement.classList.add("row");

    for (let i = 0; i < rowCells.length; i++) {
        rowElement.appendChild(renderCell(rowCells[i]));
    }

    return rowElement;
}

function renderCell(cell) {
    if (!cell instanceof Cell) {
        throw new Error("Invalid arguments");
    }

    const cellElement = document.createElement("div");
    cellElement.classList.add("cell");

    const inputElement = document.createElement("input");
    inputElement.type = "text";
    inputElement.value = cell.value;
    // inputElement.value = cell.row + "," + cell.column;
    inputElement.id = indexToReference(cell.row, cell.column);

    if (cell.isHeader) {
        cellElement.classList.add("header");
        inputElement.value =
            cell.row === 0
                ? cell.column === 0
                    ? ""
                    : indexToLetters(cell.column)
                : cell.row;
        inputElement.disabled = true;
    }

    inputElement.addEventListener("change", function (event) {
        cell.updateValue(event.target.value);
        console.log(
            `Cell ${cell.row},${cell.column} value changed to ${cell.value}`,
            cell
        );
    });

    inputElement.addEventListener("focus", function (event) {
        const cellSelectedInputElement =
            document.getElementById("cellSelectedInput");
        const cellSelectedValueInputElement = document.getElementById(
            "cellSelectedValueInput"
        );

        cellSelectedInputElement.value = event.target.id;
        cellSelectedValueInputElement.value = cell.formula
            ? cell.formula
            : cell.value;
    });

    cellElement.appendChild(inputElement);

    cell.onChange = () => {
        inputElement.value = cell.value;
        console.log(`Cell ${cell.row},${cell.column} value changed`, cell);
    };

    return cellElement;
}

function initRefreshButton(grid, spreadsheetContainer) {
    if (!grid instanceof Grid && !spreadsheetContainer instanceof HTMLElement) {
        throw new Error("Invalid arguments");
    }

    const refreshButton = document.getElementById("refreshButton");
    refreshButton.addEventListener("click", function (event) {
        spreadsheetContainer.innerHTML = "";
        // this is too fast to see the change, so we add a delay
        // renderSpreadsheet(grid, spreadsheetContainer);
        setTimeout(function () {
            renderSpreadsheet(grid, spreadsheetContainer);
        }, 300);
    });
}

function initCheckboxes() {
    const boldCheckbox = document.getElementById("boldCheckbox");
    boldCheckbox.addEventListener("change", (event) => {
        const refElement = document.getElementById("cellSelectedInput");
        const cell = document.getElementById(refElement.value);

        if (event.target.checked) {
            cell.style.fontWeight = "bold";
        } else {
            cell.style.fontWeight = "normal";
        }
    });

    const italicCheckbox = document.getElementById("italicCheckbox");
    italicCheckbox.addEventListener("change", (event) => {
        const refElement = document.getElementById("cellSelectedInput");
        const cell = document.getElementById(refElement.value);

        if (event.target.checked) {
            cell.style.fontStyle = "italic";
        } else {
            cell.style.fontStyle = "normal";
        }
    });

    const underlineCheckbox = document.getElementById("underlineCheckbox");
    underlineCheckbox.addEventListener("change", (event) => {
        const refElement = document.getElementById("cellSelectedInput");
        const cell = document.getElementById(refElement.value);

        if (event.target.checked) {
            cell.style.textDecoration = "underline";
        } else {
            cell.style.textDecoration = "none";
        }
    });
}

function indexToLetters(index) {
    if (!Number.isInteger(index) || index < 0) {
        throw new Error("Invalid arguments");
    }

    let result = "";
    while (index > 0) {
        result = String.fromCharCode(65 + ((index - 1) % 26)) + result;
        index = Math.floor((index - 1) / 26);
    }
    return result;
}

function lettersToIndex(letters) {
    if (typeof letters !== "string") {
        throw new Error("Invalid arguments");
    }

    let result = 0;
    for (let i = 0; i < letters.length; i++) {
        result *= 26;
        result += letters.charCodeAt(i) - 64;
    }
    return result;
}

function indexToReference(row, column) {
    if (
        !Number.isInteger(row) ||
        row < 0 ||
        !Number.isInteger(column) ||
        column < 0
    ) {
        throw new Error("Invalid arguments");
    }

    return indexToLetters(column) + row;
}

function referenceToIndex(reference) {
    if (typeof reference !== "string") {
        throw new Error("Invalid arguments");
    }

    const [col, row] = reference.match(/[0-9]+|[A-Z]+/g);
    return [parseInt(row), lettersToIndex(col)];
}

function isFormula(value) {
    if (value === null || value === undefined) {
        return false;
    }

    return value.startsWith("=");
}

function getAllDependentCellsFromFormula(formula, grid) {
    if (!grid instanceof Grid) {
        throw new Error("Invalid arguments");
    }

    if (this.isFormula(formula)) {
        // Remove '=' and all whitespace from the formula
        const expression = formula.slice(1).replace(/\s/g, "");
        // Split the expression based on operators (+, -, *, /), exclude the operators
        let parts = expression.split(/([+*/-])/).filter((part) => {
            return !part.match(/[+*/-]/);
        });

        let cells = [];
        for (const part of parts) {
            if (part.startsWith("SUM")) {
                const [startRef, endRef] = part
                    .slice(4, -1)
                    .split(":")
                    .map((ref) => referenceToIndex(ref));
                for (let i = startRef[0]; i <= endRef[0]; i++) {
                    for (let j = startRef[1]; j <= endRef[1]; j++) {
                        cells.push(grid.getCell(i, j));
                    }
                }
            } else if (part.match(/[A-Z]+[0-9]+/)) {
                cells.push(grid.getCell(...referenceToIndex(part)));
            }
        }

        return cells;
    }

    return [];
}

function isCircularReference(cell, formula) {
    if (!cell instanceof Cell || typeof formula !== "string") {
        throw new Error("Invalid arguments");
    }

    // Get all cell references in the formula
    const matches = getAllDependentCellsFromFormula(formula, cell.grid);
    if (matches.includes(cell)) {
        return true;
    }

    // Check if any of the referenced cells have circular references
    for (const match of matches) {
        if (match.formula && isCircularReference(cell, match.formula)) {
            return true;
        }
    }

    return false;
}

class Grid {
    constructor(rows, columns) {
        this.cells = new Array(rows);

        for (let i = 0; i < rows; i++) {
            this.cells[i] = new Array(columns);

            for (let j = 0; j < columns; j++) {
                this.cells[i][j] = new Cell(this);
                this.cells[i][j].row = i;
                this.cells[i][j].column = j;

                if (i === 0 || j === 0) {
                    this.cells[i][j].isHeader = true;
                }
            }
        }
    }

    getCell(row, column) {
        return this.cells[row][column];
    }
}

class Cell {
    constructor(grid, value = "", formula = "", isHeader = false) {
        this.grid = grid;
        this._value = value;
        this.formula = formula;
        this.isHeader = isHeader;
        this.row = null;
        this.columns = null;
        this.onChange = null; // call back function when value changes
        this.dependencies = []; // list of cells that depend on this cell
    }

    set value(value) {
        this._value = value;
        if (this.onChange) {
            this.onChange();
        }
    }

    get value() {
        return this._value;
    }

    addDependency(cell) {
        if (!cell instanceof Cell) {
            throw new Error("Invalid arguments");
        }

        if (!this.dependencies.includes(cell)) {
            this.dependencies.push(cell);
        }
    }

    removeDependency(cell) {
        if (!cell instanceof Cell) {
            throw new Error("Invalid arguments");
        }

        const index = this.dependencies.indexOf(cell);
        if (index !== -1) {
            this.dependencies.splice(index, 1);
        }
    }

    updateDependencies() {
        this.dependencies.forEach((cell) => {
            cell.computeFormula();
            cell.updateDependencies();
        });
    }

    /**
     * Updates the value of the cell.
     * @param {*} value - The new value to be assigned.
     * steps:
     * 1. check the value is a formula or not, if not:
     *   a. assign the value to the cell value
     *   b. remove all dependencies this cell depends on
     *   c. reset the formula to empty string
     *   d. update all cells that depend on this cell
     * 2. if the value is a formula and it is a circular reference, return an error value "#REF!"
     * 3. if the value is a formula, check if it is a circular reference, if not:
     *   a. remove all dependencies this cell depends on
     *   b. assign the formula to the cell formula
     *   c. add all dependencies this cell depends on
     *   d. compute the value of the formula
     *   e. update all cells that depend on this cell
     */
    updateValue(value) {
        if (!isFormula(value)) {
            this.value = value;
            clearDependenciesForCell(this, this.grid);
            this.formula = "";
            this.updateDependencies();
            return this.value;
        }

        if (isCircularReference(this, value)) {
            this.value = "#REF!";
            return this.value;
        }

        clearDependenciesForCell(this, this.grid);
        this.formula = value;
        addDependenciesForCell(this, this.grid);
        this.computeFormula();
        this.updateDependencies();
    }

    computeFormula() {
        console.log(
            `Computing formula for cell ${this.row}, ${this.column}`,
            this
        );

        // Remove '=' and all whitespace from the formula
        const expression = this.formula.slice(1).replace(/\s/g, "");

        // Split the expression based on operators (+, -, *, /)
        let parts = expression.split(/([+*/-])/);

        // Replace cell references with their values
        parts = parts.map((part) => {
            if (part.startsWith("SUM")) {
                const [startRef, endRef] = part
                    .slice(4, -1)
                    .split(":")
                    .map((ref) => referenceToIndex(ref));
                let sum = 0;
                for (let i = startRef[0]; i <= endRef[0]; i++) {
                    for (let j = startRef[1]; j <= endRef[1]; j++) {
                        this.grid.getCell(i, j)
                            ? (sum += Number(this.grid.getCell(i, j).value))
                            : (sum += 0);
                    }
                }
                return sum;
            }

            if (part.match(/[A-Z]+[0-9]+/)) {
                const refCell = this.grid.getCell(...referenceToIndex(part));
                return refCell ? refCell.value : 0;
            }

            return part;
        });

        // Rebuild the formula using the computed parts
        const newFormula = parts.join("");

        // Evaluate the formula and assign the result to the cell
        // Todo: the eval function is dangerous, we should use a library like math.js
        try {
            this.value = eval(newFormula);
        } catch (e) {
            this.value = "#VALUE!";
        }
    }
}

function addDependenciesForCell(cell, grid) {
    if (!cell instanceof Cell || !grid instanceof Grid) {
        throw new Error("Invalid arguments");
    }

    for (const depentCell of getAllDependentCellsFromFormula(
        cell.formula,
        grid
    )) {
        depentCell.addDependency(cell);
    }
}

function clearDependenciesForCell(cell, grid) {
    if (!cell instanceof Cell || !grid instanceof Grid) {
        throw new Error("Invalid arguments");
    }

    for (const depentCell of getAllDependentCellsFromFormula(
        cell.formula,
        grid
    )) {
        depentCell.removeDependency(cell);
    }
}
