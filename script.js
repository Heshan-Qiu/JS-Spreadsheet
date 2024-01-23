const SPREADSHEET_ROWS = 101;
const SPREADSHEET_COLUMNS = 101;

document.addEventListener("DOMContentLoaded", function () {
    const spreadsheetContainer = document.getElementById(
        "spreadsheetContainer"
    );
    const grid = new Grid(SPREADSHEET_ROWS, SPREADSHEET_COLUMNS);

    renderSpreadsheet(grid, spreadsheetContainer);
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
    // inputElement.value = cell.value;
    inputElement.value = cell.row + "," + cell.column;
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

    inputElement.addEventListener("input", function (event) {
        cell.value = event.target.value;
        console.log(
            `Cell ${cell.row},${cell.column} value changed to ${cell.value}`,
            cell
        );
    });

    cellElement.appendChild(inputElement);

    return cellElement;
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
        this.value = value;
        this.formula = formula;
        this.isHeader = isHeader;
        this.row = null;
        this.columns = null;
    }
}
