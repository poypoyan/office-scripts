/*
Indicate Duplicate from Color-highlighted Cells

Usage: In a worksheet,
Run this script by going to Automate menu > new script/finding and clicking
the saved script. A prompt appears, then input the column to check, start
row, end row, and the column to indicate.

Distributed under the MIT software license. See the accompanying
file LICENSE or https://opensource.org/license/mit/.
*/

function main(workbook: ExcelScript.Workbook, checkcol: string, startrow: string, endrow: string, writecol: string) {
    let selectedSheet = workbook.getActiveWorksheet();

    let checkColRan = selectedSheet.getRange(checkcol + startrow + ":" + checkcol + endrow);
    let checkColVals = checkColRan.getValues();
    let distinctIdxs = new Set<number>();
    let distinctVals = new Set<string>();

    // get distinct values from highlighted cells
    for (let row = 0; row < checkColVals.length; row++) {
        const cellColor = checkColRan.getCell(row, 0).getFormat().getFill().getColor();
        const cellValue = String(checkColVals[row][0]);
        if (cellColor && cellColor !== "#FFFFFF" && cellColor !== "") {
            if (cellValue !== "" && !distinctVals.has(cellValue)) {
                distinctIdxs.add(row);
                distinctVals.add(cellValue);
            }
        }
    }

    // check whole range for duplicate
    const duplStr = "duplicate";
    let writeColRan = selectedSheet.getRange(writecol + startrow + ":" + writecol + endrow);
    for (let row = 0; row < checkColVals.length; row++) {
        const cellValue = String(checkColVals[row][0]);
        if (!distinctIdxs.has(row) && distinctVals.has(cellValue)) {
            writeColRan.getCell(row, 0).setValue(duplStr);
        }
    }
}