/*
Paste to Visible Cells

This is the Office Script version of https://www.reddit.com/r/excel/comments/wkrnfv/comment/l5gdqye

Usage: In a worksheet,
1) Select the block of cells to copy.
2) Run this script by going to Automate menu > finding and clicking the saved
script. A prompt appears, then input the starting cell to paste.

Distributed under the MIT software license. See the accompanying
file LICENSE or https://opensource.org/license/mit/.
*/
function main(workbook: ExcelScript.Workbook, startCell: string) {
    let selectedSheet = workbook.getActiveWorksheet();
    let toCopy = workbook.getSelectedRange();

    let chkCell = selectedSheet.getRange(startCell);
    let totalRow = toCopy.getRowCount(), totalCol = toCopy.getColumnCount();
    let chkRow = 1, chkCol = 1;

    while (chkRow < totalRow) {
        chkCell = chkCell.getOffsetRange(1, 0);
        if (!chkCell.getHidden()) chkRow++;
    }
    while (chkCol < totalCol) {
        chkCell = chkCell.getOffsetRange(0, 1);
        if (!chkCell.getHidden()) chkCol++;
    }
    let toPaste = selectedSheet.getRange(startCell + ':' + chkCell.getAddress())
        .getVisibleView();
    toPaste.setValues(toCopy.getValues());
}