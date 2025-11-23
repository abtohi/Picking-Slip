/**
 * SheetHelper
 * Provides utility functions to interact with Google Sheets using dynamic column mapping.
 */
var SheetHelper = (function () {

    var columnCache = {};

    /**
     * Gets the column index (1-based) for a given header name.
     * Performs case-insensitive and trimmed comparison.
     * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 
     * @param {string} headerName 
     * @param {number} headerRowIndex (1-based, default 1)
     * @returns {number} 1-based column index, or -1 if not found.
     */
    function getColumnIndex(sheet, headerName, headerRowIndex) {
        headerRowIndex = headerRowIndex || 1;
        var cacheKey = sheet.getName() + "_" + headerName;

        if (columnCache[cacheKey]) {
            return columnCache[cacheKey];
        }

        var lastCol = sheet.getLastColumn();
        if (lastCol === 0) return -1;

        var headers = sheet.getRange(headerRowIndex, 1, 1, lastCol).getValues()[0];

        // Normalize target header
        var target = String(headerName).trim().toLowerCase();

        for (var i = 0; i < headers.length; i++) {
            var h = String(headers[i]).trim().toLowerCase();
            if (h === target) {
                columnCache[cacheKey] = i + 1; // Convert to 1-based
                return i + 1;
            }
        }

        return -1;
    }

    /**
     * Reads a row and returns an object mapping header names to values.
     * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 
     * @param {number} rowIndex 
     * @param {Object} schemaDefinition The schema definition for this sheet (from Schema.js)
     * @returns {Object} Data object
     */
    function getRowData(sheet, rowIndex, schemaDefinition) {
        var headers = schemaDefinition.HEADERS;
        var headerRow = schemaDefinition.TABLE_START_ROW ? schemaDefinition.TABLE_START_ROW - 1 : 1;

        var rowData = {};
        var lastCol = sheet.getLastColumn();
        var values = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];

        headers.forEach(function (header) {
            var colIdx = getColumnIndex(sheet, header, headerRow);
            if (colIdx !== -1) {
                rowData[header] = values[colIdx - 1];
            }
        });

        return rowData;
    }

    return {
        getColumnIndex: getColumnIndex,
        getRowData: getRowData
    };

})();
