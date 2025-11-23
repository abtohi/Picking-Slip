/**
 * ValidationService
 * Handles data validations and conditional formatting.
 */
var ValidationService = (function () {

    function setAllValidations(ss) {
        var sheets = [
            ss.getSheetByName(Config.SHEET_NAME.FORM_PICKING_SLIP_ADV),
            ss.getSheetByName(Config.SHEET_NAME.FORM_PICKING_SLIP_KYM)
        ];

        sheets.forEach(function (sheet) {
            if (!sheet) return;

            // 1. Requestor Validation
            var reqSheet = ss.getSheetByName(Config.SHEET_NAME.REQUESTOR);
            if (reqSheet) {
                var range = reqSheet.getRange("A2:A"); // Assuming list starts at A2
                var rule = SpreadsheetApp.newDataValidation()
                    .requireValueInRange(range, true)
                    .setAllowInvalid(false)
                    .setHelpText("Requestor tidak valid")
                    .build();

                // Apply to Requestor Column in Form
                // Need to find column index for Requestor
                var reqCol = SheetHelper.getColumnIndex(sheet, Schema.PICKING_SLIP.FORM_PICKING_SLIP.COLUMNS.REQUESTOR, Schema.PICKING_SLIP.FORM_PICKING_SLIP.TABLE_START_ROW - 1);
                if (reqCol > 0) {
                    sheet.getRange(Schema.PICKING_SLIP.FORM_PICKING_SLIP.TABLE_START_ROW, reqCol, sheet.getMaxRows() - Schema.PICKING_SLIP.FORM_PICKING_SLIP.TABLE_START_ROW + 1).setDataValidation(rule);
                }
            }

            // 2. Date Validation
            var dateRule = SpreadsheetApp.newDataValidation()
                .requireDate()
                .setAllowInvalid(false)
                .setHelpText("Masukkan tanggal yang valid")
                .build();
            sheet.getRange(Schema.PICKING_SLIP.FORM_PICKING_SLIP.CELLS.DATE).setDataValidation(dateRule);

            // 3. Kategori Validation
            var catRule = SpreadsheetApp.newDataValidation()
                .requireValueInList(["RM", "PM"])
                .setAllowInvalid(false)
                .setHelpText("Pilih RM atau PM")
                .build();
            sheet.getRange(Schema.PICKING_SLIP.FORM_PICKING_SLIP.CELLS.KATEGORI).setDataValidation(catRule);
        });
    }

    function setConditionalFormatting(ss) {
        var sheets = [
            ss.getSheetByName(Config.SHEET_NAME.FORM_PICKING_SLIP_ADV),
            ss.getSheetByName(Config.SHEET_NAME.FORM_PICKING_SLIP_KYM),
            ss.getSheetByName(Config.SHEET_NAME.PRINT)
        ];

        sheets.forEach(function (sheet) {
            if (!sheet) return;

            sheet.clearConditionalFormatRules();
            var rules = [];

            // Expiry Date Formatting
            // Find Expiry Date Column
            var headerRow = (sheet.getName() === Config.SHEET_NAME.PRINT) ? Schema.PICKING_SLIP.PRINT.TABLE_START_ROW - 1 : Schema.PICKING_SLIP.FORM_PICKING_SLIP.TABLE_START_ROW - 1;
            var colName = (sheet.getName() === Config.SHEET_NAME.PRINT) ? Schema.PICKING_SLIP.PRINT.COLUMNS.EXPIRY_DATE : Schema.PICKING_SLIP.FORM_PICKING_SLIP.COLUMNS.EXPIRY_DATE;

            var expCol = SheetHelper.getColumnIndex(sheet, colName, headerRow);

            if (expCol > 0) {
                var colLetter = getColLetter(expCol);
                var startRow = headerRow + 1;
                var range = sheet.getRange(startRow, expCol, sheet.getMaxRows() - startRow + 1);
                var firstCell = colLetter + startRow;

                // OK (> 90 days)
                rules.push(SpreadsheetApp.newConditionalFormatRule()
                    .whenFormulaSatisfied("=" + firstCell + "-TODAY()>90")
                    .setBackground("#b7e1cd")
                    .setRanges([range])
                    .build());

                // Warning (0-90 days)
                rules.push(SpreadsheetApp.newConditionalFormatRule()
                    .whenFormulaSatisfied("=AND(" + firstCell + "-TODAY()>0," + firstCell + "-TODAY()<=90)")
                    .setBackground("#ffe599")
                    .setRanges([range])
                    .build());

                // Expired (< 0 days)
                rules.push(SpreadsheetApp.newConditionalFormatRule()
                    .whenFormulaSatisfied("=AND(" + firstCell + "-TODAY()<0," + firstCell + "<>\"\")")
                    .setBackground("#e06666")
                    .setRanges([range])
                    .build());
            }

            sheet.setConditionalFormatRules(rules);
        });
    }

    function getColLetter(colIndex) {
        var temp, letter = '';
        while (colIndex > 0) {
            temp = (colIndex - 1) % 26;
            letter = String.fromCharCode(temp + 65) + letter;
            colIndex = (colIndex - temp - 1) / 26;
        }
        return letter;
    }

    return {
        setAllValidations: setAllValidations,
        setConditionalFormatting: setConditionalFormatting
    };

})();
