/**
 * PickingSlipService
 * Handles business logic for Picking Slip forms (Save, Clear, Validate).
 * Restored exact logic from original Form Picking Slip.js
 */
var PickingSlipService = (function () {

    var SCHEMA = Schema.PICKING_SLIP;

    function clearForm(sheet) {
        var startRow = SCHEMA.FORM_PICKING_SLIP.TABLE_START_ROW;
        var lastRow = sheet.getLastRow();

        // Clear Data Range
        // Original used config.database_structure.picking_slip.form_picking_slip.data (A8:M)
        // We should clear based on actual content to be safe, or use the fixed range if preferred.
        // Let's use dynamic range starting from startRow
        if (lastRow >= startRow) {
            sheet.getRange(startRow, 1, lastRow - startRow + 1, 13).clearContent(); // A to M

            // Clear Validations for Location and Area Gudang
            var colIndices = getColIndices(sheet);
            if (colIndices.LOCATION > 0) sheet.getRange(startRow, colIndices.LOCATION, lastRow - startRow + 1).clearDataValidations();
            if (colIndices.AREA_GUDANG > 0) sheet.getRange(startRow, colIndices.AREA_GUDANG, lastRow - startRow + 1).clearDataValidations();
        }

        // Clear Header Cells
        sheet.getRange(SCHEMA.FORM_PICKING_SLIP.CELLS.DATE).clearContent();
        sheet.getRange(SCHEMA.FORM_PICKING_SLIP.CELLS.PICKER).clearContent();
        sheet.getRange(SCHEMA.FORM_PICKING_SLIP.CELLS.CHECKER).clearContent();
        sheet.getRange(SCHEMA.FORM_PICKING_SLIP.CELLS.KATEGORI).clearContent();

        // Clear Print Sheet
        var ss = sheet.getParent();
        var printSheet = ss.getSheetByName(Config.SHEET_NAME.PRINT);
        if (printSheet) {
            var printStartRow = SCHEMA.PRINT.TABLE_START_ROW;
            var printLastRow = printSheet.getLastRow();
            if (printLastRow >= printStartRow) {
                printSheet.getRange(printStartRow, 2, printLastRow - printStartRow + 1, 15).clearContent(); // B to O
            }
            printSheet.getRange(SCHEMA.PRINT.CELLS.DATE).clearContent();
            printSheet.getRange(SCHEMA.PRINT.CELLS.PICKER).clearContent();
        }
    }

    function saveData(sheet) {
        try {
            var date = sheet.getRange(SCHEMA.FORM_PICKING_SLIP.CELLS.DATE).getValue();
            var picker = sheet.getRange(SCHEMA.FORM_PICKING_SLIP.CELLS.PICKER).getValue();
            var checker = sheet.getRange(SCHEMA.FORM_PICKING_SLIP.CELLS.CHECKER).getValue();
            var kategori = sheet.getRange(SCHEMA.FORM_PICKING_SLIP.CELLS.KATEGORI).getValue();

            // Validation: Header Fields
            if (date === "") return Browser.msgBox("Silahkan masukkan tanggal terlebih dahulu pada cell " + SCHEMA.FORM_PICKING_SLIP.CELLS.DATE);
            if (kategori === "") return Browser.msgBox("Silahkan pilih kategori pada cell " + SCHEMA.FORM_PICKING_SLIP.CELLS.KATEGORI);
            if (picker === "") return Browser.msgBox("Silahkan masukkan picker pada cell " + SCHEMA.FORM_PICKING_SLIP.CELLS.PICKER);
            if (checker === "") return Browser.msgBox("Silahkan masukkan checker pada cell " + SCHEMA.FORM_PICKING_SLIP.CELLS.CHECKER);

            var startRow = SCHEMA.FORM_PICKING_SLIP.TABLE_START_ROW;
            var lastRow = sheet.getLastRow();
            if (lastRow < startRow) return;

            var colIndices = getColIndices(sheet);
            var dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn());
            var values = dataRange.getValues();

            // Validation: Rows
            for (var i = 0; i < values.length; i++) {
                var row = values[i];
                var itemCode = row[colIndices.ITEM_CODE - 1];

                // Check if row has partial data but missing Item Code
                var hasOtherData = row[colIndices.PO_NUMBER - 1] || row[colIndices.VENDOR - 1];
                if (itemCode === "" && hasOtherData) {
                    return Browser.msgBox("Silahkan lengkapi item code pada baris " + (startRow + i));
                }

                if (itemCode !== "") {
                    if (row[colIndices.REQUESTOR - 1] === "") return Browser.msgBox("Silahkan lengkapi requestor pada baris " + (startRow + i));
                    if (row[colIndices.LOCATION - 1] === "") return Browser.msgBox("Silahkan lengkapi location pada baris " + (startRow + i));
                    if (row[colIndices.QTY_REQ - 1] === "") return Browser.msgBox("Silahkan lengkapi quantity request pada baris " + (startRow + i));
                    if (Number(row[colIndices.QTY_REQ - 1]) > Number(row[colIndices.QTY_STOCK - 1])) {
                        return Browser.msgBox("Qty req tidak dapat melebihi qty stock pada baris " + (startRow + i));
                    }
                }
            }

            var conf = Browser.msgBox("Konfirmasi Data Benar", "Apakah data Anda sudah benar?", Browser.Buttons.YES_NO);
            if (conf.toString().toUpperCase() !== "YES") return;

            // --- WAVE CALCULATION ---
            var ss = sheet.getParent();
            var historySheet = ss.getSheetByName(Config.SHEET_NAME.HISTORY);
            var historyData = historySheet.getDataRange().getValues();
            var historyCols = getHistoryIndices(historySheet);

            // Filter history by date
            var formattedDate = Utils.formatDefault(date);
            var currData = historyData.filter(function (row) {
                return Utils.formatDefault(row[historyCols.DATE - 1]) === formattedDate;
            });

            // Get unique timestamps (waves)
            var timestamps = currData.map(function (row) { return row[historyCols.TIMESTAMP - 1]; });
            // Simple unique count
            var uniqueWaves = [];
            timestamps.forEach(function (ts) {
                var t = new Date(ts).getTime();
                if (uniqueWaves.indexOf(t) === -1) uniqueWaves.push(t);
            });

            var wave = uniqueWaves.length + 1;

            // --- PREPARE DATA ---
            var rowsToSave = [];
            var rowsToPrint = [];
            var timestamp = new Date();

            for (var i = 0; i < values.length; i++) {
                var row = values[i];
                var itemCode = row[colIndices.ITEM_CODE - 1];
                if (itemCode === "") continue;

                var qtyReq = row[colIndices.QTY_REQ - 1];
                var qtyStock = row[colIndices.QTY_STOCK - 1];
                var uom = row[colIndices.UOM - 1];

                // Convert KG to Grams if needed (Original logic)
                var qtyReqGram = (uom === "KG") ? Number(qtyReq) * 1000 : qtyReq;
                var qtyStockGram = (uom === "KG") ? Number(qtyStock) * 1000 : qtyStock;

                // Construct History Row (Order matters based on Schema/Original)
                // Original: ID, Timestamp, Date, Wave, Kategori, PO, Vendor, ItemCode, ItemName, UOM, Requestor, Batch, Area, Loc, Exp, Stock, StockG, Req, ReqG, Picked, PickedG, Sisa, SisaG, Picker, Checker
                var newRow = [];
                // We need to map to the exact column order of History Sheet.
                // Let's assume Schema.HISTORY.HEADERS order is correct and matches original.
                // But to be safe, let's build an object and map it.

                var hObj = {};
                hObj["ID"] = Utils.generateRandomStringLC(8);
                hObj["Timestamp"] = timestamp;
                hObj["Date"] = formattedDate;
                hObj["Wave"] = wave;
                hObj["Kategori"] = kategori;
                hObj["PO Number"] = row[colIndices.PO_NUMBER - 1];
                hObj["Vendor"] = row[colIndices.VENDOR - 1];
                hObj["Item Code"] = itemCode;
                hObj["Item Name"] = row[colIndices.ITEM_NAME - 1];
                hObj["UOM"] = uom;
                hObj["Requestor"] = row[colIndices.REQUESTOR - 1];
                hObj["Batch/Lot"] = row[colIndices.BATCH_LOT - 1];
                hObj["Area Gudang"] = row[colIndices.AREA_GUDANG - 1];
                hObj["Location"] = row[colIndices.LOCATION - 1];
                hObj["Expiry Date"] = row[colIndices.EXPIRY_DATE - 1] ? Utils.formatDefault(row[colIndices.EXPIRY_DATE - 1]) : "";
                hObj["Qty Stock"] = qtyStock;
                hObj["Qty Stock (g)"] = qtyStockGram;
                hObj["Qty Req"] = qtyReq;
                hObj["Qty Req (g)"] = qtyReqGram;
                hObj["Picker"] = picker;
                hObj["Checker"] = checker;

                // Map to array based on History Headers
                var hHeaders = Schema.PICKING_SLIP.HISTORY.HEADERS;
                var hRowArr = hHeaders.map(function (h) { return hObj[h] || ""; });
                rowsToSave.push(hRowArr);

                // Construct Print Row
                // No, PO, Vendor, Code, Name, QtyReq, Picked(Empty), UOM, Requestor, Batch, Area, Loc, Exp, Sisa(Empty)
                rowsToPrint.push([
                    row[colIndices.NO - 1],
                    row[colIndices.PO_NUMBER - 1],
                    row[colIndices.VENDOR - 1],
                    itemCode,
                    row[colIndices.ITEM_NAME - 1],
                    qtyReq,
                    "",
                    uom,
                    row[colIndices.REQUESTOR - 1],
                    row[colIndices.BATCH_LOT - 1],
                    row[colIndices.AREA_GUDANG - 1],
                    row[colIndices.LOCATION - 1],
                    row[colIndices.EXPIRY_DATE - 1] ? Utils.formatDefault(row[colIndices.EXPIRY_DATE - 1]) : "",
                    ""
                ]);
            }

            // --- WRITE TO HISTORY ---
            if (rowsToSave.length > 0) {
                historySheet.getRange(historySheet.getLastRow() + 1, 1, rowsToSave.length, rowsToSave[0].length).setValues(rowsToSave);
            }

            // --- WRITE TO PRINT ---
            var printSheet = ss.getSheetByName(Config.SHEET_NAME.PRINT);
            if (printSheet && rowsToPrint.length > 0) {
                var printStartRow = SCHEMA.PRINT.TABLE_START_ROW;

                // Clear old
                printSheet.getRange(printStartRow, 2, printSheet.getMaxRows() - printStartRow + 1, 15).clearContent();

                // Write new
                var targetRange = printSheet.getRange(printStartRow, 2, rowsToPrint.length, rowsToPrint[0].length);
                targetRange.setValues(rowsToPrint);

                // Formatting (Borders)
                targetRange.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.DOTTED);
                targetRange.setFontWeight('normal');

                // Center Align specific columns (Item Code, Batch, UOM) - Indices relative to range (start col 2)
                // Item Code is col 4 (index 3 in array, col 5 in sheet) -> relative index 4
                // Let's just use the range offset
                // Item Code (Col 4 in array -> Col 5 in sheet)
                printSheet.getRange(printStartRow, 5, rowsToPrint.length, 1).setHorizontalAlignment('center');
                // Batch (Col 9 in array -> Col 10 in sheet)
                printSheet.getRange(printStartRow, 11, rowsToPrint.length, 1).setHorizontalAlignment('center'); // Batch is 10th in array? No, 9th index -> 10th col. Wait.
                // Array: No, PO, Vendor, Code, Name, QtyReq, Picked, UOM, Requestor, Batch
                // Idx:   0   1   2       3     4     5       6       7    8          9
                // Sheet: B   C   D       E     F     G       H       I    J          K
                // Batch is K -> 11. 
                printSheet.getRange(printStartRow, 11, rowsToPrint.length, 1).setHorizontalAlignment('center');
                // UOM is I -> 9
                printSheet.getRange(printStartRow, 9, rowsToPrint.length, 1).setHorizontalAlignment('center');

                // Header Info
                var dateFormatted = (wave > 0) ? formattedDate + " - " + wave : formattedDate;
                printSheet.getRange(SCHEMA.PRINT.CELLS.DATE).setValue(dateFormatted);
                printSheet.getRange(SCHEMA.PRINT.CELLS.PICKER).setValue(picker);

                // Signatures
                var lastRowPrint = printStartRow + rowsToPrint.length;
                printSheet.getRange(lastRowPrint + 2, 5).setValue("Picker").setFontWeight("bold").setHorizontalAlignment("center");
                printSheet.getRange(lastRowPrint + 5, 5).setValue("( " + picker + " )").setFontWeight("bold").setHorizontalAlignment("center");

                printSheet.getRange(lastRowPrint + 2, 11).setValue("Checker").setFontWeight("bold").setHorizontalAlignment("center");
                printSheet.getRange(lastRowPrint + 5, 11).setValue("( " + checker + " )").setFontWeight("bold").setHorizontalAlignment("center");

                // Activate Print Sheet
                printSheet.activate();

                // Update Data (Sync)
                SyncService.updateData(false);
            }

            var confClear = Browser.msgBox("Konfirmasi Hapus Data", "Data Berhasil Tersimpan\\n\\nApakah Anda ingin menghapus data pada form?", Browser.Buttons.YES_NO);
            if (confClear.toString().toUpperCase() === "YES") {
                clearForm(sheet);
            }

        } catch (e) {
            Utils.logError(Config.SSID.ERROR_LOG, Config.SHEET_NAME.ERROR, e);
            Browser.msgBox("Error: " + e.message);
        }
    }

    function getColIndices(sheet) {
        var headerRow = SCHEMA.FORM_PICKING_SLIP.TABLE_START_ROW - 1;
        var indices = {};
        for (var key in SCHEMA.FORM_PICKING_SLIP.COLUMNS) {
            indices[key] = SheetHelper.getColumnIndex(sheet, SCHEMA.FORM_PICKING_SLIP.COLUMNS[key], headerRow);
        }
        return indices;
    }

    function getHistoryIndices(sheet) {
        var headerRow = 2; // Assumption
        var indices = {};
        // Map Schema keys to indices
        for (var key in SCHEMA.HISTORY.COLUMNS) {
            indices[key] = SheetHelper.getColumnIndex(sheet, SCHEMA.HISTORY.COLUMNS[key], headerRow);
        }
        return indices;
    }

    return {
        saveData: saveData,
        clearForm: clearForm
    };

})();
