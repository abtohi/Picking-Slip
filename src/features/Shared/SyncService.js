/**
 * SyncService
 * Handles synchronization with external spreadsheets (Inventory, Movement).
 */
var SyncService = (function () {

    function updateData(showUi) {
        try {
            var ss = SpreadsheetApp.getActiveSpreadsheet();
            var inventorySS = SpreadsheetApp.openById(Config.SSID.INVENTORY_STOCK);

            // 1. Sync Inventory Data
            var invDataSheet = inventorySS.getSheetByName(Config.SHEET_NAME.DATA);
            // Assuming Inventory Sheet has headers at row 3, data at row 4 (based on original code)
            // Original code used hardcoded indices from config.database_structure.inventory_stock.data
            // We should try to be dynamic if possible, or stick to a known schema for the external sheet.
            // Since I don't have the external sheet schema in my new Schema.js yet, I'll assume a standard structure or use the indices from the old code as a guide but implemented cleaner.

            // Let's read the whole data range
            var invValues = invDataSheet.getDataRange().getValues();
            // Filter and Map
            // Original: item_code != '' && location != ''
            // Map to: Item Code, Desc, PO, Type, Supplier, UOM, Lot Adev, Location, Expiry, Qty Update

            // We need to find columns in the external sheet.
            // Let's use SheetHelper for the external sheet too.
            var invHeadersRow = 3; // Assumption based on original code
            var invColIdx = {
                ITEM_CODE: SheetHelper.getColumnIndex(invDataSheet, "Item Code", invHeadersRow),
                DESC: SheetHelper.getColumnIndex(invDataSheet, "Material Description", invHeadersRow),
                PO: SheetHelper.getColumnIndex(invDataSheet, "PO Number", invHeadersRow),
                TYPE: SheetHelper.getColumnIndex(invDataSheet, "Type", invHeadersRow),
                SUPPLIER: SheetHelper.getColumnIndex(invDataSheet, "Supplier/Consumers", invHeadersRow),
                UOM: SheetHelper.getColumnIndex(invDataSheet, "UOM", invHeadersRow),
                LOT: SheetHelper.getColumnIndex(invDataSheet, "Lot Number ADEV", invHeadersRow),
                LOC: SheetHelper.getColumnIndex(invDataSheet, "Location", invHeadersRow),
                EXP: SheetHelper.getColumnIndex(invDataSheet, "Expiry Date", invHeadersRow),
                QTY: SheetHelper.getColumnIndex(invDataSheet, "Quantity Update", invHeadersRow)
            };

            var processedData = [];
            for (var i = invHeadersRow; i < invValues.length; i++) {
                var row = invValues[i];
                var itemCode = row[invColIdx.ITEM_CODE - 1];
                var loc = row[invColIdx.LOC - 1];

                if (itemCode && loc) {
                    processedData.push([
                        itemCode,
                        row[invColIdx.DESC - 1],
                        row[invColIdx.PO - 1],
                        row[invColIdx.TYPE - 1],
                        row[invColIdx.SUPPLIER - 1],
                        row[invColIdx.UOM - 1],
                        row[invColIdx.LOT - 1],
                        loc,
                        row[invColIdx.EXP - 1],
                        row[invColIdx.QTY - 1]
                    ]);
                }
            }

            // Split into PM and RM
            var pmData = processedData.filter(function (r) { return r[3] === "PM"; }); // Index 3 is Type
            var rmData = processedData.filter(function (r) { return r[3] === "RM"; });

            // Update Local Sheets
            updateSheet(ss, Config.SHEET_NAME.MASTER_INVENTORY, processedData);
            updateSheet(ss, Config.SHEET_NAME.PM_LIST, pmData);
            updateSheet(ss, Config.SHEET_NAME.RM_LIST, rmData);

            // 2. Sync Detail Lokasi
            syncDetailLokasi(ss);

            // 3. Refresh Validations & Formatting
            ValidationService.setAllValidations(ss);
            ValidationService.setConditionalFormatting(ss);

            if (showUi) {
                Browser.msgBox("Data berhasil diupdate");
            }

        } catch (e) {
            Utils.logError(Config.SSID.ERROR_LOG, Config.SHEET_NAME.ERROR, e);
            Browser.msgBox("Error updating data: " + e.message);
        }
    }

    function updateSheet(ss, sheetName, data) {
        var sheet = ss.getSheetByName(sheetName);
        if (!sheet) return;

        // Clear old data (assuming header is row 1, data starts row 2)
        var lastRow = sheet.getLastRow();
        if (lastRow > 1) {
            sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
        }

        if (data.length > 0) {
            sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
        }
    }

    function syncDetailLokasi(ss) {
        var movementSS = SpreadsheetApp.openById(Config.SSID.MOVEMENT_BARANG);
        var movSheet = movementSS.getSheetByName(Config.SHEET_NAME.DETAIL_LOKASI);

        // Read from Movement
        // Columns: Jenis Lokasi, Sloc
        var movValues = movSheet.getDataRange().getValues();
        var headerRow = 1; // Assumption
        var colIdx = {
            JENIS: SheetHelper.getColumnIndex(movSheet, "Jenis Lokasi", headerRow),
            SLOC: SheetHelper.getColumnIndex(movSheet, "Sloc", headerRow)
        };

        var lokasiData = [];
        for (var i = headerRow; i < movValues.length; i++) {
            var row = movValues[i];
            var jenis = row[colIdx.JENIS - 1];
            if (jenis) {
                var gudang = (jenis === "Gudang RM Factory") ? "ADV" : "KYM";
                lokasiData.push([gudang, row[colIdx.SLOC - 1]]);
            }
        }

        // Update Local
        updateSheet(ss, Config.SHEET_NAME.DETAIL_LOKASI, lokasiData);
    }

    return {
        updateData: updateData
    };

})();
