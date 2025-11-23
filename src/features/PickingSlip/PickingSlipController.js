/**
 * PickingSlipController
 * Handles onEdit events for Picking Slip sheets.
 * Restored exact logic from original onEdit.js
 * Added safety checks for column indices.
 */
var PickingSlipController = (function () {

    var SCHEMA = Schema.PICKING_SLIP;

    function handleOnEdit(e) {
        var sheet = e.source.getActiveSheet();
        var range = e.range;
        var row = range.getRow();
        var col = range.getColumn();
        var val = range.getValue();
        var address = range.getA1Notation();

        var headerRow = SCHEMA.FORM_PICKING_SLIP.TABLE_START_ROW - 1;
        var startDataRow = SCHEMA.FORM_PICKING_SLIP.TABLE_START_ROW;

        // Helper to get all column indices dynamically
        var colIndices = {};
        for (var key in SCHEMA.FORM_PICKING_SLIP.COLUMNS) {
            colIndices[key] = SheetHelper.getColumnIndex(sheet, SCHEMA.FORM_PICKING_SLIP.COLUMNS[key], headerRow);
        }

        // Helper to safely get range
        function safeGetRange(r, c) {
            if (c > 0) return sheet.getRange(r, c);
            return null;
        }

        // 1. Handle Kategori Change
        if (address === SCHEMA.FORM_PICKING_SLIP.CELLS.KATEGORI) {
            if (colIndices.ITEM_CODE > 0) {
                var itemCodeRange = sheet.getRange(startDataRow, colIndices.ITEM_CODE, sheet.getMaxRows() - startDataRow + 1);
                itemCodeRange.clearDataValidations();

                if (val === "PM") {
                    setDataValidationFromSheet(sheet, itemCodeRange, Config.SHEET_NAME.PM_LIST, "Item Code");
                } else if (val === "RM") {
                    setDataValidationFromSheet(sheet, itemCodeRange, Config.SHEET_NAME.RM_LIST, "Item Code");
                }
            }
            return;
        }

        // 2. Handle Table Edits
        if (row >= startDataRow) {
            var kategori = sheet.getRange(SCHEMA.FORM_PICKING_SLIP.CELLS.KATEGORI).getValue();

            // --- TRIGGER: ITEM CODE ---
            if (col === colIndices.ITEM_CODE) {
                if (val === "") {
                    sheet.deleteRow(row);
                    if (colIndices.NO > 0 && colIndices.ITEM_CODE > 0) {
                        resetNumber(sheet, colIndices.NO, startDataRow, colIndices.ITEM_CODE);
                    }
                } else {
                    var df = data_filtered(kategori, val, null);

                    if (row > startDataRow) {
                        var dp = data_prev(sheet, row, colIndices, startDataRow);
                        var dm = data_mixing(dp, df);
                        df = dm.except_data;
                    }

                    if (df.length > 0) {
                        var item = df[0];
                        var mCols = getMasterIndices();

                        // Safe clears
                        [colIndices.REQUESTOR, colIndices.AREA_GUDANG, colIndices.LOCATION, colIndices.QTY_REQ, colIndices.QTY_STOCK].forEach(function (c) {
                            if (c > 0) sheet.getRange(row, c).clearContent();
                        });
                        if (colIndices.LOCATION > 0) sheet.getRange(row, colIndices.LOCATION).clearDataValidations();

                        // Safe sets
                        if (colIndices.NO > 0) sheet.getRange(row, colIndices.NO).setValue(row - headerRow);
                        if (colIndices.PO_NUMBER > 0) sheet.getRange(row, colIndices.PO_NUMBER).setValue(item[mCols.PO_NUMBER]);
                        if (colIndices.VENDOR > 0) sheet.getRange(row, colIndices.VENDOR).setValue(item[mCols.VENDOR]);
                        if (colIndices.ITEM_NAME > 0) sheet.getRange(row, colIndices.ITEM_NAME).setValue(item[mCols.ITEM_NAME]);
                        if (colIndices.UOM > 0) sheet.getRange(row, colIndices.UOM).setValue(item[mCols.UOM]);
                        if (colIndices.BATCH_LOT > 0) sheet.getRange(row, colIndices.BATCH_LOT).setValue(item[mCols.BATCH_LOT]);
                        if (colIndices.EXPIRY_DATE > 0) sheet.getRange(row, colIndices.EXPIRY_DATE).setValue(item[mCols.EXPIRY_DATE]);

                        if (colIndices.AREA_GUDANG > 0) {
                            resetAreaGudang(sheet, row, item[mCols.BATCH_LOT], df, colIndices);
                        }
                    }
                }
            }

            // --- TRIGGER: AREA GUDANG ---
            if (col === colIndices.AREA_GUDANG) {
                if (val === "") {
                    if (colIndices.LOCATION > 0) sheet.getRange(row, colIndices.LOCATION).clearDataValidations().clearContent();
                    if (colIndices.QTY_STOCK > 0) sheet.getRange(row, colIndices.QTY_STOCK).clearContent();
                } else {
                    var itemCode = (colIndices.ITEM_CODE > 0) ? sheet.getRange(row, colIndices.ITEM_CODE).getValue() : "";
                    var batchLot = (colIndices.BATCH_LOT > 0) ? sheet.getRange(row, colIndices.BATCH_LOT).getValue() : "";

                    var df = data_filtered(kategori, itemCode, batchLot);

                    if (row > startDataRow) {
                        var dp = data_prev(sheet, row, colIndices, startDataRow);
                        var dm = data_mixing(dp, df);
                        df = dm.except_data;
                    }

                    var detailLokasi = getDetailLokasiData();
                    var locations = [];
                    var mCols = getMasterIndices();
                    var dlCols = getDetailLokasiIndices();

                    for (var i = 0; i < df.length; i++) {
                        for (var j = 0; j < detailLokasi.length; j++) {
                            var locMaster = String(df[i][mCols.LOCATION]).toUpperCase();
                            var locDetail = String(detailLokasi[j][dlCols.LOKASI_RAK]).toUpperCase();
                            var areaDetail = detailLokasi[j][dlCols.AREA_GUDANG];

                            if (locDetail == locMaster && areaDetail == val) {
                                locations.push(detailLokasi[j][dlCols.LOKASI_RAK]);
                            }
                        }
                    }

                    if (colIndices.LOCATION > 0) {
                        var rule = SpreadsheetApp.newDataValidation().requireValueInList(locations).setAllowInvalid(false).build();
                        sheet.getRange(row, colIndices.LOCATION).setDataValidation(rule);
                    }
                }
            }

            // --- TRIGGER: LOCATION ---
            if (col === colIndices.LOCATION) {
                var itemCode = (colIndices.ITEM_CODE > 0) ? sheet.getRange(row, colIndices.ITEM_CODE).getValue() : "";
                var batchLot = (colIndices.BATCH_LOT > 0) ? sheet.getRange(row, colIndices.BATCH_LOT).getValue() : "";
                var qtyReq = (colIndices.QTY_REQ > 0) ? sheet.getRange(row, colIndices.QTY_REQ).getValue() : 0;

                var df = data_filtered(kategori, itemCode, batchLot);
                var mCols = getMasterIndices();

                df = df.filter(function (a) { return a[mCols.LOCATION] == val; });

                if (row > startDataRow) {
                    var dp = data_prev(sheet, row, colIndices, startDataRow);
                    var dm = data_mixing(dp, df);
                    df = dm.except_data;
                }

                if (df.length > 0) {
                    var sisaStock = df[0][mCols.SISA_STOCK];

                    if (qtyReq > sisaStock) {
                        if (colIndices.QTY_STOCK > 0) sheet.getRange(row, colIndices.QTY_STOCK).setValue(sisaStock);

                        var dfNext = data_filtered(kategori, itemCode, null);
                        if (row + 1 > startDataRow) {
                            var dpNext = data_prev(sheet, row + 1, colIndices, startDataRow);
                            var dmNext = data_mixing(dpNext, dfNext);
                            dfNext = dmNext.except_data;
                        }

                        if (dfNext.length > 0) {
                            var nextItem = dfNext[0];
                            var reqRemaining = Number(qtyReq) - Number(sisaStock);
                            var requestor = (colIndices.REQUESTOR > 0) ? sheet.getRange(row, colIndices.REQUESTOR).getValue() : "";

                            sheet.insertRowAfter(row);
                            if (colIndices.NO > 0 && colIndices.ITEM_CODE > 0) resetNumber(sheet, colIndices.NO, startDataRow, colIndices.ITEM_CODE);

                            var nextRow = row + 1;
                            if (colIndices.PO_NUMBER > 0) sheet.getRange(nextRow, colIndices.PO_NUMBER).setValue(nextItem[mCols.PO_NUMBER]);
                            if (colIndices.VENDOR > 0) sheet.getRange(nextRow, colIndices.VENDOR).setValue(nextItem[mCols.VENDOR]);
                            if (colIndices.ITEM_CODE > 0) sheet.getRange(nextRow, colIndices.ITEM_CODE).setValue(nextItem[mCols.ITEM_CODE]);
                            if (colIndices.ITEM_NAME > 0) sheet.getRange(nextRow, colIndices.ITEM_NAME).setValue(nextItem[mCols.ITEM_NAME]);
                            if (colIndices.UOM > 0) sheet.getRange(nextRow, colIndices.UOM).setValue(nextItem[mCols.UOM]);
                            if (colIndices.REQUESTOR > 0) sheet.getRange(nextRow, colIndices.REQUESTOR).setValue(requestor);
                            if (colIndices.BATCH_LOT > 0) sheet.getRange(nextRow, colIndices.BATCH_LOT).setValue(nextItem[mCols.BATCH_LOT]);
                            if (colIndices.QTY_REQ > 0) sheet.getRange(nextRow, colIndices.QTY_REQ).setValue(reqRemaining);
                            if (colIndices.EXPIRY_DATE > 0) sheet.getRange(nextRow, colIndices.EXPIRY_DATE).setValue(nextItem[mCols.EXPIRY_DATE]);

                            if (colIndices.QTY_REQ > 0) sheet.getRange(row, colIndices.QTY_REQ).setValue(sisaStock);

                            if (colIndices.LOCATION > 0) sheet.getRange(nextRow, colIndices.LOCATION).clearDataValidations();
                            if (colIndices.AREA_GUDANG > 0) resetAreaGudang(sheet, nextRow, nextItem[mCols.BATCH_LOT], dfNext, colIndices);
                        }
                    } else {
                        if (colIndices.QTY_STOCK > 0) sheet.getRange(row, colIndices.QTY_STOCK).setValue(sisaStock);
                    }
                }
            }

            // --- TRIGGER: QTY REQ ---
            if (col === colIndices.QTY_REQ) {
                var qtyStock = (colIndices.QTY_STOCK > 0) ? sheet.getRange(row, colIndices.QTY_STOCK).getValue() : 0;
                if (val > qtyStock && qtyStock !== "") {
                    var itemCode = (colIndices.ITEM_CODE > 0) ? sheet.getRange(row, colIndices.ITEM_CODE).getValue() : "";
                    var sisaStock = qtyStock;

                    var dfNext = data_filtered(kategori, itemCode, null);
                    if (row + 1 > startDataRow) {
                        var dpNext = data_prev(sheet, row + 1, colIndices, startDataRow);
                        var dmNext = data_mixing(dpNext, dfNext);
                        dfNext = dmNext.except_data;
                    }

                    if (dfNext.length > 0) {
                        var nextItem = dfNext[0];
                        var mCols = getMasterIndices();
                        var reqRemaining = Number(val) - Number(sisaStock);
                        var requestor = (colIndices.REQUESTOR > 0) ? sheet.getRange(row, colIndices.REQUESTOR).getValue() : "";

                        sheet.insertRowAfter(row);
                        if (colIndices.NO > 0 && colIndices.ITEM_CODE > 0) resetNumber(sheet, colIndices.NO, startDataRow, colIndices.ITEM_CODE);

                        var nextRow = row + 1;
                        if (colIndices.PO_NUMBER > 0) sheet.getRange(nextRow, colIndices.PO_NUMBER).setValue(nextItem[mCols.PO_NUMBER]);
                        if (colIndices.VENDOR > 0) sheet.getRange(nextRow, colIndices.VENDOR).setValue(nextItem[mCols.VENDOR]);
                        if (colIndices.ITEM_CODE > 0) sheet.getRange(nextRow, colIndices.ITEM_CODE).setValue(nextItem[mCols.ITEM_CODE]);
                        if (colIndices.ITEM_NAME > 0) sheet.getRange(nextRow, colIndices.ITEM_NAME).setValue(nextItem[mCols.ITEM_NAME]);
                        if (colIndices.UOM > 0) sheet.getRange(nextRow, colIndices.UOM).setValue(nextItem[mCols.UOM]);
                        if (colIndices.REQUESTOR > 0) sheet.getRange(nextRow, colIndices.REQUESTOR).setValue(requestor);
                        if (colIndices.BATCH_LOT > 0) sheet.getRange(nextRow, colIndices.BATCH_LOT).setValue(nextItem[mCols.BATCH_LOT]);
                        if (colIndices.QTY_REQ > 0) sheet.getRange(nextRow, colIndices.QTY_REQ).setValue(reqRemaining);
                        if (colIndices.EXPIRY_DATE > 0) sheet.getRange(nextRow, colIndices.EXPIRY_DATE).setValue(nextItem[mCols.EXPIRY_DATE]);

                        if (colIndices.QTY_REQ > 0) sheet.getRange(row, colIndices.QTY_REQ).setValue(sisaStock);
                        if (colIndices.LOCATION > 0) sheet.getRange(nextRow, colIndices.LOCATION).clearDataValidations();
                        if (colIndices.AREA_GUDANG > 0) resetAreaGudang(sheet, nextRow, nextItem[mCols.BATCH_LOT], dfNext, colIndices);
                    }
                }
            }
        }
    }

    // --- PRIVATE HELPER FUNCTIONS ---

    function getMasterIndices() {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName(Config.SHEET_NAME.MASTER_INVENTORY);
        var hRow = 1;
        // Helper to safe get index
        var getIdx = function (name) { return SheetHelper.getColumnIndex(sheet, name, hRow) - 1; };

        return {
            ITEM_CODE: getIdx("Item Code"),
            ITEM_NAME: getIdx("Item Name"),
            PO_NUMBER: getIdx("PO Number"),
            TYPE: getIdx("Type"),
            VENDOR: getIdx("Vendor"),
            UOM: getIdx("UOM"),
            BATCH_LOT: getIdx("Batch / Lot"),
            LOCATION: getIdx("Location"),
            EXPIRY_DATE: getIdx("Expiry Date"),
            SISA_STOCK: getIdx("Sisa Stock")
        };
    }

    function getDetailLokasiIndices() {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName(Config.SHEET_NAME.DETAIL_LOKASI);
        var hRow = 1;
        return {
            AREA_GUDANG: SheetHelper.getColumnIndex(sheet, "Area Gudang", hRow) - 1,
            LOKASI_RAK: SheetHelper.getColumnIndex(sheet, "Lokasi Rak", hRow) - 1
        };
    }

    function getDetailLokasiData() {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName(Config.SHEET_NAME.DETAIL_LOKASI);
        return sheet.getDataRange().getValues();
    }

    function data_filtered(kategori, item_code, batch_lot) {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName(Config.SHEET_NAME.MASTER_INVENTORY);
        var data = sheet.getDataRange().getValues();
        var mCols = getMasterIndices();

        // Safety check if master columns are found
        if (mCols.ITEM_CODE < 0 || mCols.SISA_STOCK < 0) return [];

        var filtered = data.filter(function (row) {
            var matchCode = String(row[mCols.ITEM_CODE]) == String(item_code);
            var matchStock = Number(row[mCols.SISA_STOCK]) > 0;
            var matchBatch = batch_lot ? String(row[mCols.BATCH_LOT]) == String(batch_lot) : true;

            if (kategori === "RM") {
                return matchCode && matchStock && matchBatch && row[mCols.EXPIRY_DATE] !== "";
            } else {
                return matchCode && matchStock && matchBatch;
            }
        });

        if (kategori === "RM" && mCols.EXPIRY_DATE >= 0) {
            filtered.sort(function (a, b) {
                return new Date(a[mCols.EXPIRY_DATE]) - new Date(b[mCols.EXPIRY_DATE]);
            });
        }

        return filtered;
    }

    function data_prev(sheet, currentRow, colIndices, startRow) {
        if (currentRow <= startRow) return [];
        if (colIndices.ITEM_CODE < 0) return [];

        var numRows = currentRow - startRow;
        var range = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn());
        var values = range.getValues();

        return values.map(function (row) {
            return {
                item_code: (colIndices.ITEM_CODE > 0) ? row[colIndices.ITEM_CODE - 1] : "",
                batch_lot: (colIndices.BATCH_LOT > 0) ? row[colIndices.BATCH_LOT - 1] : "",
                location: (colIndices.LOCATION > 0) ? row[colIndices.LOCATION - 1] : "",
                qty_stock: (colIndices.QTY_STOCK > 0) ? row[colIndices.QTY_STOCK - 1] : ""
            };
        });
    }

    function data_mixing(dataPrev, dataFilter) {
        var mCols = getMasterIndices();
        var exceptData = [];

        for (var i = 0; i < dataFilter.length; i++) {
            var found = false;
            for (var j = 0; j < dataPrev.length; j++) {
                if (
                    String(dataFilter[i][mCols.ITEM_CODE]) === String(dataPrev[j].item_code) &&
                    String(dataFilter[i][mCols.BATCH_LOT]) === String(dataPrev[j].batch_lot) &&
                    String(dataFilter[i][mCols.LOCATION]) === String(dataPrev[j].location) &&
                    String(dataFilter[i][mCols.SISA_STOCK]) === String(dataPrev[j].qty_stock)
                ) {
                    found = true;
                    break;
                }
            }
            if (!found) {
                exceptData.push(dataFilter[i]);
            }
        }
        return { except_data: exceptData };
    }

    function resetNumber(sheet, noColIndex, startRow, itemCodeColIndex) {
        var lastRow = sheet.getLastRow();
        if (lastRow < startRow) return;
        if (noColIndex < 1 || itemCodeColIndex < 1) return;

        var range = sheet.getRange(startRow, itemCodeColIndex, lastRow - startRow + 1, 1);
        var values = range.getValues();
        var numbers = [];
        var count = 1;

        for (var i = 0; i < values.length; i++) {
            if (values[i][0] !== "") {
                numbers.push([count++]);
            } else {
                numbers.push([""]);
            }
        }

        sheet.getRange(startRow, noColIndex, numbers.length, 1).setValues(numbers);
    }

    function resetAreaGudang(sheet, row, batchLot, df, colIndices) {
        if (colIndices.AREA_GUDANG < 1) return;

        var detailLokasi = getDetailLokasiData();
        var dlCols = getDetailLokasiIndices();
        var mCols = getMasterIndices();
        var areas = [];

        for (var i = 0; i < df.length; i++) {
            for (var j = 0; j < detailLokasi.length; j++) {
                var locMaster = String(df[i][mCols.LOCATION]).toUpperCase();
                var locDetail = String(detailLokasi[j][dlCols.LOKASI_RAK]).toUpperCase();

                if (locDetail == locMaster && df[i][mCols.BATCH_LOT] == batchLot) {
                    var area = detailLokasi[j][dlCols.AREA_GUDANG];
                    if (areas.indexOf(area) === -1) areas.push(area);
                }
            }
        }

        var rule = SpreadsheetApp.newDataValidation().requireValueInList(areas).setAllowInvalid(false).build();
        sheet.getRange(row, colIndices.AREA_GUDANG).setDataValidation(rule);
    }

    function setDataValidationFromSheet(sheet, targetRange, sourceSheetName, colHeader) {
        var ss = sheet.getParent();
        var sourceSheet = ss.getSheetByName(sourceSheetName);
        if (!sourceSheet) return;

        var colIdx = SheetHelper.getColumnIndex(sourceSheet, colHeader, 1);
        if (colIdx === -1) return;

        var lastRow = sourceSheet.getLastRow();
        if (lastRow < 2) return;

        var range = sourceSheet.getRange(2, colIdx, lastRow - 1);
        var rule = SpreadsheetApp.newDataValidation().requireValueInRange(range, true).setAllowInvalid(false).build();
        targetRange.setDataValidation(rule);
    }

    return {
        handleOnEdit: handleOnEdit
    };

})();
