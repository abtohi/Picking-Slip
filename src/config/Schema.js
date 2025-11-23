/**
 * Database Structure and Schema Definitions
 * Defines the expected column headers and structure for each sheet.
 */
var Schema = {
    PICKING_SLIP: {
        MASTER_INVENTORY: {
            HEADERS: [
                "Item Code", "Item Name", "PO Number", "Type", "Vendor",
                "UOM", "Batch / Lot", "Location", "Expiry Date", "Sisa Stock"
            ],
            // Mapping internal keys to header names for easier reference
            COLUMNS: {
                ITEM_CODE: "Item Code",
                ITEM_NAME: "Item Name",
                PO_NUMBER: "PO Number",
                TYPE: "Type",
                VENDOR: "Vendor",
                UOM: "UOM",
                BATCH_LOT: "Batch / Lot",
                LOCATION: "Location",
                EXPIRY_DATE: "Expiry Date",
                SISA_STOCK: "Sisa Stock"
            }
        },
        DETAIL_LOKASI: {
            HEADERS: ["Area Gudang", "Lokasi Rak"],
            COLUMNS: {
                AREA_GUDANG: "Area Gudang",
                LOKASI_RAK: "Lokasi Rak"
            }
        },
        FORM_PICKING_SLIP: {
            // Single Value Cells (Not Table Headers)
            CELLS: {
                DATE: "B3", // Value cell
                PICKER: "E3",
                CHECKER: "E4",
                KATEGORI: "B4",
                DATE_LABEL: "A3",
                PICKER_LABEL: "D3",
                CHECKER_LABEL: "D4",
                KATEGORI_LABEL: "A4"
            },
            // Table Headers (Row 7 in original, data starts Row 8)
            TABLE_START_ROW: 8,
            HEADERS: [
                "No", "PO Number", "Vendor", "Item Code", "Item Name",
                "UOM", "Requestor", "Batch / Lot", "Area Gudang", "Location",
                "Qty Stock", "Qty Req", "Expiry Date"
            ],
            COLUMNS: {
                NO: "No",
                PO_NUMBER: "PO Number",
                VENDOR: "Vendor",
                ITEM_CODE: "Item Code",
                ITEM_NAME: "Item Name",
                UOM: "UOM",
                REQUESTOR: "Requestor",
                BATCH_LOT: "Batch / Lot",
                AREA_GUDANG: "Area Gudang",
                LOCATION: "Location",
                QTY_STOCK: "Qty Stock",
                QTY_REQ: "Qty Req",
                EXPIRY_DATE: "Expiry Date"
            }
        },
        HISTORY: {
            HEADERS: [
                "ID", "Timestamp", "Date", "Wave", "Kategori", "PO Number", "Vendor",
                "Item Code", "Item Name", "UOM", "Requestor", "Batch / Lot", "Area Gudang",
                "Location", "Expiry Date", "Qty Stock", "Qty Stock (g)", "Qty Req",
                "Qty Req (g)", "Qty Pengambilan", "Qty Pengambilan (g)", "Sisa Stock",
                "Sisa Stock (g)", "Picker", "Checker", "Tgl Pengambilan Barang"
            ],
            COLUMNS: {
                ID: "ID",
                TIMESTAMP: "Timestamp",
                DATE: "Date",
                WAVE: "Wave",
                KATEGORI: "Kategori",
                PO_NUMBER: "PO Number",
                VENDOR: "Vendor",
                ITEM_CODE: "Item Code",
                ITEM_NAME: "Item Name",
                UOM: "UOM",
                REQUESTOR: "Requestor",
                BATCH_LOT: "Batch / Lot",
                AREA_GUDANG: "Area Gudang",
                LOCATION: "Location",
                EXPIRY_DATE: "Expiry Date",
                QTY_STOCK: "Qty Stock",
                QTY_REQ: "Qty Req",
                PICKER: "Picker",
                CHECKER: "Checker"
            }
        },
        PRINT: {
            CELLS: {
                DATE: "C4",
                PICKER: "M4"
            },
            TABLE_START_ROW: 7,
            HEADERS: [
                "No", "PO Number", "Vendor", "Item Code", "Item Name", "Qty Req",
                "Qty Picked", "UOM", "Requestor", "Batch / Lot", "Area Gudang",
                "Location", "Expiry Date", "Sisa Stock"
            ],
            COLUMNS: {
                NO: "No",
                PO_NUMBER: "PO Number",
                VENDOR: "Vendor",
                ITEM_CODE: "Item Code",
                ITEM_NAME: "Item Name",
                QTY_REQ: "Qty Req",
                QTY_PICKED: "Qty Picked",
                UOM: "UOM",
                REQUESTOR: "Requestor",
                BATCH_LOT: "Batch / Lot",
                AREA_GUDANG: "Area Gudang",
                LOCATION: "Location",
                EXPIRY_DATE: "Expiry Date",
                SISA_STOCK: "Sisa Stock"
            }
        }
    }
};
