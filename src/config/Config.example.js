/**
 * Configuration File (TEMPLATE)
 * Copy this file to Config.js and fill in the actual values.
 */
var Config = {
    ADMIN_EMAIL: "admin@example.com",

    // UI Settings
    UI: {
        BACKGROUND_CELL_AUTO: "#cfe2f3",
        BACKGROUND_CELL_INPUT: "#ffffff",
        BACKGROUND_CELL_HEADER: "#e6b8af"
    },

    // Email Notifications
    EMAIL_NOTIF: {
        REKAP_BARANG_DATANG: {
            to: "email@example.com",
            cc: "email@example.com",
            chat_id_telegram: "12345678"
        },
        MUTASI_RM_PM: {
            to: "email@example.com",
            cc: "email@example.com",
            chat_id_telegram: "12345678"
        }
    },

    // Spreadsheet IDs (SSID)
    SSID: {
        MASTER_INVENTORY: "YOUR_SSID_HERE",
        INVENTORY_STOCK: "YOUR_SSID_HERE",
        MOVEMENT_BARANG: "YOUR_SSID_HERE",
        MUTASI_RM_PM: "YOUR_SSID_HERE",
        PICKING_SLIP: "YOUR_SSID_HERE",
        REKAP_BARANG_DATANG: "YOUR_SSID_HERE",
        TEMPLATE_FORM_MUTASI_RM_PM: "YOUR_SSID_HERE",
        TEMPLATE_FORM_REKAP_BARANG_DATANG: "YOUR_SSID_HERE",
        ERROR_LOG: "YOUR_SSID_HERE"
    },

    // Sheet Names
    SHEET_NAME: {
        MASTER_INVENTORY: "Master Inventory",
        DETAIL_LOKASI: "Detail Lokasi",
        FORM_PICKING_SLIP_ADV: "Form Picking Slip ADV",
        FORM_PICKING_SLIP_KYM: "Form Picking Slip KYM",
        HISTORY: "History",
        PRINT: "Print",
        PM_LIST: "PM List",
        RM_LIST: "RM List",
        ERROR: "Error Log"
    }
};
