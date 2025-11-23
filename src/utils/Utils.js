/**
 * Utils
 * General utility functions.
 */
var Utils = {

    /**
     * Generates a random string of lower case characters.
     * @param {number} length 
     * @returns {string}
     */
    generateRandomStringLC: function (length) {
        var result = '';
        var characters = 'abcdefghijklmnopqrstuvwxyz0123456789';
        var charactersLength = characters.length;
        for (var i = 0; i < length; i++) {
            result += characters.charAt(Math.floor(Math.random() * charactersLength));
        }
        return result;
    },

    /**
     * Sets the current timestamp.
     * @returns {Date}
     */
    setTimeStamp: function () {
        return new Date();
    },

    /**
     * Formats a date to a string (default format).
     * @param {Date|string} date 
     * @returns {string}
     */
    formatDefault: function (date) {
        if (!date) return "";
        return Utilities.formatDate(new Date(date), "Asia/Jakarta", "yyyy-MM-dd");
    },

    /**
     * Logs errors to the Error Log sheet.
     * @param {string} ssid 
     * @param {string} sheetName 
     * @param {Error} error 
     */
    logError: function (ssid, sheetName, error) {
        try {
            var ss = SpreadsheetApp.openById(ssid);
            var sheet = ss.getSheetByName(sheetName);
            if (sheet) {
                sheet.appendRow([new Date(), error.message, error.stack]);
            }
        } catch (e) {
            console.error("Failed to log error: " + e.message);
        }
    }
};
