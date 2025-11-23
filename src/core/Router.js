/**
 * Router
 * Dispatches events to the appropriate controller based on the active sheet.
 */
var Router = (function () {

    function routeOnEdit(e) {
        var sheet = e.source.getActiveSheet();
        var sheetName = sheet.getName();

        // Route for Picking Slip Forms
        if (sheetName === Config.SHEET_NAME.FORM_PICKING_SLIP_ADV ||
            sheetName === Config.SHEET_NAME.FORM_PICKING_SLIP_KYM) {
            PickingSlipController.handleOnEdit(e);
        }

        // Add other routes here as needed
        // if (sheetName === Config.SHEET_NAME.OTHER_FORM) { ... }
    }

    return {
        routeOnEdit: routeOnEdit
    };

})();
