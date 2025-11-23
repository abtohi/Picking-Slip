/**
 * Main Entry Point for onEdit Trigger
 */
function onEdit(e) {
  try {
    Router.routeOnEdit(e);
  } catch (error) {
    Utils.logError(Config.SSID.ERROR_LOG, Config.SHEET_NAME.ERROR, error);
    Browser.msgBox("Error in onEdit: " + error.message);
  }
}

/**
 * Main Entry Point for onOpen Trigger
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation')
    .addItem('Update Data', 'updateData')
    .addToUi();
}

/**
 * Exposed functions for Buttons/Menus
 */
function updateData() {
  SyncService.updateData(true);
}

function savePickingSlip() {
  var sheet = SpreadsheetApp.getActiveSheet();
  PickingSlipService.saveData(sheet);
}

function clearPickingSlip() {
  var sheet = SpreadsheetApp.getActiveSheet();
  PickingSlipService.clearForm(sheet);
}