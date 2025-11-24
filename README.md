# Warehouse Management System (WMS) Script

## Overview
This project contains a modular Google Apps Script solution for a Warehouse Management System (WMS). It is designed to manage inventory, picking slips, and data synchronization efficiently using Google Sheets as the backend.

## Key Features
- **Modular Architecture**: Code is organized into `core`, `features`, `config`, and `utils` for better maintainability.
- **Picking Slip Management**: Automated validation, stock checking, and row splitting based on inventory logic.
- **FEFO Logic**: First-Expired-First-Out logic for automatic batch selection.
- **Dynamic Schema**: Flexible column mapping that adapts to header name changes in the sheet.
- **Inventory Sync**: Automatic synchronization of stock levels.

## Architecture
The project structure is organized as follows:
- `src/config`: Configuration and schema definitions.
- `src/core`: Core routing and sheet helpers.
- `src/features`: Business logic for Picking Slips and shared services.
- `src/utils`: General utility functions.

## Setup & Usage
1.  **Dependencies**: Ensure all dependencies are listed in `appsscript.json`.
2.  **Configuration**: Update `src/config/Config.js` with your Spreadsheet IDs and Sheet Names.
3.  **Deployment**: Push the script to your Google Apps Script project using `clasp`.

For detailed developer documentation, please refer to `PANDUAN_DEVELOPER.md`.

---
*Note: This project is private and intended for internal use.*
