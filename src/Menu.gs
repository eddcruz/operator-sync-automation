/**
 * @function onOpen
 * Runs automatically when the Google Sheets file is opened.
 * Creates a custom menu called "Automation" in the toolbar.
 * From this menu, automated actions related to operators can be triggered.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi(); // Get the spreadsheet UI instance

  // Create the custom menu with the available actions
  ui.createMenu('Automation')
    .addItem('Import operators (email)', 'procesarCorreosOperadores')
    // Extracts weekly operator files and updates the database
    .addItem('Query operator deductions (email)', 'revisarFaltantesOperador')
    // Searches and replies to operator deduction emails
    .addToUi(); // Attach the menu to the toolbar
}