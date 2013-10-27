function onOpen() {
  // Create spreadsheet menu, runs when file is opened
  var ss = SpreadsheetApp.getActive();
  var items = [
    {name: 'Enter Query', functionName: 'menuQuery'},
    null, // Results in a line separator.
    {name: 'Example: New Orders', functionName: 'example1'}
  ];
  ss.addMenu('3DCart', items);
}



function menuQuery() {
  // First menu item, allows user to enter a query into a dialog box.
  var userQuery = Browser.inputBox('Enter an SQL query');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  sheet.appendRow([["This query:"],[userQuery]]);
  
  var result = queryCart(userQuery);
  resultToSheet(result);
}



function example1() {
  // Example query in menu, list all orders with an order_status of 1 (new orders)
  Browser.msgBox('The current sheet will be appended with the results of this query.');
  
  var query = "SELECT invoicenum,ofirstname,olastname from orders WHERE order_status=1;";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  sheet.appendRow([["This query:"],[query]]);
  
  var result = queryCart(query);
  resultToSheet(result);
}
