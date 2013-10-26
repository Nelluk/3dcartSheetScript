// THIS SCRIPT NEEDS TO BE ATTACHED TO A SPREADSHEET
// Create a Google Spreadsheet, go to Tools>Script Editor, and copy each of these files into a blank script.

function exampleQuery() {
  // This is an example of how to run a query with your own function (not from the menu). Results will be appended
  // to the first sheet.
  
  var query = "SELECT invoicenum,ofirstname,olastname from orders WHERE order_status=2";
  var result = queryCart(query);

  resultToSheet(result);
}
