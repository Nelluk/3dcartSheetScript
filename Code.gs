function exampleQuery() {
  // This is an example of how to run a query with your own function (not from the menu). Results will be appended
  // to the first sheet.
  
  var query = "SELECT TOP 5 email,custenabled,maillist from customers WHERE maillist=1;";
  var result = queryCart(query);

  var result_array = xmlResultToArray(result, true);
  arrayToSheet(result_array);
  Logger.log("Number of results:" + result_array.length);
  addQueryHistory(query, result_array.length);
}
