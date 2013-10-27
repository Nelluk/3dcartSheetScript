var api_key = "";
var store_url = "";
// These settings can be hardcoded or placed in the "SETTINGS" sheet

function queryCart(strSql) {
  // This script relies on the SoapService which was deprecated beginning 7/2013 and will be unavailable 2/2014!
  
  loadSettingsFromSheet();
  var wsdl = SoapService.wsdl("http://api.3dcart.com/cart_advanced.asmx?WSDL");
  var cartAPI = wsdl.getService("cartAPIAdvanced");
  
  var param = Xml.element("runQuery", [
                Xml.attribute("xmlns", "http://3dcart.com/"),
                Xml.element("storeUrl", [store_url]),
                Xml.element("userKey", [api_key]),
                Xml.element("callBackURL", [""]),
                Xml.element("sqlStatement", [strSql])
              ]);
  
  
//  var envelope = cartAPI.getSoapEnvelope("runQuery", param); // Use to examine request without sending
  
  var result_full = cartAPI.invokeOperation("runQuery", [param]);
  var result = result_full.Envelope.Body.runQueryResponse.runQueryResult;

  return result;
}






function xmlResultToArray(xml, include_headers) {
  // Given raw XML from queryCart() function, transforms into a two-dimensional array:
  // [[Result1-Field1,Result1-Field2],[Result2-Field1,Result2-Field2]]
  // This is how Google Docs represents a spreadsheet's rows and columns.
  // If 'true' is given as the second argument, the first row of the returned array will include field names.
  
  var return_array = [];
  var headers = [];
  
  if ("Error" in xml) {
    var error_msg = xml.Error.getText();
    Logger.log("Error in result: " + error_msg);
    return_array.push([["Error"],[error_msg]]);
    return return_array;
  }
  
  if ("Error" in xml.runQueryResponse) {
    var error_msg = xml.runQueryResponse.Error["Description"].getText();
    Logger.log("Error in result: " + error_msg);
    return_array.push([["Error"],[error_msg]]);
    return return_array;
  }
  
  // No errors in results if we've reached this point.
  var results = [].concat(xml.runQueryResponse.runQueryRecord);
  
  for (var i = 0; i < results.length; i++) {
    var elems = results[i].getElements();
    var row = [];
    for (var j = 0; j < elems.length; j++) {
      row.push(elems[j].getText());
      if(i == 0) {
        headers.push(elems[j].getName().getLocalName());
      }
      
    }
    if(i == 0 && include_headers == true) {return_array.push(headers);}
    return_array.push(row);
  }
  
  return return_array;
}




function arrayToSheet(data_array) {
  // Takes a two dimensional array as data_array (from the xmlResultToArray function)
  // and replaces the contents of the 'Results' sheet with that data.
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Results");
  
  if (!sheet) {
    // If there is not already a 'Results' sheet, create a fresh one.
    sheet = ss.insertSheet("Results",0);
  } 
  
  
  
  sheet.getDataRange().clearContent(); // Clear contents of sheet
  sheet.getRange(1, 1, data_array.length, data_array[0].length).setValues(data_array);
}



function loadSettingsFromSheet() {
  
  // Looks for strSettingsName in column 1 of the Settings sheet. Returns value in the adjacent row in column 2
  
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var arrSettings = doc.getSheetByName("SETTINGS").getDataRange().getValues();
      
  for (var i = 0; i < arrSettings.length; i++) {
    if (arrSettings[i][0] == "api_key") {
      api_key = arrSettings[i][1];
    }
    if (arrSettings[i][0] == "store_url") {
      store_url = arrSettings[i][1];
    }
  }
  Logger.log("Loaded API key from settings: " + api_key);
  Logger.log("Loaded store URL from settings: " + store_url);
  return;
}



function addQueryHistory(query,num_results) {

  //For todays date;
  Date.prototype.today = function(){ 
    return ((this.getDate() < 10)?"0":"") + this.getDate() +"/"+(((this.getMonth()+1) < 10)?"0":"") + (this.getMonth()+1) +"/"+ this.getFullYear();
  };
  //For the time now
  Date.prototype.timeNow = function(){
     return ((this.getHours() < 10)?"0":"") + this.getHours() +":"+ ((this.getMinutes() < 10)?"0":"") + this.getMinutes() +":"+ ((this.getSeconds() < 10)?"0":"") + this.getSeconds();
  };
  
  var newDate = new Date();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("History");
  
  if (!sheet) {
    Logger.log("No sheet named 'History'. Nothing added.");
    return;
  }
  
  sheet.appendRow([newDate.today() + " @ " + newDate.timeNow(), num_results,query]);
  
}