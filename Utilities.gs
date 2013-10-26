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
  
  var result = cartAPI.invokeOperation("runQuery", [param]);
  var result_json = xmlToJson_(result.Envelope.Body.runQueryResponse.runQueryResult);
  
  return result_json;
}


function exampleQuery3() {
  // This is an example of how to run a query with your own function (not from the menu). Results will be appended
  // to the first sheet.
  
  var query = "SELECT invoicenum,ofirstname,olastname FROM orders WHERE order_status IN (1,2) AND (oshipaddress <> oaddress)";
  var result = queryCart(query);

  resultToSheet(result);
}




function resultToSheet(result) {
  // Takes a result of a query, converts it to arrays of data, and appends it to the spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  
  if ("Error" in result) {
    Logger.log("Error during call: " + result["Error"]);
    sheet.appendRow([["Error"],[result["Error"]]]);
    return;
  }
  
  
  
  if ("Error" in result['runQueryResponse']) {
    Logger.log("Error during call: " + result['runQueryResponse']['Error']['Description']);
    sheet.appendRow([["Error"],[result['runQueryResponse']['Error']['Description']]]);
    return;
  }
  
  
  
  var data = [].concat(result["runQueryResponse"]["runQueryRecord"]);
  Logger.log(data.length);
  for (var i = 0; i < data.length; i++) {
    var row = [];
    var header = [];
    for (var keys in data[i]) {
      row.push(data[i][keys]);
      if (i == 0) {
        header.push(keys);
      }
    }
    if (i==0) {sheet.appendRow(header);};
    sheet.appendRow(row);
  }
}






function xmlToJson_(xml) {
 // Converts a given xml element to JSON object. If no XML is supplied will use the root-level element of the
 // request object. Adapted from http://davidwalsh.name/convert-xml-json
  
  // Create the return object
  var obj = {};
  var xmltext = xml.getText();
  var foo = xmltext.length;
  if (xmltext.length == 0) { // element
    // do attributes
    if (xml.getAttributes().length > 0) {
    obj["@attributes"] = {};
      for (var j = 0; j < xml.getAttributes().length; j++) {
        var attribute = xml.getAttributes()[j];
        obj["@attributes"][attribute.getName().getLocalName()] = attribute.getValue();
      }
    }
  } else { // text
    obj = xmltext;
  }

  // do children
  if (xml.getElements().length) {
    for(var i = 0; i < xml.getElements().length; i++) {
      var item = xml.getElements()[i];
      var nodeName = item.getName().getLocalName();
      if (typeof(obj[nodeName]) == "undefined") {
        obj[nodeName] = this.xmlToJson_(item);
      } else {
        if (typeof(obj[nodeName].push) == "undefined") {
          var old = obj[nodeName];
          obj[nodeName] = [];
          obj[nodeName].push(old);
        }
        obj[nodeName].push(this.xmlToJson_(item));
      }
    }
  }
  return obj;
};



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