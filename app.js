const fs = require('fs');

var CloudmersiveConvertApiClient = require('cloudmersive-convert-api-client');
var defaultClient = CloudmersiveConvertApiClient.ApiClient.instance;

// Configure API key authorization: Apikey
var Apikey = defaultClient.authentications['Apikey'];
Apikey.apiKey = 'YOUR-API-KEY';


var apiInstance = new CloudmersiveConvertApiClient.EditDocumentApi();

var inputFile = Buffer.from(fs.readFileSync("C:\\temp2\\input.xlsx").buffer); // File | Input file to perform the operation on.

// Read an XLSX

var callback = function(error, data, response) {
  if (error) {
    console.error(error);
  } else {
    console.log('API called successfully. Returned data: ' + data);

    var input = new CloudmersiveConvertApiClient.GetXlsxRowsAndCellsRequest(); // GetXlsxRowsAndCellsRequest | Document input request

    input.InputFileUrl = data;

    var callback2 = function(error, data, response2) {
        if (error) {
          console.error(error);
        } else {
          console.log('API called successfully. Returned data: ' + data);
        }
      };
      apiInstance.editDocumentXlsxGetRowsAndCells(input, callback2);

  }
};
apiInstance.editDocumentBeginEditing(inputFile, callback);

// Create an XLSX

var apiInstance = new CloudmersiveConvertApiClient.EditDocumentApi();

var input = new CloudmersiveConvertApiClient.CreateSpreadsheetFromDataRequest(); // CreateSpreadsheetFromDataRequest | Document input request

input.Rows = new CloudmersiveConvertApiClient.XlsxSpreadsheetRow[1];
input.Rows[0] = new CloudmersiveConvertApiClient.XlsxSpreadsheetRow();
input.Rows[0].Cells = new CloudmersiveConvertApiClient.XlsxSpreadsheetCell[1];
input.Rows[0].Cells[0].CellIdentifier = "A1";
input.Rows[0].Cells[0].TextValue = "Hello, world";

var callback3 = function(error, data, response) {
  if (error) {
    console.error(error);
  } else {
    console.log('API editDocumentXlsxCreateSpreadsheetFromData called successfully. Returned data: ' + data);
  }
};
apiInstance.editDocumentXlsxCreateSpreadsheetFromData(input, callback3);