const fs = require('fs');

var CloudmersiveConvertApiClient = require('cloudmersive-convert-api-client');
var defaultClient = CloudmersiveConvertApiClient.ApiClient.instance;

// Configure API key authorization: Apikey
var Apikey = defaultClient.authentications['Apikey'];
Apikey.apiKey = 'YOUR-API-KEY';


var apiInstance = new CloudmersiveConvertApiClient.EditDocumentApi();

var inputFile = Buffer.from(fs.readFileSync("C:\\temp2\\input.xlsx").buffer); // File | Input file to perform the operation on.


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