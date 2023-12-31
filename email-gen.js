function extractAndTranslateCompanyName() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var startRow = 2; // Start at row 2 to skip the header
  var numRows = sheet.getLastRow() - 1; // Number of rows to process
  var emailColumn = 1; // Column A has index 1
  var companyColumn = 2; // Column B
  var japaneseCompanyColumn = 3; // Column C for Japanese company name

  // Fetch the range of cells A2:A(last row)
  var dataRange = sheet.getRange(startRow, emailColumn, numRows, 1);
  var data = dataRange.getValues();

  // OpenAI API setup
  var openaiApiKey = 'OpenAI API key';
  var openaiUrl = 'https://api.openai.com/v1/engines/text-davinci-003/completions';

  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var email = row[0];
    if (email) {
      var companyName = email.split('@')[1].split('.')[0]; // Extract the company name
      sheet.getRange(startRow + i, companyColumn).setValue(companyName); // Set the company name in column B

      // Prepare the payload for the OpenAI API call
      var payload = {
        prompt: 'describe "' + companyName + '" to a Japanese company name format":',
        temperature: 0.3,
        max_tokens: 60,
      };

      var options = {
        'method': 'post',
        'contentType': 'application/json',
        'headers': {
          'Authorization': 'Bearer ' + openaiApiKey
        },
        'payload': JSON.stringify(payload)
      };

      // Call the OpenAI API
      var response = UrlFetchApp.fetch(openaiUrl, options);
      var responseJson = JSON.parse(response.getContentText());

      if (responseJson.choices && responseJson.choices.length > 0) {
        var translatedName = responseJson.choices[0].text.trim();
        sheet.getRange(startRow + i, japaneseCompanyColumn).setValue(translatedName); // Set the translated name in column C
      }
    }
  }
}
