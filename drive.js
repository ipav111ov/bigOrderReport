function getValuesFromSS() {
  const folder = DriveApp.getFolderById(id = '1YKMMn7uFW1yYc8b8pNAvE-OdYqxwtm1l');
  const files = folder.getFiles();
  let result = [];
  while (files.hasNext()) {
    const file = files.next();
    let currentValues = SpreadsheetApp.openById(file.getId()).getSheets()[0].getDataRange().getValues()
    currentValues = currentValues[1][0] == '' ? currentValues.slice(2) : currentValues.slice(1)
    result = result.concat(currentValues)
  }
  return result;
}

function test() {
  const filesArray = readFiles();
  const data = new Data()
  filesArray.forEach(xlsx => {
    const file = Drive.Files.copy(
      { title: 'tempxlsxtosheet', mimeType: MimeType.GOOGLE_SHEETS },
      xlsx.getId(),
      { convert: true }
    );
    data.addValues(SpreadsheetApp.openById(file.getId()).getSheets()[0].getDataRange().getValues())
    // Logger.log(SpreadsheetApp.openById(file.getId()).getSheets()[0].getDataRange().getValues()[0])
  })
  data.getCalculations()
  Logger.log(data)
}