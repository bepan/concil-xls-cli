class ExcelJsService
{
  constructor(lib) {
    this.lib = lib;
    this.workbook = new lib.Workbook();
  }

  read(filePath, cb) {
    this.workbook.xlsx.readFile(filePath)
      .then(() => cb())
      .catch(err => cb(true));
  }

  getAllWorksheetNames() {
    return this.workbook.worksheets.map(ws => ws.name);
  }

  getDataset(account, startFromCell) {
    const dataset = [];
    let headersRow = null;
    const startingRow = +startFromCell.substr(1);
    const worksheet = this.workbook.getWorksheet(account);
    //Iterate over all rows that have values in a worksheet
    worksheet.eachRow(function(row, rowNumber) {
      if (rowNumber < startingRow) { return; }
      // Get Headers
      if (rowNumber === startingRow) {
        headersRow = row;
        return;
      }
      // Dataset elements
      const data = {};
      headersRow.eachCell(function(cell, colNumber) {
        data[cell.value] = row.getCell(colNumber).value;
      });
      dataset.push(data);
    });
    return dataset;
  }

  jsonToSheet(workbook, dataset, sheetName)
  {
    let sheet = workbook.addWorksheet(sheetName);
    const firstOb = dataset[0] || {};
    const headers = Object.keys(firstOb);
    const headersConfig = [];
    // First set headers
    for (let header of headers) {
      const newHeader = {header: header, key: header, width: 25};
      headersConfig.push(newHeader);
    }
    sheet.columns = headersConfig;
    for (let data of dataset) {
      sheet.addRow(data).commit();
    }
    sheet.commit();
  }

  createNewWorkbook(fileToWrite)
  {
    // construct a streaming XLSX workbook writer with styles and shared strings
    var options = {
      filename: fileToWrite
    };
    return new this.lib.stream.xlsx.WorkbookWriter(options);
  }
  
  write(workbook, cb)
  {
    workbook.commit()
      .then(() => cb())
      .catch(err => cb(true));
  }
}

module.exports = {
  ExcelJsService
};
