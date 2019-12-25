class SheetJsService
{
  constructor(lib) {
    this.lib = lib;
    this.workbook = null;
    this.fileToWrite = '';
  }

  read(filePath, cb) {
    try {
      this.workbook = this.lib.readFile(filePath);
      cb();
    } catch(error) {
      cb(true);
    }
  }

  getAllWorksheetNames() {
    // Return an array of sheet names
    return this.workbook.SheetNames;
  }

  getDataset(account, startFromCell) {
    const ws = this.workbook.Sheets[account];
    // Set the range which the database starts at
    const range = this.lib.utils.decode_range(ws['!ref']);
    range.s.c = this.lib.utils.decode_col( startFromCell[0].toUpperCase() );
    range.s.r = this.lib.utils.decode_row( startFromCell.substr(1) );
    const new_range = this.lib.utils.encode_range(range);
    // Parse and manip the data
    const dataset = this.lib.utils.sheet_to_json(ws, {range: new_range});
    return dataset;
  }

  jsonToSheet(workbook, dataset, sheetName)
  {
    var newWorksheet = this.lib.utils.json_to_sheet(dataset);
    this.lib.utils.book_append_sheet(workbook, newWorksheet, sheetName);
  }

  createNewWorkbook(fileToWrite)
  {
    this.fileToWrite = fileToWrite;
    return this.lib.utils.book_new();
  }

  write(workbook, cb)
  {
    try {
      this.lib.writeFile(workbook, this.fileToWrite);
      cb();
    } catch (error) {
      cb(true)
    }
  }
}

module.exports = {
  SheetJsService
};
