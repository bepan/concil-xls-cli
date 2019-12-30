const os = require('os');

class ExcelJsService
{
  constructor(lib) {
    this.lib = lib;
    this.workbook = new lib.Workbook();
    this.sheetHeaders = null;
  }

  async read(filePath) {
    try {
      await this.workbook.xlsx.readFile(filePath);
      return;
    } catch (err) {
      throw new Error(err.message);
    }
  }

  getAllWorksheetNames() {
    return this.workbook.worksheets.map(ws => ws.name);
  }

  getDataset(account, startFromCell) {
    const dataset = [];
    const startingRow = +startFromCell.substr(1);
    const worksheet = this.workbook.getWorksheet(account);
    //Iterate over all rows that have values in a worksheet
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber < startingRow) { return; }
      // Get Headers
      if (rowNumber === startingRow) {
        this.sheetHeaders = row;
        return;
      }
      // Dataset elements
      const data = {};
      this.sheetHeaders.eachCell((cell, colNumber) => {
        data[cell.value] = row.getCell(colNumber).value;
      });
      dataset.push(data);
    });

    return dataset;
  }

  getLetterColumn(index)
  {
    const alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']; 
    return alphabet[index];
  }

  styleCell(cell, bgColor='ffffff', align='right', color='000000', wrap=false) {
    cell.alignment = {
      wrapText: wrap,
      horizontal: align,
      vertical: 'middle'
    };
    cell.font = {color: { argb: color }};
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {argb:bgColor}
    };
  }

  styleHeaderTemplate(worksheet, datasetLength, currAccount, currAux)
  {
    // Set all Columns width
    for(let i = 0; i < this.sheetHeaders.cellCount; i++) {
      const letter = this.getLetterColumn(i);
      worksheet.getColumn(letter).width = 15;
    }

    // Style and Set first row
    worksheet.mergeCells('A1:J1');
    const firstCell = worksheet.getCell('A1');
    this.styleCell(firstCell, '00ccff', 'center', undefined, true);
    const ctaDesc = currAccount === '2102' ? 'Acreedores Diversos' : 'Deudores Diversos';
    firstCell.value = `INVERSIONES ACCIONARIAS LANDUS${os.EOL}Auxiliar${os.EOL}Cta: ${currAccount} ${ctaDesc}${os.EOL}(En Varias Monedas)`;
    worksheet.getRow(1).height = 70;

    // Color row 2,3,4,5,6
    [2, 3, 4, 5, 6].forEach(row => {
      for(let i = 0; i < this.sheetHeaders.cellCount; i++) {
        const cell = worksheet.getCell(`${this.getLetterColumn(i)}${row}`);
        this.styleCell(cell, '00ccff');
        cell.value = '';
      }
    });

    // Set text and formulas for header
    worksheet.getCell('F3').value = 'Total Cargos';
    worksheet.getCell('F4').value = { formula: `SUM(F8:F${datasetLength+7})`};
    worksheet.getCell('G3').value = 'Total Abonos';
    worksheet.getCell('G4').value = { formula: `SUM(G8:G${datasetLength+7})`};
    worksheet.getCell('H3').value = 'Saldo';
    worksheet.getCell('H4').value = { formula: 'G4-F4' };

    worksheet.getCell('A4').value = 'Cta';
    worksheet.getCell('A5').value = currAccount;
    worksheet.getCell('B4').value = 'Aux';
    worksheet.getCell('B5').value = currAux;


    // Color header cells
    for(let i = 0; i < this.sheetHeaders.cellCount; i++) {
      const cell = worksheet.getCell(`${this.getLetterColumn(i)}7`);
      this.styleCell(cell, 'FF6699', 'center', 'FFFFFF')
    }
  }

  jsonToSheet(workbook, dataset, sheetName, currAccount, currAux)
  {
    // Add worksheet to workbook
    let sheet = workbook.addWorksheet(sheetName);

    // Build default template for all documents
    this.styleHeaderTemplate(sheet, dataset.length, currAccount, currAux);

    // Print headers
    this.sheetHeaders.eachCell((cell, colNumber) => {
      const letter = this.getLetterColumn(colNumber-1);
      sheet.getCell(`${letter}7`).value = cell.value;
    });

    // Print row data
    for (let i = 0; i < dataset.length; i++) {
      for(let j = 0; j < this.sheetHeaders.cellCount; j++) {
        const letter = this.getLetterColumn(j);
        const prop = this.sheetHeaders.getCell(letter).value;
        sheet.getCell(`${letter}${i+8}`).value = dataset[i][prop];
      }
    }

    // Commit sheet
    sheet.commit();
  }

  createNewWorkbook(fileToWrite)
  {
    // construct a streaming XLSX workbook writer with styles and shared strings
    var options = {
      filename: fileToWrite,
      useStyles: true
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
