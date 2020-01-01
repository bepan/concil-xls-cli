const os = require('os');
const XlsxStreamReader = require("xlsx-stream-reader");
const fs = require('fs');

class ExcelJsService
{
  constructor(lib) {
    this.lib = lib;
    this.workbook = new lib.Workbook();
    this.sheetHeaders = [];
  }

  read(filePath, startFromCell) {
    return new Promise((resolve, reject) => {
      const startingRow = +startFromCell.substr(1);
      let sheets = {}, sheetRows = [];
      let workBookReader = new XlsxStreamReader();
      workBookReader.on('error', error => reject(error));
      workBookReader.on('worksheet', (workSheetReader) =>
      {
        workSheetReader.on('row', (row) => {
          if (+row.attributes.r < startingRow) return;
          if (+row.attributes.r === startingRow) {
            this.sheetHeaders = row.values;
            return;
          }
          let rowObj = {};
          row.values.forEach((rowVal, colNum) => {
            // do something with row values
            const header = this.sheetHeaders[colNum];
            if (['CargoCG', 'AbonoCG'].includes(header)) {
              rowObj[header] = +rowVal.replace(/[,-]/g, '').trim();
              return;
            }
            rowObj[header] = rowVal;
          });
          sheetRows.push(rowObj);
        });

        workSheetReader.on('end', function () {
          sheets[workSheetReader.name] = sheetRows;
          sheetRows = [];
        });
        // call process after registering handlers
        workSheetReader.process();
      });
      
      // End reading all file
      workBookReader.on('end', function() {
        resolve(sheets);
      });

      fs.createReadStream(filePath)
        .pipe(workBookReader);
    });
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
    for(let i = 0; i < this.sheetHeaders.length - 1; i++) {
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
      for(let i = 0; i < this.sheetHeaders.length - 1; i++) {
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
    for(let i = 0; i < this.sheetHeaders.length - 1; i++) {
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
    this.sheetHeaders.forEach((value, colNumber) => {
      if (colNumber === 0) { return; }
      const letter = this.getLetterColumn(colNumber-1);
      sheet.getCell(`${letter}7`).value = value;
    });

    // Print row data
    for (let i = 0; i < dataset.length; i++) {
      for(let j = 0; j < this.sheetHeaders.length - 1; j++) {
        const letter = this.getLetterColumn(j);
        const prop = this.sheetHeaders[j+1];
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
