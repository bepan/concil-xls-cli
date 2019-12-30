// const XLSX = require('xlsx');
// const { SheetJsService } = require('./services/sheet-js.service');
const ExcelJS = require('exceljs');
const { ExcelJsService } = require('../services/excel-js.service');
const path = require('path');
const groupBy = require('lodash.groupby');
const fs = require('fs');
const FileNameValidator = require('../validators/file-name.validator');
const StartFromValidator = require('../validators/start-from.validator');
const chargePaymentConcil = require('./conciliate-logic.module');
const workerpool = require('workerpool');

async function main(file, startFromCell, month, year, outDir) {
  // Create objects
  // const excel = new SheetJsService(XLSX);
  const excel = new ExcelJsService(ExcelJS);

  // Take start execution time.
  const startEx = new Date();

  // Run Argument Validations
  try {
    FileNameValidator.run(file);
    StartFromValidator.run(startFromCell);
  } 
  catch (err) {
    throw new Error(err.message);
  }

  // Read Excel file
  try {
    console.log('start reading file...');
    await excel.read(file);
    console.log('finish reading file...');
  
    // Create output root directory in destination folder
    const rootDirName = `concil_${month}_${year}__${new Date().getTime()}`;
    const rootDirFullPath = path.join(outDir, rootDirName);
    fs.mkdirSync(rootDirFullPath);

    // Loop through all Accounts
    const sheetNames = excel.getAllWorksheetNames();
    for (let account of sheetNames)
    {
      // Create directory for each account
      fs.mkdirSync(path.join(rootDirFullPath, account));

      // Group the data rows by Aux
      const grouped = groupBy(excel.getDataset(account, startFromCell), 'Aux');

      // Loop through all auxiliars
      for (let aux of Object.keys(grouped))
      {
        // Run conciliate routine per aux block
        const { matchesArr, pendingRegs } = chargePaymentConcil(grouped[aux]);

        // Remove block from memory
        delete grouped[aux];

        // Create new excel per aux
        const newFile = path.join(rootDirFullPath, account, `${account}_${aux}_${month}_${year}.xlsx`);
        var newWorkbook = excel.createNewWorkbook(newFile);
        excel.jsonToSheet(newWorkbook, pendingRegs, "Pendientes", account, aux);
        excel.jsonToSheet(newWorkbook, matchesArr, "Eliminados", account, aux);
        excel.write(newWorkbook, function(err) {
          if (err) {
            throw new Error('There was a problem creating one file, try again.');
          }
        });
      }
    }

    // Return Execution time
    return (new Date() - startEx);
  }
  catch (err) {
    throw new Error('There was a problem reading the input file, try again.');
  }
}

// create a worker and register public functions
workerpool.worker({
  conciliate: main
});

module.exports = main;



