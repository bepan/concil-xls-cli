#!/usr/bin/env node

const XLSX = require('xlsx');
const path = require('path');
const _ = require('lodash');
const os = require('os');
const fs = require('fs');
const buildOptions = require('./modules/config-options.module');
const FileNameValidator = require('./modules/file-name.validator');
const StartFromValidator = require('./modules/start-from.validator');
const getExcelData = require('./modules/get-excel-data.helper');
const conciliate = require('./modules/conciliate-logic.module');

// Take start execution time.
const startEx = new Date();

// Build agnostic User's Desktop path
const desktopPath = path.join(os.homedir(), 'Desktop');

// Configure cli options
const options = buildOptions();
const { fileName, startFromCell, month, year } = options;

// Run Argument Validations
FileNameValidator.run(fileName, desktopPath);
StartFromValidator.run(startFromCell);

// Read Excel file from Desktop
const filePath = path.join(desktopPath, fileName);
const workbook = XLSX.readFile(filePath);

// Create output root directory in Desktop
const rootDirName = `conciliacion_${new Date().getTime()}`;
const rootDirFullPath = path.join(desktopPath, rootDirName);
fs.mkdirSync(rootDirFullPath);

// Loop through all Accounts
for (let sheetName of workbook.SheetNames)
{
  // Create directory for each account
  fs.mkdirSync(path.join(rootDirFullPath, sheetName));
  // Get Worksheet
  const ws = workbook.Sheets[sheetName];
  // Group the data rows by Aux
  const grouped = _.groupBy(getExcelData(ws, startFromCell), 'Aux');
  // Get all keys (Auxiliaries)
  const auxiliaries = Object.keys(grouped);

  for (let aux of auxiliaries)
  {
    // Run conciliate routine per aux block
    const { matchesArr, pendingRegs } = conciliate(grouped[aux]);
    // Remove block from memory
    delete grouped[aux];
    // Create new excel per aux
    var newWb = XLSX.utils.book_new();
    var newWsPending = XLSX.utils.json_to_sheet(pendingRegs);
    var newWsDeleted = XLSX.utils.json_to_sheet(matchesArr);
    XLSX.utils.book_append_sheet(newWb, newWsPending, "Pendientes");
    XLSX.utils.book_append_sheet(newWb, newWsDeleted, "Eliminados");
    XLSX.writeFile(newWb, path.join(rootDirFullPath, sheetName, `${sheetName}_${aux}_${month}_${year}.xlsx`));
  }
}

// Calculate Execution time
var endEx = new Date() - startEx;
console.info('Execution time: %dms', endEx);



