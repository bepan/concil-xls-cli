var workerpool = require('workerpool');

const XLSX = require('xlsx');
const path = require('path');
const _ = require('lodash');
const fs = require('fs');
const FileNameValidator = require('./modules/file-name.validator');
const StartFromValidator = require('./modules/start-from.validator');
const getExcelData = require('./modules/get-excel-data.helper');
const conciliate = require('./modules/conciliate-logic.module');

function conciliateLogic({file, startFromCell, month, year, outDir}) {

  console.log('start generating files...');
  
  // Take start execution time.
  const startEx = new Date();

  // Run Argument Validations
  FileNameValidator.run(file);
  StartFromValidator.run(startFromCell);

  // Read Excel file
  const workbook = XLSX.readFile(file);

  // Create output root directory in Desktop
  const rootDirName = `concil_${month}_${year}__${new Date().getTime()}`;
  const rootDirFullPath = path.join(outDir, rootDirName);
  fs.mkdirSync(rootDirFullPath);

  // Loop through all Accounts
  for (let account of workbook.SheetNames)
  {
    // Create directory for each account
    fs.mkdirSync(path.join(rootDirFullPath, account));
    // Get Worksheet
    const ws = workbook.Sheets[account];
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
      XLSX.writeFile(newWb, path.join(rootDirFullPath, account, `${account}_${aux}_${month}_${year}.xlsx`));
    }
  }

  // Calculate Execution time
  var endEx = new Date() - startEx;
  console.info('Execution time: %dms', endEx);
};

// create a worker and register public functions
workerpool.worker({
  conciliate: conciliateLogic
});



