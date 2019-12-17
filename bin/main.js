#!/usr/bin/env node
const XLSX = require('xlsx');
const path = require('path');
const _ = require('lodash');
const os = require('os');
const fs = require('fs');
const yargs = require("yargs");
const FileNameValidator = require('./modules/file-name.validator');
const StartFromValidator = require('./modules/start-from.validator');

// Build the user's desktop path
const desktopPath = path.join(os.homedir(), 'Desktop');

// Configure cli options
const options = yargs
  .usage('Basic Usage: concil --fn=example.xlsx --sf=A2')
  .option("fn", { alias: "fileName", describe: "Excel to feed the app.", type: "string", demandOption: true })
  .option("sf",  { alias: "startFrom", describe: "Which Cell to start reading data from.", type: "string", demandOption: true })
  .argv;

// Run Argument Validations
FileNameValidator.run(options.fileName, desktopPath);
StartFromValidator.run(options.startFrom);

// Read Concil file
var startEx = new Date();
const filePath = path.join(desktopPath, options.fileName);
const workbook = XLSX.readFile(filePath);

// Create root output directory in Users Desktop
const rootDir = path.join(desktopPath, `conciliacion_${new Date().getTime()}`);
fs.mkdirSync(rootDir);

// Loop through all worksheets (Cuentas)
for (let sheetName of workbook.SheetNames)
{
  // Create directory for each account
  fs.mkdirSync(path.join(rootDir, sheetName));
  const ws = workbook.Sheets[sheetName];

  // Set the range which the database starts at
  var range = XLSX.utils.decode_range(ws['!ref']);
  range.s.c = XLSX.utils.decode_col( options.sf[0].toUpperCase() );
  range.s.r = XLSX.utils.decode_row( options.sf.substr(1) );
  // range.e.c = 6; // 6 == XLSX.utils.decode_col("G")
  // range.e.c = 6;
  var new_range = XLSX.utils.encode_range(range);

  // Parse and manip the data
  let concilData = XLSX.utils.sheet_to_json(ws, {range: new_range});
  var grouped = _.groupBy(concilData, 'Aux');
  const groupedKeys = Object.keys(grouped);
  concilData = [];

  let matchesArr = [];
  let pendingRegs = [];
  const processedMap = new WeakMap();

  for (let aux of groupedKeys)
  {
    let oldConcept = '', iterStart = 0;
    grouped[aux] = _.sortBy(grouped[aux], ['Concepto']);

    for (let i = 0; i < grouped[aux].length; i++) 
    {
      const currReg = grouped[aux][i];
      if (processedMap.get(currReg)) continue;
      iterStart = currReg.Concepto !== oldConcept ? i : iterStart;
      oldConcept = currReg.Concepto;
      let matchFound = false;
      const amountCurrReg = currReg.CargoCG ? 'CargoCG' : 'AbonoCG';
      const amountIterReg = amountCurrReg === 'CargoCG' ? 'AbonoCG' : 'CargoCG';

      for (let j = iterStart; j < grouped[aux].length; j++) 
      {
        const iterReg = grouped[aux][j];
        if (processedMap.get(iterReg)) continue;
        if (iterReg.Concepto !== oldConcept) break;
        if (currReg.Concepto === iterReg.Concepto && currReg[amountCurrReg] === iterReg[amountIterReg]) 
        {
          matchFound = true;
          matchesArr.push(currReg, iterReg);
          processedMap.set(currReg, true).set(iterReg, true);
          break;
        }
      }

      if (!matchFound)
      {
        processedMap.set(currReg, true);
        pendingRegs.push(currReg);
      }
    }

    delete grouped[aux];

    // Create a new excel file per aux
    var newWb = XLSX.utils.book_new();
    var newWsPending = XLSX.utils.json_to_sheet(pendingRegs);
    var newWsDeleted = XLSX.utils.json_to_sheet(matchesArr);

    XLSX.utils.book_append_sheet(newWb, newWsPending, "Pendientes");
    XLSX.utils.book_append_sheet(newWb, newWsDeleted, "Eliminados");
    XLSX.writeFile(newWb, path.join(rootDir, sheetName, `${sheetName}_${aux}.xlsx`));
    pendingRegs = [];
    matchesArr = []
  }

}

// Calculate Execution time
var endEx = new Date() - startEx;
console.info('Execution time: %dms', endEx);



