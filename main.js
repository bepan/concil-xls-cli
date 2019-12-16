const XLSX = require('xlsx');
const path = require('path');

// Read Concil file
const filePath = path.join(__dirname, '1105 30480 Comercio Otros deudores operaciones -Julio 2018.xlsb');
const workbook = XLSX.readFile(filePath, {cellDates: true});

// Get the first Worksheet
const firstSheetName = workbook.SheetNames[0];
const ws = workbook.Sheets[firstSheetName];

// Set the range which the database starts at
var range = XLSX.utils.decode_range(ws['!ref']);
range.s.c = 0; // 0 == XLSX.utils.decode_col("A")
range.s.r = 6;
// range.e.c = 6; // 6 == XLSX.utils.decode_col("G")
// range.e.c = 6;
var new_range = XLSX.utils.encode_range(range);

// Parse and manip the data
const concilData = XLSX.utils.sheet_to_json(ws, {range: new_range});
const matchesArr = [];
const pendingRegs = [];
const processedMap = new Map();

var startEx = new Date();
for (let currReg of concilData) 
{
  if (processedMap.get(currReg)) continue;
  let matchFound = false;
  const amountCurrReg = currReg.CargoCG ? 'CargoCG' : 'AbonoCG';
  const amountIterReg = amountCurrReg === 'CargoCG' ? 'AbonoCG' : 'CargoCG';

  for (let iterReg of concilData) 
  {
    if (currReg === iterReg || processedMap.get(iterReg)) continue;
    if (currReg.Concepto === iterReg.Concepto &&
        currReg[amountCurrReg] === iterReg[amountIterReg]) 
    {
      matchFound = true;
      matchesArr.push(currReg, iterReg);
      processedMap.set(currReg, true);
      processedMap.set(iterReg, true);
      break;
    }
  }

  if (!matchFound)
  {
    processedMap.set(currReg, true);
    pendingRegs.push(currReg);
  }
}

var endEx = new Date() - startEx;
console.info('Execution time: %dms', endEx);

// Write to a new book
var newWb = XLSX.utils.book_new();
var newWsPending = XLSX.utils.json_to_sheet(pendingRegs);
var newWsDeleted = XLSX.utils.json_to_sheet(matchesArr);

XLSX.utils.book_append_sheet(newWb, newWsPending, "Pendientes");
XLSX.utils.book_append_sheet(newWb, newWsDeleted, "Eliminados");
XLSX.writeFile(newWb, 'output.xlsx');

