const ExcelJS = require('exceljs');
const { ExcelJsService } = require('./src/services/excel-js.service');

const e = new ExcelJsService(ExcelJS);

e.read('/Users/beto/Downloads/Cuentas y Aux Nov 2019.xlsx', function(){
  const workbook = e.createWorkbook('/Users/beto/Desktop/foo.xlsx');
  const worksheet = workbook.addWorksheet('foo');
  worksheet.columns = [
    { header: 'Cta', key: 'Cta', width: 20 },
    { header: 'Aux', key: 'Aux', width: 20 },
    { header: 'Fecha', key: 'Fecha', width: 20},
    { header: 'Pol', key: 'Pol', width: 20 },
    { header: 'Concepto', key: 'Concepto', width: 20 },
    { header: 'CargoCG', key: 'CargoCG', width: 20},
    { header: 'AbonoCG', key: 'AbonoCG', width: 20 },
    { header: 'NomCta', key: 'NomCta', width: 20 },
    { header: 'NomAux', key: 'NomAux', width: 20 },
    { header: 'Cia', key: 'Cia', width: 20 }
  ];
  
  const dataset = e.getDataset('2102', 'a2');

  for(let i = 0; i < dataset.length; i++) {
    worksheet.addRow(dataset[i]).commit();
    dataset[i] = null;
  }

  // Finished adding data. Commit the worksheet
  worksheet.commit();

  // Finished the workbook.
  workbook.commit()
    .then(function() {
      console.log('done!');
    });
  
  // workbook.xlsx.writeFile('/Users/beto/Desktop/foo.xlsx')
  //   .then(function() {
  //     // done
  //     console.log('done!');
  //   });
});

