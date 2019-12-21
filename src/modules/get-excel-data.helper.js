const XLSX = require('xlsx');

module.exports = function(ws, startFromCell) {
  // Set the range which the database starts at
  const range = XLSX.utils.decode_range(ws['!ref']);
  range.s.c = XLSX.utils.decode_col( startFromCell[0].toUpperCase() );
  range.s.r = XLSX.utils.decode_row( startFromCell.substr(1) );
  // range.e.c = 6; // 6 == XLSX.utils.decode_col("G")
  // range.e.c = 6;
  const new_range = XLSX.utils.encode_range(range);

  // Parse and manip the data
  return XLSX.utils.sheet_to_json(ws, {range: new_range});
};
