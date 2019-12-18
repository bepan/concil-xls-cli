const yargs = require("yargs");

module.exports = function() {
  return yargs
    .usage('Usage: concil --fn=example.xlsx --sfc=A2')
    .usage('Usage with spaces in file name: concil --fn="example 2019.xlsx" --sfc=A2')
    // Options
    .option("f", { 
      alias: "fileName", 
      describe: "Excel to feed the app.", 
      type: "string", 
      demandOption: true 
    })
    .option("st", {
      alias: "startFromCell", 
      describe: "Which Cell to start reading data from.",
      type: "string",
      demandOption: true
    })
    .option("m", {
      alias: "month", 
      describe: "Month that will be part of the output files.",
      type: "string",
      demandOption: true
    })
    .option("y", {
      alias: "year", 
      describe: "Year that will be part of the output files.",
      type: "string",
      demandOption: true
    })
    .argv;
}
