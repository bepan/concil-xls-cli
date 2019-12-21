const yargs = require("yargs");

module.exports = function() {
  return yargs
    .usage('Usage: concil --fn=example.xlsx --sfc=A2')
    .usage('Usage with spaces in file name: concil --fn="example 2019.xlsx" --sfc=A2')
    // Options
    .option("f", { 
      alias: "file", type: "string", demandOption: true,
      describe: "Full path Excel file to feed the app.", 
    })
    .option("st", {
      alias: "startFromCell", type: "string", demandOption: true,
      describe: "Which Cell to start reading data from.",
    })
    .option("m", {
      alias: "month", type: "string", demandOption: true,
      describe: "Month that will be part of the output files.",
    })
    .option("y", {
      alias: "year", type: "string", demandOption: true,
      describe: "Year that will be part of the output files.",
    })
    .option("o", {
      alias: "outDir", type: "string", demandOption: true,
      describe: "Output folder to put the results.",
    })
    .argv;
}
