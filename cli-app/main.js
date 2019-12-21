const buildOptions = require('./config-options.module');
const fileGenerator = require('../src/main');

// Configure and get cli options
const options = buildOptions();
fileGenerator(options);
