const fs = require('fs');
const path = require('path');

class StartFromValidator
{
  run(startFrom)
  {
    if (startFrom.trim() === '')
    {
      console.error('The cell value should not be empty.');
      process.exit(1);
    }
    
    if ( !/[A-Za-z]/.test(startFrom[0]) )
    {
      console.error('The first character of provided cell is not a letter.');
      process.exit(1);
    }
    
    if ( !/^\d+$/.test(startFrom.substr(1)) )
    {
      console.error('The cell value after the first letter must be a number.');
      process.exit(1);
    }
  }
}

module.exports = new StartFromValidator();
