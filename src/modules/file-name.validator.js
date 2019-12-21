const fs = require('fs');

class FileNameValidator
{
  run(file)
  {
    // Check if file name is empty
    if (file.trim() === '')
    {
      console.error('The provided file path is empty.');
      process.exit(1);
    }

    // Check if file exists
    if ( !fs.existsSync(file) ) 
    {
      console.error('The provided file does not exists.');
      process.exit(1);
    }
  }
}

module.exports = new FileNameValidator();
