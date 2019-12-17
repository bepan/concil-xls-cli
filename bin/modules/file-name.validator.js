const fs = require('fs');
const path = require('path');

class FileNameValidator
{
  run(fileName, desktopPath)
  {
    // Check if file name is empty
    if (fileName.trim() === '')
    {
      console.error('The provided file name is empty.');
      process.exit(1);
    }

    // Check if file exists in Desktop
    if ( !fs.existsSync(path.join(desktopPath, fileName)) ) 
    {
      console.error('The provided file is not placed in your Desktop.');
      process.exit(1);
    }
  }
}

module.exports = new FileNameValidator();
