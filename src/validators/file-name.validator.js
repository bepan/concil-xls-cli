const fs = require('fs');

class FileNameValidator
{
  run(file)
  {
    // Check if file name is empty
    if (file.trim() === '')
    {
      throw new Error('El Archivo Base no puede estar vacio.');
    }

    // Check if file exists
    if ( !fs.existsSync(file) ) 
    {
      throw new Error('El Archivo Base no existe');
    }
  }
}

module.exports = new FileNameValidator();
