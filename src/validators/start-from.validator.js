const fs = require('fs');
const path = require('path');

class StartFromValidator
{
  run(startFrom)
  {
    if (startFrom.trim() === '')
    {
      throw new Error('La Celda Inicial no puede estar vacia.');
    }
    
    if ( !/[A-Za-z]/.test(startFrom[0]) )
    {
      throw new Error('El primer caracter de la Celda Inicial debe ser una letra.');
    }
    
    if ( !/^\d+$/.test(startFrom.substr(1)) )
    {
      throw new Error('El contenido, de la Celda Inicial, despues del primer caracter debe ser un numero.');
    }
  }
}

module.exports = new StartFromValidator();
