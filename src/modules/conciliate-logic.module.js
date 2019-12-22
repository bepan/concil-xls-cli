const sortBy = require('lodash/sortBy');

module.exports = function(dataBlock) {

  const processedMap = new WeakMap();
  const matchesArr = [], pendingRegs = [];
  let oldConcept = '', iterStart = 0;
  dataBlock = sortBy(dataBlock, ['Concepto']);

  for (let i = 0; i < dataBlock.length; i++) 
  {
    const currReg = dataBlock[i];
    if (processedMap.get(currReg)) continue;
    if (currReg.Concepto !== oldConcept) {
      iterStart = i;
    }
    oldConcept = currReg.Concepto;
    let matchFound = false;
    const amountCurrReg = currReg.CargoCG ? 'CargoCG' : 'AbonoCG';
    const amountIterReg = amountCurrReg === 'CargoCG' ? 'AbonoCG' : 'CargoCG';

    for (let j = iterStart; j < dataBlock.length; j++) 
    {
      const iterReg = dataBlock[j];
      if (processedMap.get(iterReg)) continue;
      if (iterReg.Concepto !== oldConcept) break;
      if (currReg.Concepto === iterReg.Concepto && 
        currReg[amountCurrReg] === iterReg[amountIterReg]) 
      {
        matchFound = true;
        matchesArr.push(currReg, iterReg);
        processedMap.set(currReg, true).set(iterReg, true);
        break;
      }
    }

    if (!matchFound)
    {
      processedMap.set(currReg, true);
      pendingRegs.push(currReg);
    }
  }

  return {
    matchesArr,
    pendingRegs
  };

};
