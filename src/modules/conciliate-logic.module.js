const sortBy = require('lodash.sortby');

module.exports = function(dataBlock) {

  const processedMap = new WeakMap();
  const matchesArr = [], pendingRegs = [];
  let oldConcept = '', iterStart = 0;
  dataBlock = sortBy(dataBlock, ['concepto']);

  for (let i = 0; i < dataBlock.length; i++) 
  {
    const currReg = dataBlock[i];
    if (processedMap.get(currReg)) continue;
    if (currReg.concepto !== oldConcept) {
      iterStart = i;
    }
    oldConcept = currReg.concepto;
    let matchFound = false;
    const amountCurrReg = currReg.cargocg ? 'cargocg' : 'abonocg';
    const amountIterReg = amountCurrReg === 'cargocg' ? 'abonocg' : 'cargocg';

    for (let j = iterStart; j < dataBlock.length; j++) 
    {
      const iterReg = dataBlock[j];
      if (processedMap.get(iterReg)) continue;
      if (iterReg.concepto !== oldConcept) break;
      if (currReg.concepto === iterReg.concepto && 
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
