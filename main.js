const fs = require('fs')
const ConvertService = require("./service/convert.service");


(async () => {
  const inputDocs = []
  
  for (let i = 1; i <= 10; i++) {
    const fileName = `input/doc${i}.docx`;
    inputDocs.push(fileName);
  }

  const outputDir = 'output/';

  const pdfPaths = await ConvertService.convertDocxCollectionToPdf(inputDocs, outputDir);

  const pdfBuffer = await ConvertService.combinePdfModels(pdfPaths.map(path => ({ pathPdf: path })));

  await fs.promises.writeFile('output/final.pdf', pdfBuffer);
})();
