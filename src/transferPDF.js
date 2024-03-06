const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Paths
const sourceFolder = 'path';
const destinationFolder = 'path;
const excelFilePath = 'path/relatorio.xlsx';

// Funcao para ler o excel e pegar os codigos
function getCodesFromExcel() {
  const workbook = xlsx.readFile(excelFilePath);
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  const cells = xlsx.utils.sheet_to_json(sheet);

  // Valida os Codigos
  const validCodes = cells
    .map(row => {
      const codePI = `PI${row.PI || ''}`;
      const codePG = `PG${row.PI || ''}`;

      // Concatena "Pi" e "PG" com os codigos
      const concatenatedCodePI = codePI.replace(/[^a-zA-Z0-9]/g, '');
      const concatenatedCodePG = codePG.replace(/[^a-zA-Z0-9]/g, '');

      return [concatenatedCodePI, concatenatedCodePG];
    })
    .flat()
    .filter(code => /^(PI|PG)\d+$/i.test(code));

  return validCodes;
}

// Funcao para chegar se o arquivos são pdfs
function isFileValid(file) {
  const lowerCaseFile = file.toLowerCase();
  return (lowerCaseFile.endsWith('.pdf') || lowerCaseFile.endsWith('.PDF')) && /^(PI|PG)\d+\.pdf$/i.test(file);
}

// Funcao para mover o pdf baseados nos codigos do excel
function moveFilesBasedOnExcel() {
  const excelCodes = getCodesFromExcel();
  let movedFilesCount = 0;
 
  // Lê o arquivo da pasta source
  fs.readdir(sourceFolder, (err, files) => {
    if (err) {
      console.error('Error reading the source folder:', err);
      return;
    }

    // Filtra apenas arquivos validos
    const validFiles = files.filter(isFileValid);

    console.log('Valid Files:', validFiles);
    
    validFiles.forEach(file => {
      const fileCode = file.replace(/^(PI|PG)(\d+).pdf$/i, '$2'); 

      // Checha se o codigo do arquivo é igual ao do excel
      if (excelCodes.includes(`PI${fileCode}`) || excelCodes.includes(`PG${fileCode}`)) {
        const sourcePath = path.join(sourceFolder, file);
        const destinationPath = path.join(destinationFolder, file);

        // Recorta e cola na pasta de destino
        fs.renameSync(sourcePath, destinationPath);

        console.log(`File moved: ${file}`);
        movedFilesCount++;
      }
    });

    // Verifica quantos arquivos foram movidos
    if (movedFilesCount > 0) {
      console.log(`${movedFilesCount} files were moved successfully.`);
    } else {
      console.log('No files were moved.');
    }
  });
}

moveFilesBasedOnExcel();

