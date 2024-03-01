const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Paths
const sourceFolder = 'C:/Users/timed/Documents/PDFS RAIZ';
const destinationFolder = 'C:/Users/timed/Desktop/PDF FINAL';
const excelFilePath = 'C:/Users/timed/Downloads/relatorio.xlsx';
const wordFolder = 'C:/Users/timed/Desktop/WORD FINAL';

//Funcao para ler o arquivo excel e os codigos
function getCodesFromExcel() {
  const workbook = xlsx.readFile(excelFilePath);
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  const cells = xlsx.utils.sheet_to_json(sheet);

  // Filtra codigos validos
  const validCodes = cells
    .map(row => {
      const codePI = `PI${row.PI || ''}`;
      const codePG = `PG${row.PI || ''}`;

      //Concatena "PI" e "PG" com os codigos, removendo espaçoes em branco e outros caracteres
      const concatenatedCodePI = codePI.replace(/[^a-zA-Z0-9]/g, '');
      const concatenatedCodePG = codePG.replace(/[^a-zA-Z0-9]/g, '');

      return [concatenatedCodePI, concatenatedCodePG];
    })
    .flat()
    .filter(code => /^(PI|PG)\d+$/i.test(code));

  return validCodes;
}

// Função para checar se os arquivos são pdfs
function isFileValid(file) {
  return file.toLowerCase().endsWith('.pdf') && /^(PI|PG)\d+\.pdf$/i.test(file);
}

// Funcao para mover o pdf e arquivos baseados nos PIs do Excel
function moveFilesBasedOnExcel() {
  const excelCodes = getCodesFromExcel();
  let movedFilesCount = 0;
 
  // Lê os arquivos da pasta RAIZ
  fs.readdir(sourceFolder, (err, files) => {
    if (err) {
      console.error('Erro ao ler a pasta RAIZ:', err);
      return;
    }

    // Filtra apenas os arquivos validos
    const validFiles = files.filter(isFileValid);

    console.log('Arquivos Validos:', validFiles);
    
    validFiles.forEach(file => {
      const fileCode = file.replace(/^(PI|PG)(\d+).pdf$/, '$2'); // Extrai o codigo

      // Checa se o PI do arquivo é igual ao do Excel
      if (excelCodes.includes(`PI${fileCode}`) || excelCodes.includes(`PG${fileCode}`)) {
        const sourcePath = path.join(sourceFolder, file);
        const destinationPath = path.join(destinationFolder, file);

        // Move o arquivo (recorta e cola)
        fs.renameSync(sourcePath, destinationPath);

        console.log(`Arquivo movido: ${file}`);
        movedFilesCount++;
      }
    });

    // Checa quantos arquivos foram movidos
    if (movedFilesCount > 0) {
      console.log(`${movedFilesCount} Arquivos movidos com sucesso`);
    } else {
      console.log('Nenhum arquivo movido.');
    }
  });
}

moveFilesBasedOnExcel();


// Function to check if a Word file with the same name exists in another folder
function doesWordFileExist(fileName, wordFolder) {
  const wordFilePath = path.join(wordFolder, fileName.replace(/\.pdf$/, '.docx'));

  return fs.existsSync(wordFilePath);
}

// Function to move PDF files based on Excel codes and check for corresponding Word files
function moveAndCheckWordFiles() {
  const excelCodes = getCodesFromExcel();
  let movedFilesCount = 0;

  // Read files from the source folder
  fs.readdir(sourceFolder, (err, files) => {
    if (err) {
      console.error('Error reading source folder:', err);
      return;
    }

    // Filter only valid files
    const validFiles = files.filter(isFileValid);

    console.log('Valid files:', validFiles);

    // Iterate over each valid file
    validFiles.forEach(file => {
      const fileCode = file.replace(/^(PI|PG)(\d+).pdf$/, '$2'); // Extract the file code

      // Check if the file code is in the Excel codes
      if (excelCodes.includes(`PI${fileCode}`) || excelCodes.includes(`PG${fileCode}`)) {
        const sourcePath = path.join(sourceFolder, file);
        const destinationPath = path.join(destinationFolder, file);

        // Move the PDF file (cut and paste)
        fs.renameSync(sourcePath, destinationPath);

        if (doesWordFileExist(file, wordFolder)) {
          console.log(`Word file found for ${file}`);
          // Add your logic to move the Word file to another folder if needed
          // Example: fs.renameSync(wordFilePath, path.join(anotherFolder, wordFileName));
        } else {
          console.log(`No Word file found for ${file}`);
        }

        console.log(`Word file moved: ${file}`);
        movedFilesCount++;
      }
    });

    // Check how many files were moved
    if (movedFilesCount > 0) {
      console.log(`${movedFilesCount}Word files were successfully moved.`);
    } else {
      console.log('No Word file was moved.');
    }
  });
}

// Call the extended function to move PDF files, check for corresponding Word files, and take additional actions
moveAndCheckWordFiles();