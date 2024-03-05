const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Paths
const sourceFolder = 'path';
const destinationFolder = 'path;
const excelFilePath = 'path/relatorio.xlsx';

// Function to read the Excel file and get the codes
function getCodesFromExcel() {
  const workbook = xlsx.readFile(excelFilePath);
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  const cells = xlsx.utils.sheet_to_json(sheet);

  // Filter valid codes
  const validCodes = cells
    .map(row => {
      const codePI = `PI${row.PI || ''}`;
      const codePG = `PG${row.PI || ''}`;

      // Concatenate "PI" and "PG" with the codes, removing whitespace and other characters
      const concatenatedCodePI = codePI.replace(/[^a-zA-Z0-9]/g, '');
      const concatenatedCodePG = codePG.replace(/[^a-zA-Z0-9]/g, '');

      return [concatenatedCodePI, concatenatedCodePG];
    })
    .flat()
    .filter(code => /^(PI|PG)\d+$/i.test(code));

  return validCodes;
}

// Function to check if the files are PDFs
function isFileValid(file) {
  const lowerCaseFile = file.toLowerCase();
  return (lowerCaseFile.endsWith('.pdf') || lowerCaseFile.endsWith('.PDF')) && /^(PI|PG)\d+\.pdf$/i.test(file);
}

// Function to move the PDF and files based on the Excel PIs
function moveFilesBasedOnExcel() {
  const excelCodes = getCodesFromExcel();
  let movedFilesCount = 0;
 
  // Read files from the source folder
  fs.readdir(sourceFolder, (err, files) => {
    if (err) {
      console.error('Error reading the source folder:', err);
      return;
    }

    // Filter only valid files
    const validFiles = files.filter(isFileValid);

    console.log('Valid Files:', validFiles);
    
    validFiles.forEach(file => {
      const fileCode = file.replace(/^(PI|PG)(\d+).pdf$/i, '$2'); // Extract the code

      // Check if the file's code is equal to the one in Excel
      if (excelCodes.includes(`PI${fileCode}`) || excelCodes.includes(`PG${fileCode}`)) {
        const sourcePath = path.join(sourceFolder, file);
        const destinationPath = path.join(destinationFolder, file);

        // Move the file (cut and paste)
        fs.renameSync(sourcePath, destinationPath);

        console.log(`File moved: ${file}`);
        movedFilesCount++;
      }
    });

    // Check how many files were moved
    if (movedFilesCount > 0) {
      console.log(`${movedFilesCount} files were moved successfully.`);
    } else {
      console.log('No files were moved.');
    }
  });
}

moveFilesBasedOnExcel();

