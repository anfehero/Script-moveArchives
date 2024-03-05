const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Paths
const sourceFolder = 'path';
const destinationFolder = 'path';
const excelFilePath = 'path/relatorio.xlsx';

// Funcao para ler o excel e pegar os codigos
function getCodesFromExcel() {
  const workbook = xlsx.readFile(excelFilePath);
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  const cells = xlsx.utils.sheet_to_json(sheet);

  // Filtra os codigos validos
  const validCodes = cells.map(row => {
    const codeImage = `F${row.codigo}`; // Adiciona o 'F' no prefixo do codigo do excel
    return codeImage;
  });

  return validCodes;
}

// Função chea se o arquivos é uma imagem e se tem o formato correto
function isImageValid(file) {
  const excelImageCodes = getCodesFromExcel();

  // Extrai o codigo da imagem do nome do arquivo
  const imageCode = file.match(/^(f|F)\d{3,6}-?\d{2,3}-\d{3}/);

  // Checa se o arquivos é uma imagem e se o codigo esta no excel
  return imageCode && excelImageCodes.includes(imageCode[0].toUpperCase());
}

function moveImagesBasedOnExcel() {
  let movedImagesCount = 0;

  // Lê os arquivos da pasta raiz
  fs.readdir(sourceFolder, (err, files) => {
    if (err) {
      console.error('Erro ao ler a pasta RAIZ da imagem: ', err);
      return;
    }

    // Filtra apenas as imagens validas
    const validImages = files.filter(isImageValid);

    // Itera cada imagem valida
    validImages.forEach(image => {
      const sourceImagePath = path.join(sourceFolder, image);
      const destinationImagePath = path.join(destinationFolder, image);

      // Move as imagens (recorte e cola)
      fs.renameSync(sourceImagePath, destinationImagePath);

      movedImagesCount++;
    });

    // Checa quantas imagens foi movida
    if (movedImagesCount > 0) {
      console.log(`${movedImagesCount} As imagens foram movidas com sucesso.`);
    } else {
      console.log('Nenhuma imagem movida.');
    }
  });
}

moveImagesBasedOnExcel();
