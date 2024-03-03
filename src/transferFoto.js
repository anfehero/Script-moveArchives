const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Caminhos para a pasta de origem e destino
const pastaOrigemImagens = 'C:/Users/RealMayer/Documents/VIVO';
const pastaDestinoImagens = 'C:/Users/RealMayer/Desktop/MORTO';
const caminhoArquivoExcelImagens = 'C:/Users/RealMayer/Documents/relatorio.xlsx';

// Função para ler o arquivo Excel e obter os códigos
function obterCodigosDoExcelImagens() {
  const workbook = xlsx.readFile(caminhoArquivoExcelImagens);
  const nomePrimeiraPlanilha = workbook.SheetNames[0];
  const planilha = workbook.Sheets[nomePrimeiraPlanilha];
  const celulas = xlsx.utils.sheet_to_json(planilha);

  // Filtra os códigos válidos
  const codigosValidos = celulas
    .map(linha => {
      const codigoImagem = `F${linha.codigo}`; // Adiciona o prefixo 'F' ao código do Excel
      return codigoImagem;
    });

  return codigosValidos;
}

// Função para verificar se um arquivo é uma imagem e tem o formato esperado
function isImagemValida(arquivo) {
  const codigosDoExcelImagens = obterCodigosDoExcelImagens();

  // Extrai o código da imagem do nome do arquivo
  const codigoDaImagem = arquivo.match(/^(f|F)\d{3,6}-?\d{2,3}-\d{3}/);

  // Verifica se o arquivo é uma imagem e se o código está nos códigos do Excel
  return codigoDaImagem && codigosDoExcelImagens.includes(codigoDaImagem[0].toUpperCase());
}

function moverImagensComBaseNoExcel() {
  let imagensMovidas = 0;

  // Lê os arquivos da pasta de origem de imagens
  fs.readdir(pastaOrigemImagens, (err, arquivos) => {
    if (err) {
      console.error('Erro ao ler a pasta de origem de imagens:', err);
      return;
    }

    // Filtra apenas os arquivos de imagem válidos
    const imagensValidas = arquivos.filter(isImagemValida);

    // Itera sobre cada imagem válida
    imagensValidas.forEach(imagem => {
      const caminhoOrigemImagem = path.join(pastaOrigemImagens, imagem);
      const caminhoDestinoImagem = path.join(pastaDestinoImagens, imagem);

      // Move a imagem (recorta e cola)
      fs.renameSync(caminhoOrigemImagem, caminhoDestinoImagem);

      imagensMovidas++;
      
      console.log(`Imagem movida: ${imagem}`);
    });

    // Verifica quantas imagens foram movidas
    if (imagensMovidas > 0) {
      console.log(`${imagensMovidas} imagens foram movidas com sucesso.`);
    } else {
      console.log('Nenhuma imagem foi movida.');
    }
  });
}

moverImagensComBaseNoExcel();
