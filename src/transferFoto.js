const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Caminhos para a pasta de origem e destino
const pastaOrigemImagens = 'C:/Users/timed/Documents/IMAGENS RAIZ';
const pastaDestinoImagens = 'C:/Users/timed/Desktop/IMAGENS FINAL';
const caminhoArquivoExcelImagens = 'C:/Users/timed/Downloads/relatorio_imagens.xlsx';

// Função para ler o arquivo Excel e obter os códigos
function obterCodigosDoExcelImagens() {
  const workbook = xlsx.readFile(caminhoArquivoExcelImagens);
  const nomePrimeiraPlanilha = workbook.SheetNames[0];
  const planilha = workbook.Sheets[nomePrimeiraPlanilha];
  const celulas = xlsx.utils.sheet_to_json(planilha);

  console.log('Conteúdo lido do Excel para imagens:', celulas);

  // Filtra os códigos válidos
  const codigosValidos = celulas
    .map(linha => {
      const codigoImagem = `f${linha.Codigo || ''}`;
      console.log('Código da imagem:', codigoImagem);

      // Concatena "f" com o código, removendo espaços em branco e outros caracteres não alfanuméricos
      const codigoConcatenado = codigoImagem.replace(/[^a-zA-Z0-9]/g, '');
      console.log('Código após a transformação:', codigoConcatenado);

      return codigoConcatenado;
    })
    .filter(codigo => /^f\d+$/i.test(codigo));

  console.log('Códigos válidos para imagens:', codigosValidos);

  return codigosValidos;
}

// Função para verificar se um arquivo é uma imagem e tem o formato esperado
function isImagemValida(arquivo) {
  return /\.(jpg|jpeg|png)$/i.test(arquivo);
}

// Função para mover arquivos de imagem com base nos códigos do Excel
function moverImagensComBaseNoExcel() {
  const codigosDoExcelImagens = obterCodigosDoExcelImagens();
  let imagensMovidas = 0;

  // Lê os arquivos da pasta de origem de imagens
  fs.readdir(pastaOrigemImagens, (err, arquivos) => {
    if (err) {
      console.error('Erro ao ler a pasta de origem de imagens:', err);
      return;
    }

    // Filtra apenas os arquivos de imagem válidos
    const imagensValidas = arquivos.filter(isImagemValida);

    console.log('Imagens válidas:', imagensValidas);

    // Itera sobre cada imagem válida
    imagensValidas.forEach(imagem => {
      const codigoDaImagem = imagem.replace(/^f(\d+)\.(jpg|jpeg|png)$/i, '$1'); // Extrai o código da imagem

      // Verifica se o código da imagem está nos códigos do Excel para imagens
      if (codigosDoExcelImagens.includes(`f${codigoDaImagem}`)) {
        const caminhoOrigemImagem = path.join(pastaOrigemImagens, imagem);
        const caminhoDestinoImagem = path.join(pastaDestinoImagens, imagem);

        // Move a imagem (recorta e cola)
        fs.renameSync(caminhoOrigemImagem, caminhoDestinoImagem);

        console.log(`Imagem movida: ${imagem}`);
        imagensMovidas++;
      }
    });

    // Verifica quantas imagens foram movidas
    if (imagensMovidas > 0) {
      console.log(`${imagensMovidas} imagens foram movidas com sucesso.`);
    } else {
      console.log('Nenhuma imagem foi movida.');
    }
  });
}

// Chama a função para mover as imagens com base nos códigos do Excel para imagens
moverImagensComBaseNoExcel();
