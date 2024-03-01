const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Caminhos para a pasta de origem e destino
const pastaOrigem = 'C:/Users/timed/Documents/PDFS RAIZ';
const pastaDestino = 'C:/Users/timed/Desktop/PDF FINAL';
const caminhoArquivoExcel = 'C:/Users/timed/Downloads/relatorio.xlsx';

// ...

// Função para ler o arquivo Excel e obter os códigos
function obterCodigosDoExcel() {
  const workbook = xlsx.readFile(caminhoArquivoExcel);
  const nomePrimeiraPlanilha = workbook.SheetNames[0];
  const planilha = workbook.Sheets[nomePrimeiraPlanilha];
  const celulas = xlsx.utils.sheet_to_json(planilha);

  console.log('Conteúdo lido do Excel:', celulas);

  // Filtra os códigos válidos
  const codigosValidos = celulas
    .map(linha => {
      const codigoPI = `PI${linha.CodigoPi || ''}`;
      const codigoPG = `PG${linha.CodigoPi || ''}`;
      console.log('Código PI:', codigoPI);
      console.log('Código PG:', codigoPG);

      // Concatena "PI" e "PG" com o código, removendo espaços em branco e outros caracteres não alfanuméricos
      const codigoConcatenadoPI = codigoPI.replace(/[^a-zA-Z0-9]/g, '');
      const codigoConcatenadoPG = codigoPG.replace(/[^a-zA-Z0-9]/g, '');
      console.log('Código após a transformação (PI):', codigoConcatenadoPI);
      console.log('Código após a transformação (PG):', codigoConcatenadoPG);

      return [codigoConcatenadoPI, codigoConcatenadoPG];
    })
    .flat()
    .filter(codigo => /^(PI|PG)\d+$/i.test(codigo));

  console.log('Códigos válidos:', codigosValidos);

  return codigosValidos;
}

// Função para verificar se um arquivo é um PDF e tem o formato esperado
function isArquivoValido(arquivo) {
  return arquivo.toLowerCase().endsWith('.pdf') && /^(PI|PG)\d+$/i.test(arquivo);
}

// Função para mover arquivos PDF com base nos códigos do Excel
function moverArquivosComBaseNoExcel() {
  const codigosDoExcel = obterCodigosDoExcel();
  let arquivosMovidos = 0;

  // Lê os arquivos da pasta de origem
  fs.readdir(pastaOrigem, (err, arquivos) => {
    if (err) {
      console.error('Erro ao ler a pasta de origem:', err);
      return;
    }

    // Filtra apenas os arquivos válidos
    const arquivosValidos = arquivos.filter(isArquivoValido);

    console.log('Arquivos válidos:', arquivosValidos);

    // Itera sobre cada arquivo válido
    arquivosValidos.forEach(arquivo => {
      const codigoDoArquivo = arquivo.replace(/^(PI|PG)(\d+).pdf$/, '$2'); // Extrai o código do arquivo
      if (codigosDoExcel.includes(`PI${codigoDoArquivo}`) || codigosDoExcel.includes(`PG${codigoDoArquivo}`)) {
        const caminhoOrigem = path.join(pastaOrigem, arquivo);
        const caminhoDestino = path.join(pastaDestino, arquivo);

        // Move o arquivo (recorta e cola)
        fs.renameSync(caminhoOrigem, caminhoDestino);

        console.log(`Arquivo movido: ${arquivo}`);
        arquivosMovidos++;
      }
    });

    // Verifica quantos arquivos foram movidos
    if (arquivosMovidos > 0) {
      console.log(`${arquivosMovidos} arquivos foram movidos com sucesso.`);
    } else {
      console.log('Nenhum arquivo foi movido.');
    }
  });
}

// Chama a função para mover os arquivos PDF com base nos códigos do Excel
moverArquivosComBaseNoExcel();