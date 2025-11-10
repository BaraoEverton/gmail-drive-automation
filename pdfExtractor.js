
const CONFIG = {
  // Aba onde você grava os dados do PDF (no MESMO spreadsheet ativo)
  OUTPUT_SHEET_NAME: 'DOCUMENTOS',

  // Cabeçalho da faixa H1:L1 (exatamente 5 colunas)
  OUTPUT_HEADER: ['VALOR', 'VENCIMENTO', 'CODIGO DO BOLETO', 'Nº DOCUMENTO', 'REGISTRO'],

  // Planilha EXTERNA para pilha de envios (onde faz appendRow)
  EXTERNAL_SPREADSHEET_ID: 'PASTE_SPREADSHEET_ID',
  EXTERNAL_SHEET_NAME: 'OUTBOX',
};

// --- ACESSA DADOS DA PLANILHA (ATUAL) ---
const planilhaDasPdf = SpreadsheetApp.getActiveSpreadsheet();
const abaDasPdf = planilhaDasPdf.getSheetByName(CONFIG.OUTPUT_SHEET_NAME);

// --- ACESSA PLANILHA EXTERNA DE FLUXOS ---
let planilhaFluxos = SpreadsheetApp.openById(CONFIG.EXTERNAL_SPREADSHEET_ID);
let abaFluxos = planilhaFluxos.getSheetByName(CONFIG.EXTERNAL_SHEET_NAME);
let intervaloFluxos = abaFluxos.getRange('A:J').getValues();
const ultimaLinhaFluxos = abaFluxos.getLastRow();
Logger.log(ultimaLinhaFluxos);

// ----------------------

function cabecalhoPdfDas() {
  const controleAbaPdf = abaDasPdf.getRange('H1').getValue();
  if (controleAbaPdf === '') {
    // Escreve cabeçalho em H1:L1 (5 colunas)
    abaDasPdf.getRange(1, 8, 1, CONFIG.OUTPUT_HEADER.length).setValues([CONFIG.OUTPUT_HEADER]);
    Logger.log('Cabeçalho criado com sucesso');
  }
}

function extrairInfoPdf() {
  // Cria cabeçalho se necessário
  cabecalhoPdfDas();

  // ÚLTIMA LINHA RECEBIDA
  const intervaloExtracao = abaDasPdf.getLastRow();
  const ultimLinhaExtraida = abaDasPdf.getRange(intervaloExtracao, 7, 1, 1).getValue();   // Coluna G
  const ultimLinhaRegistrada = abaDasPdf.getRange(intervaloExtracao, 12, 1, 1).getValue(); // Coluna L

  Logger.log('Dados última linha (G): ' + ultimLinhaExtraida);
  Logger.log('Dados última linha (L): ' + ultimLinhaRegistrada);

  if (ultimLinhaExtraida !== ultimLinhaRegistrada) {
    // EXTRAI O LINK DA TABELA (coluna D)
    const intervaloLinks = abaDasPdf.getRange(intervaloExtracao, 4).getValue();
    Logger.log(intervaloLinks);

    // CONVERTE O LINK EM ID
    const regex = /\/d\/([^/]+)/;
    const matchLink = intervaloLinks.match(regex)[1];
    Logger.log(matchLink);

    // COPIA PDF → DOC (OCR) via Drive API avançado
    const docFileDas = Drive.Files.copy(
      { title: 'TEMP_OCR_', mimeType: 'application/vnd.google-apps.document' },
      matchLink,
      { ocr: true }
    );

    // PEGA DADOS DO NOVO DOCUMENTO
    const docFileId = docFileDas.id;
    Logger.log(docFileId);
    const docDas = DocumentApp.openById(docFileId);
    const texto = docDas.getBody().getText();

    // RASTREIA DADOS IMPORTANTES (regex mantidos)
    const valorDas = texto.match(/Valor:\s*([\d.,/-]+)/)[1];
    const vencimentoDas = texto.match(/Pagar\s*at[eé]\s*:\s*(\d{2}\/\d{2}\/\d{4})/i)[1];
    const nDocDas = texto.match(/N[úu]mero\s*:\s*(\S+)/i)[1];
    const numBoletoDas1 = texto.match(/(\d{10,})(?:\s+(\d.))?/g)[0];
    const numBoletoDas2 = texto.match(/(\d{10,})(?:\s+(\d.))?/g)[1];
    const numBoletoDas3 = texto.match(/(\d{10,})(?:\s+(\d.))?/g)[2];
    const numBoletoDas4 = texto.match(/(\d{10,})(?:\s+(\d.))?/g)[3];
    const numBoletoInteiro = numBoletoDas1 + numBoletoDas2 + numBoletoDas3 + numBoletoDas4;

    Logger.log(valorDas);
    Logger.log(vencimentoDas);
    Logger.log(nDocDas);
    Logger.log(numBoletoInteiro);

    // ESCREVE NA FAIXA H:L DA ÚLTIMA LINHA
    abaDasPdf
      .getRange(intervaloExtracao, 8, 1, 5)
      .setValues([[valorDas, vencimentoDas, numBoletoInteiro, nDocDas, 'TRUE']]);

    // --- ENVIA DADOS PARA A PLANILHA DE FLUXOS ---
    // OBS: mantém exatamente a chamada à variável global "aba" (definida no seu outro arquivo)
    const dadosFluxo = aba.getRange(intervaloExtracao, 1, 2, 12).getDisplayValues()[0];
    abaFluxos.appendRow(dadosFluxo);

    // EXCLUI O DOCUMENTO TEMPORÁRIO
    Drive.Files.remove(docFileId);
  } else {
    Logger.log('Extração PDF ainda não realizada');
  }
}
