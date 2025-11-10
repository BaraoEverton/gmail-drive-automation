
const CONFIG = {
  // Ex.: 'from:(*@empresa.com)' ou 'from:(pessoa@empresa.com)'
  GMAIL_SEARCH_QUERY: 'from:(*@exemplo.com)',

  // Palavra-chave do assunto para filtrar os e-mails (ex.: 'DAS', 'DARF', etc.)
  SUBJECT_KEYWORD: 'ASSUNTO EMAIL',

  // Nome da aba (sheet) onde os dados serão escritos
  SHEET_NAME: 'DOCUMENTOS',

  // Pasta de destino no Drive (ID). Use uma pasta técnica do seu projeto.
  DRIVE_FOLDER_ID: 'COLOQUE_AQUI_O_ID_DA_PASTA',

  // Formato de data para registro na planilha
  DATE_FORMAT: 'dd/MM/yyyy',
};

// --- ACESSO A DADOS DO GMAIL ---
const threadsSelecionadas = GmailApp.search(CONFIG.GMAIL_SEARCH_QUERY);

// --- ACESSO A DADOS DA PLANILHA ---
let planilha = SpreadsheetApp.getActiveSpreadsheet();
let aba = planilha.getSheetByName(CONFIG.SHEET_NAME);
let intervalo = aba.getRange('A1:F');
let timezoneForm = Session.getScriptTimeZone();

// --- ACESSO AO DIRETÓRIO NO DRIVE ---
const pastaDestino = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID); // salva na pasta configurada

// --- CABEÇALHO PADRÃO ---
function criaCabecalho() {
  const cabecalho = [
    'ID',
    'DATA EMAIL',
    'NOME',
    'LINK DO PDF',
    'DESTINATÁRIO',
    'REMETENTE',
    'COMPROVANTE',
    'RECEBIMENTO',
  ];

  const celula = aba.getRange('A1');
  const verificacao = celula.getValue();

  if (verificacao === '') {
    // LINHA INICIAL, COLUNA INICIAL, NUMERO DE LINHAS, NUMERO DE COLUNAS
    aba.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]);
  } else {
    Logger.log('Cabeçalho já criado.');
  }
} // FIM CABEÇALHO

function iteraThreads() {
  criaCabecalho();

  // --- 1. PEGA A DATA E OS IDs DA ÚLTIMA EXTRAÇÃO ---
  const ultimaLinha = aba.getLastRow();

  // Se já tem dados, pega a última data registrada na coluna B
  let ultimaExtracaoData = null;
  if (ultimaLinha > 1) {
    ultimaExtracaoData = aba.getRange(ultimaLinha, 2).getValue(); // coluna B
    Logger.log('Última data registrada: ' + ultimaExtracaoData);
  }

  // Pega todos os IDs existentes (coluna A, a partir da linha 2)
  let idsExistentes = [];
  if (ultimaLinha > 1) {
    idsExistentes = aba
      .getRange(2, 1, ultimaLinha - 1, 1)
      .getValues()
      .flat()
      .filter(String);
  }
  const idsSet = new Set(idsExistentes);

  // --- 2. ITERA AS THREADS E MENSAGENS ---
  for (let thread of threadsSelecionadas) {
    const mensagens = thread.getMessages();

    for (let msg of mensagens) {
      const idMsg = msg.getId();
      const dataEmail = msg.getDate();
      const assunto = msg.getSubject();
      const remetente = msg.getFrom();
      const destinatario = msg.getTo();
      const anexos = msg.getAttachments();
      const dataTratada = Utilities.formatDate(
        dataEmail,
        timezoneForm,
        CONFIG.DATE_FORMAT
      );

      // Se já temos esse ID, pula
      if (idsSet.has(idMsg)) {
        Logger.log('🟡 ID já registrado: ' + idMsg);
        continue;
      }

      // Controle por data (mantido exatamente como no original)
      if (ultimaExtracaoData && dataEmail > ultimaExtracaoData) {
        Logger.log('⏩ E-mail fora do escopo de data: ' + dataTratada);
        continue;
      }

      // --- 3. FILTRA PELO REMETENTE/ASSUNTO ---
      // Mantém a checagem por domínio/ocorrência no endereço e palavra-chave no assunto
      if (remetente.includes('@') && assunto.includes(CONFIG.SUBJECT_KEYWORD)) {
        // --- 4. PROCESSA E SALVA ---
        for (let anexo of anexos) {
          const arquivo = pastaDestino.createFile(anexo);
          const linkArquivo = arquivo.getUrl();
          const nomeArquivo = arquivo.getName();
          const hiperlink = '=HYPERLINK("' + linkArquivo + '";"' + nomeArquivo + '")';
          const novaLinha = aba.getLastRow() + 1;

          // Observação: Mantém a MESMA ordem e quantidade de colunas definidas abaixo
          aba.getRange(novaLinha, 1, 1, 8).setValues([
            [
              idMsg,          // A: ID
              dataTratada,    // B: DATA EMAIL
              nomeArquivo,    // C: NOME
              linkArquivo,    // D: LINK DO PDF
              destinatario,   // E: DESTINATÁRIO
              remetente,      // F: REMETENTE
              '',             // G: COMPROVANTE
              'TRUE',         // H: RECEBIMENTO
            ],
          ]);

          Logger.log('✅ Documento registrado: ' + nomeArquivo);

          // Mantido exatamente como no original (se existir no seu projeto)
          extrairInfoPdf();
        }
      }
    }
  }
}
