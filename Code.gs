
/**
 * GERADOR DE FICHAS DE INSPEÇÃO DE SAÚDE
 * Desenvolvido para integração com APIs externas de Fichas e Inspecionados.
 */

const CONFIG = {
  SS: SpreadsheetApp.getActiveSpreadsheet(),
  // API que retorna as fichas (Aba FICHAS)
  API_FICHAS: "https://script.google.com/macros/s/AKfycbzKbu1m8zsQfaqUXyufwVqNmxQRw0pHOo6H528muJ3FIm49zonO537amN309LuRhz52Dw/exec",
  // API que retorna os dados dos inspecionados (para busca de Nº do Arquivo/Prontuário)
  API_INSPECIONADOS: "https://script.google.com/macros/s/AKfycbzydoxvSMbWIJvCktWSdzEhP6g5dC_Wh7e5PqID-D1qJCMNVe8cnGCVrnZwWdtJH20/exec"
};

/**
 * Função principal para servir a interface ou responder como API JSON.
 */
function doGet(e) {
  // Se o parâmetro "action" for "data", retorna o JSON das inspeções do dia
  if (e && e.parameter && e.parameter.action === 'data') {
    const data = getPendingInspections();
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Caso contrário, serve a interface HTML do React
  return HtmlService.createHtmlOutputFromFile('index.html')
    .setTitle('Gerador de Fichas de Saúde')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Helper para buscar dados de APIs externas via GET.
 */
function fetchFromApi(url) {
  try {
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      followRedirects: true,
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) return null;
    const content = response.getContentText();
    return JSON.parse(content);
  } catch (err) {
    Logger.log("Erro ao acessar API: " + url + " - " + err.message);
    return null;
  }
}

/**
 * Busca as fichas da API externa e filtra somente as do dia atual.
 */
function getPendingInspections() {
  const result = fetchFromApi(CONFIG.API_FICHAS);
  const rows = Array.isArray(result) ? result : (result?.data || []);
  
  if (rows.length === 0) return { inspections: [], stats: { totalFichas: 0, uniqueInspecionandos: 0, homens: 0, mulheres: 0 }, printUrl: "" };

  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  const uniqueCpfs = new Set();
  let homens = 0;
  let mulheres = 0;

  // O cabeçalho é removido e os dados mapeados
  const inspections = rows.slice(1).map((row, index) => {
    const dtInsp = formatValue(row[0]);
    
    // Filtro rigoroso: Somente fichas do dia
    if (dtInsp !== todayStr) return null;

    const cpf = String(row[6] || "").replace(/\D/g, '').padStart(11, '0');
    const sexo = String(row[7] || "").toUpperCase();
    const dtNasc = row[4] instanceof Date ? row[4] : new Date(row[4]);
    
    let age = 0;
    if (!isNaN(dtNasc.getTime())) {
      const today = new Date();
      age = today.getFullYear() - dtNasc.getFullYear();
      const m = today.getMonth() - dtNasc.getMonth();
      if (m < 0 || (m === 0 && today.getDate() < dtNasc.getDate())) age--;
    }

    if (cpf) uniqueCpfs.add(cpf);
    if (sexo.includes("MASC") || sexo === "M") homens++;
    else if (sexo.includes("FEM") || sexo === "F") mulheres++;

    return {
      dtInsp: dtInsp,
      codInsp: row[1],
      rg: row[2],
      dtNascimento: formatValue(row[4]),
      cpf: cpf,
      sexo: sexo,
      nome: row[11],
      om: row[13],
      posto: row[8],
      quadro: row[9],
      especialidade: row[10],
      dtPraca: formatValue(row[14]),
      originalIndex: index + 2, // Índice da linha original para processamento
      vinculo: row[15],
      finalidade: row[16],
      grupo: row[20],
      idade: age,
      controle: row[41] || row[37] || ""
    };
  }).filter(item => item !== null);

  return {
    inspections,
    stats: {
      totalFichas: inspections.length,
      uniqueInspecionandos: uniqueCpfs.size,
      homens: homens,
      mulheres: mulheres
    },
    printUrl: getPrintUrl()
  };
}

/**
 * Gera as fichas na aba IMPRESSAO local.
 */
function generateFichasAction(selectedIndices) {
  const ss = CONFIG.SS;
  const sheetTemplate = ss.getSheetByName("TEMPLATE");
  const sheetImpressao = ss.getSheetByName("IMPRESSAO");
  
  if (!sheetTemplate || !sheetImpressao) {
    return { success: false, message: "Abas TEMPLATE ou IMPRESSAO não encontradas neste arquivo." };
  }

  sheetImpressao.clear();
  let currentRow = 1;

  // Busca dados atualizados para processamento
  const fichasApi = fetchFromApi(CONFIG.API_FICHAS);
  const fichas = Array.isArray(fichasApi) ? fichasApi : (fichasApi?.data || []);
  
  const inspecionadosApi = fetchFromApi(CONFIG.API_INSPECIONADOS);
  const inspecionados = Array.isArray(inspecionadosApi) ? inspecionadosApi : (inspecionadosApi?.data || []);

  // Cache para busca rápida do Nº do Arquivo (CPF na Coluna E[4], Arquivo na Coluna F[5])
  const archiveCache = {};
  inspecionados.forEach(row => {
    const cpf = String(row[4] || "").replace(/\D/g, '').padStart(11, '0');
    const arquivo = row[5];
    if (cpf && arquivo) archiveCache[cpf] = arquivo;
  });

  selectedIndices.sort((a, b) => a - b).forEach(idx => {
    const rowData = fichas[idx - 1]; // Ajuste para 0-indexed do array
    if (rowData) {
      populateTemplate(rowData, sheetTemplate, sheetImpressao, currentRow, archiveCache);
      currentRow += 64;
    }
  });

  return {
    success: true,
    message: `${selectedIndices.length} fichas geradas com sucesso!`,
    printUrl: getPrintUrl()
  };
}

/**
 * Preenche o template e copia para a aba de impressão.
 */
function populateTemplate(row, template, target, startRow, archiveCache) {
  // Limpezas prévias
  template.getRange("S4").clearContent(); 
  template.getRange("Q4").clearContent(); 
  template.getRange("N1").clearContent();

  const cpf = String(row[6] || "").replace(/\D/g, '').padStart(11, '0');
  const dtNasc = row[4] instanceof Date ? row[4] : new Date(row[4]);
  const dtPraca = row[14] instanceof Date ? row[14] : new Date(row[14]);
  const om = row[13];
  const posto = row[8];
  const quadro = row[9];
  const espec = row[10];

  // Cálculos Automáticos
  const today = new Date();
  if (!isNaN(dtNasc.getTime())) {
    let age = today.getFullYear() - dtNasc.getFullYear();
    const m = today.getMonth() - dtNasc.getMonth();
    if (m < 0 || (m === 0 && today.getDate() < dtNasc.getDate())) age--;
    template.getRange("A11").setValue(age);
  }

  if (!isNaN(dtPraca.getTime())) {
    template.getRange("A15").setValue(calculateServiceTime(dtPraca, today));
  }

  // Preenchimento de Campos
  template.getRange("H4").setValue(row[1]); // Cod Insp
  template.getRange("B11").setValue(dtNasc);
  template.getRange("L11").setValue(row[5]); // Naturalidade
  template.getRange("G15").setValue(cpf);
  template.getRange("F11").setValue(row[7]); // Sexo
  template.getRange("Q9").setValue(`${posto} ${quadro} ${espec}`.trim());
  template.getRange("A9").setValue(row[11]); // Nome
  template.getRange("O8").setValue(row[12]); // Saram
  template.getRange("M15").setValue(om);
  template.getRange("D6").setValue("Letra(s) " + (row[15] || ""));
  template.getRange("Q11").setValue(row[17]); // Email
  template.getRange("A13").setValue(row[18]); // Endereço
  template.getRange("P15").setValue(row[19]); // Telefone
  template.getRange("H2").setValue(row[20]);  // Grupo
  template.getRange("G11").setValue(row[22]); // Cor

  // Nacionalidade baseada na naturalidade
  const naturalidade = String(row[5]).toUpperCase();
  const estados = ["AM", "AC", "AL", "AP", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RO", "RS", "RR", "SC", "SP", "SE", "TO"];
  template.getRange("I11").setValue(estados.includes(naturalidade) ? "BRASILEIRA" : "ESTRANGEIRA");

  // Regras de OM (SGPO / BAVEX)
  const fullSpec = `${posto} ${quadro} ${espec}`.toUpperCase();
  const keywords = ["BCT", "BCO", "CTA", "PTA", "OEA", "ATCO"];
  const isSgpo = keywords.some(k => fullSpec.includes(k));
  
  if (om === "4 BAVEX") template.getRange("N1").setValue("4 BAVEX");
  else if (isSgpo) template.getRange("N1").setValue("SGPO");

  // Nº do Arquivo (Prontuário)
  if (archiveCache[cpf]) {
    template.getRange("S4").setValue(archiveCache[cpf]);
  }

  // Cópia para Impressão
  template.getRange("A1:T64").copyTo(target.getRange(startRow, 1));
}

function calculateServiceTime(start, end) {
  let years = end.getFullYear() - start.getFullYear();
  let months = end.getMonth() - start.getMonth();
  let days = end.getDate() - start.getDate();
  if (days < 0) {
    months--;
    const lastMonth = new Date(end.getFullYear(), end.getMonth(), 0);
    days += lastMonth.getDate();
  }
  if (months < 0) {
    years--;
    months += 12;
  }
  let res = [];
  if (years > 0) res.push(`${years} ano(s)`);
  if (months > 0) res.push(`${months} mes(es)`);
  if (days > 0) res.push(`${days} dia(s)`);
  return res.join(", ") || "0 dias";
}

function formatValue(val) {
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
  return val ? String(val) : "";
}

function getPrintUrl() {
  const ss = CONFIG.SS;
  const sheet = ss.getSheetByName("IMPRESSAO");
  return sheet ? `${ss.getUrl()}#gid=${sheet.getSheetId()}` : ss.getUrl();
}
