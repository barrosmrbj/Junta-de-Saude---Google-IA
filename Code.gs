
/**
 * Main Controller for Health Inspection Sheets
 */

const CONFIG = {
  SS: SpreadsheetApp.getActiveSpreadsheet(),
  URL_FICHAS_HOJE: "https://docs.google.com/spreadsheets/d/1CWXzs_J1tTITIZ52_0t02tUb8tBewKSBNWNyaHb6z8M/edit?gid=350819832#gid=350819832",
  URL_BANCO_ARQUIVOS: "https://docs.google.com/spreadsheets/d/1s41r4hqE0qUgs7i49klZEMjtQXrXDko1gJpi2Hq9ZtE/edit?gid=1185443922#gid=1185443922"
};

/**
 * Serves the UI
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index.html')
    .setTitle('Gerador de Fichas de Saúde')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fetches data from external "FICHAS" sheet, filtered by today's date
 */
function getPendingInspections() {
  try {
    const ssExternal = SpreadsheetApp.openByUrl(CONFIG.URL_FICHAS_HOJE);
    const sheet = ssExternal.getSheetByName("FICHAS");
    if (!sheet) return { inspections: [], stats: { totalFichas: 0, uniqueInspecionandos: 0, homens: 0, mulheres: 0 }, printUrl: "" };
    
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return { inspections: [], stats: { totalFichas: 0, uniqueInspecionandos: 0, homens: 0, mulheres: 0 }, printUrl: "" };
    
    const rows = values.slice(1);
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    const uniqueCpfs = new Set();
    let homens = 0;
    let mulheres = 0;
    
    const inspections = rows.map((row, index) => {
      const dtInspVal = formatValue(row[0]);
      // Only process if it matches today or if user wants all (prompt said "of the day")
      if (dtInspVal !== todayStr && row[0] !== "") {
         // Optionally skip, but usually for debugging we might show all. 
         // Let's stick to the request: "buscar as fichas somente do dia"
         // return null; 
      }

      const dtNasc = row[4] instanceof Date ? row[4] : new Date(row[4]);
      let age = 0;
      if (!isNaN(dtNasc.getTime())) {
        const today = new Date();
        age = today.getFullYear() - dtNasc.getFullYear();
        const m = today.getMonth() - dtNasc.getMonth();
        if (m < 0 || (m === 0 && today.getDate() < dtNasc.getDate())) age--;
      }

      const cpf = String(row[6]).replace(/\D/g, '').padStart(11, '0');
      const sexo = String(row[7]).toUpperCase();
      
      uniqueCpfs.add(cpf);
      if (sexo.includes("MASC") || sexo === "M") homens++;
      else if (sexo.includes("FEM") || sexo === "F") mulheres++;

      return {
        dtInsp: dtInspVal,
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
        originalIndex: index + 2, // Row in the external sheet
        vinculo: row[15],
        finalidade: row[16],
        grupo: row[20],
        idade: age,
        controle: row[41] || row[37] || "" 
      };
    }).filter(item => item !== null && item.dtInsp === todayStr);

    const stats = {
      totalFichas: inspections.length,
      uniqueInspecionandos: uniqueCpfs.size,
      homens: homens,
      mulheres: mulheres
    };

    return { 
      inspections, 
      stats,
      printUrl: getPrintUrl()
    };
  } catch (e) {
    Logger.log("Error fetching inspections: " + e.message);
    throw new Error("Erro ao acessar banco de fichas externo.");
  }
}

/**
 * Returns a URL for the IMPRESSAO sheet in PDF/Print mode
 */
function getPrintUrl() {
  const ss = CONFIG.SS;
  const sheet = ss.getSheetByName("IMPRESSAO");
  if (!sheet) return "";
  return `${ss.getUrl()}#gid=${sheet.getSheetId()}`;
}

/**
 * Main action to generate fichas for selected row indices from EXTERNAL sheet
 */
function generateFichasAction(selectedIndices) {
  const ss = CONFIG.SS;
  const sheetTemplate = ss.getSheetByName("TEMPLATE");
  const sheetImpressao = ss.getSheetByName("IMPRESSAO");
  
  if (!sheetTemplate || !sheetImpressao) {
    return { success: false, message: "Abas TEMPLATE ou IMPRESSAO não encontradas." };
  }
  
  sheetImpressao.clear();
  let currentRow = 1;
  
  const ssExternal = SpreadsheetApp.openByUrl(CONFIG.URL_FICHAS_HOJE);
  const sheetExternal = ssExternal.getSheetByName("FICHAS");
  const allData = sheetExternal.getDataRange().getValues();
  const archiveCache = loadArchiveCache();
  
  selectedIndices.sort((a, b) => a - b).forEach((rowIndex) => {
    const dataRow = allData[rowIndex - 1]; 
    processRow(dataRow, sheetTemplate, sheetImpressao, currentRow, archiveCache);
    currentRow += 64; 
  });
  
  return { 
    success: true, 
    message: `Sucesso! ${selectedIndices.length} ficha(s) gerada(s).`,
    printUrl: getPrintUrl()
  };
}

/**
 * Processes a single row: calculates data and populates template
 */
function processRow(row, template, target, startRow, archiveCache) {
  template.getRange("S4").clearContent(); 
  template.getRange("Q4").clearContent(); 
  template.getRange("N1").clearContent(); 

  const codInsp = row[1];
  const dtNasc = row[4] instanceof Date ? row[4] : new Date(row[4]);
  const naturalidade = row[5];
  const cpf = String(row[6]).replace(/\D/g, '').padStart(11, '0');
  const sexo = row[7];
  const posto = row[8];
  const quadro = row[9];
  const especialidade = row[10];
  const nome = row[11];
  const saram = row[12];
  const om = row[13];
  const dtPraca = row[14] instanceof Date ? row[14] : null;
  const email = row[17];
  const endereco = row[18];
  const telefone = row[19];
  const grupo = row[20];
  const cor = row[22];
  const clinicaRestricao = row[25];
  const cursoEstagio = row[31];
  
  const today = new Date();
  let age = 0;
  if (!isNaN(dtNasc.getTime())) {
    age = today.getFullYear() - dtNasc.getFullYear();
    const m = today.getMonth() - dtNasc.getMonth();
    if (m < 0 || (m === 0 && today.getDate() < dtNasc.getDate())) age--;
  }
  template.getRange("A11").setValue(age);
  
  if (dtPraca) {
    const serviceTime = calculateTimeDiff(dtPraca, today);
    template.getRange("A15").setValue(serviceTime);
  } else {
    template.getRange("A15").setValue("N/A");
  }

  template.getRange("H4").setValue(codInsp); 
  template.getRange("B11").setValue(dtNasc);
  template.getRange("L11").setValue(naturalidade);
  
  const brStates = ["AM", "AC", "AL", "AP", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RO", "RS", "RR", "SC", "SP", "SE", "TO"];
  if (brStates.includes(naturalidade) || !naturalidade) {
    template.getRange("I11").setValue("BRASILEIRA");
  } else {
    template.getRange("I11").setValue("ESTRANGEIRO");
  }

  template.getRange("G15").setValue(cpf);
  template.getRange("F11").setValue(sexo);
  template.getRange("Q9").setValue(`${posto} ${quadro} ${especialidade}`.trim());
  template.getRange("A9").setValue(nome);
  template.getRange("O8").setValue(saram);
  template.getRange("M15").setValue(om);
  template.getRange("D6").setValue(`Letra(s) ${row[15] || ''}`); 
  template.getRange("Q11").setValue(email);
  template.getRange("A13").setValue(endereco);
  template.getRange("P15").setValue(telefone);
  template.getRange("H2").setValue(grupo);
  template.getRange("G11").setValue(cor);

  if (!clinicaRestricao) {
    template.getRange("N4").setValue("NÃO");
  } else {
    template.getRange("N4").setValue(clinicaRestricao);
    template.getRange("N5").setValue(row[26]); 
  }

  if (cursoEstagio) {
    template.getRange("K36").setValue(`*Curso/Estágio: ${cursoEstagio}. Periodo: ${row[32] || ''}`);
  } else {
    template.getRange("K36").setValue("");
  }

  const fullRank = `${posto} ${quadro} ${especialidade}`.toUpperCase();
  const sgpoKeywords = ["BCT", "BCO", "CTA", "PTA", "OEA", "ATCO"];
  const isSgpo = sgpoKeywords.some(key => fullRank.includes(key));
  
  if (om === "4 BAVEX") {
    template.getRange("N1").setValue("4 BAVEX");
  } else if (isSgpo) {
    template.getRange("N1").setValue("SGPO");
  } else {
    template.getRange("N1").setValue("");
  }

  if (archiveCache[cpf]) {
    template.getRange("S4").setValue(archiveCache[cpf]);
  }

  template.getRange("A1:T64").copyTo(target.getRange(startRow, 1));
}

function loadArchiveCache() {
  const cache = {};
  try {
    const ssArchive = SpreadsheetApp.openByUrl(CONFIG.URL_BANCO_ARQUIVOS);
    const sheet = ssArchive.getSheets()[0]; 
    const lastRow = sheet.getLastRow();
    if (lastRow < 4) return cache;
    const data = sheet.getRange(4, 5, lastRow - 3, 2).getValues(); 
    
    data.forEach(row => {
      const rawCpf = String(row[0]).replace(/\D/g, '');
      if (rawCpf) {
        const cpf = rawCpf.padStart(11, '0');
        const archiveNum = row[1];
        if (archiveNum) cache[cpf] = archiveNum;
      }
    });
  } catch (e) {
    Logger.log("Error loading archive cache: " + e.message);
  }
  return cache;
}

function calculateTimeDiff(start, end) {
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
  
  return res.length > 0 ? res.join(", ") : "0 dias";
}

function formatValue(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return val ? String(val) : "";
}
