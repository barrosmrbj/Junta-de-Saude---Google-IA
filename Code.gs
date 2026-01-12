/**
 * GERADOR DE FICHAS DE INSPEÇÃO DE SAÚDE - V6.5
 */

const CONFIG = {
  SS: SpreadsheetApp.getActiveSpreadsheet(),
  ID_BANCO_DADOS: "1CWXzs_J1tTITIZ52_0t02tUb8tBewKSBNWNyaHb6z8M", 
  ID_BANCO_ARQUIVO: "1s41r4hqE0qUgs7i49klZEMjtQXrXDko1gJpi2Hq9ZtE"
};

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Gerenciador FIS - Realtime')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Converte data ISO (yyyy-mm-dd) para formato string PT-BR (dd/mm/yyyy)
 */
function isoToBrDate(isoDate) {
  if (!isoDate) return "";
  const parts = isoDate.split("-");
  if (parts.length !== 3) return isoDate;
  return `${parts[2]}/${parts[1]}/${parts[0]}`;
}

/**
 * Busca dados da aba FICHAS e formata para o frontend
 */
function fetchData(filterDateIso) {
  try {
    const filterDate = isoToBrDate(filterDateIso); 
    const ssBanco = SpreadsheetApp.openById(CONFIG.ID_BANCO_DADOS);
    
    // ABA CORRIGIDA DE 'DADOS' PARA 'FICHAS'
    const sheetDados = ssBanco.getSheetByName("FICHAS") || ssBanco.getSheets()[0];
    
    if (!sheetDados) return { success: false, error: "Aba 'FICHAS' não encontrada no banco." };

    const rawDados = sheetDados.getDataRange().getValues();
    const inspections = [];
    const stats = { totalFichas: 0, impressas: 0, uniquePessoas: 0, homens: 0, mulheres: 0 };
    const uniqueMap = new Map();

    /**
     * Função robusta para extrair apenas a data (DD/MM/YYYY) para comparação
     */
    const formatDataCompare = (val) => {
      if (val instanceof Date) return Utilities.formatDate(val, "GMT-3", "dd/MM/yyyy");
      if (typeof val === 'string') {
        // Tenta capturar dd/mm/yyyy de uma string que pode conter horas
        const match = val.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
        if (match) {
          const d = match[1].padStart(2, '0');
          const m = match[2].padStart(2, '0');
          let y = match[3];
          if (y.length === 2) y = "20" + y;
          return `${d}/${m}/${y}`;
        }
      }
      return String(val || "").split(" ")[0].trim();
    };

    rawDados.forEach((row, idx) => {
      if (idx === 0) return; // Pular cabeçalho
      
      const dataNaPlanilha = formatDataCompare(row[0]);
      
      // Filtragem por data
      if (filterDate && dataNaPlanilha !== filterDate) return;

      const cpf = String(row[6] || "").replace(/\D/g, '').padStart(11, '0');
      const statusFicha = String(row[39] || "AGUARDANDO").toUpperCase(); 
      const sexo = String(row[7] || "").toUpperCase();

      inspections.push({
        originalIndex: idx + 1,
        dtInsp: dataNaPlanilha,
        nome: String(row[11] || "NOME NÃO INFORMADO"),
        posto: String(row[8] || ""),
        quadro: String(row[9] || ""),
        especialidade: String(row[10] || ""),
        dtNascimento: formatDataCompare(row[4]),
        naturalidade: String(row[5] || ""),
        idade: (row[4] instanceof Date) ? calculateAge(row[4]) : calculateAgeFromStr(row[4]),
        cpf: cpf,
        saram: String(row[12] || ""),
        om: String(row[13] || ""),
        vinculo: String(row[15] || ""),
        grupo: String(row[21] || ""), 
        finalidade: String(row[16] || ""), 
        codFicha: String(row[37] || ""),
        status: statusFicha
      });

      stats.totalFichas++;
      if (statusFicha === "IMPRESSO") stats.impressas++;
      if (!uniqueMap.has(cpf)) {
        uniqueMap.set(cpf, true);
        stats.uniquePessoas++;
        if (sexo.includes("MASC") || sexo === "M") stats.homens++;
        else if (sexo.includes("FEM") || sexo === "F") stats.mulheres++;
      }
    });

    return { 
      success: true, 
      inspections: inspections, 
      stats: stats, 
      serverTime: Utilities.formatDate(new Date(), "GMT-3", "HH:mm:ss") 
    };
  } catch (err) { 
    return { success: false, error: err.toString() }; 
  }
}

function deleteFichaAction(rowIndex) {
  try {
    const ssBanco = SpreadsheetApp.openById(CONFIG.ID_BANCO_DADOS);
    const sheetDados = ssBanco.getSheetByName("FICHAS") || ssBanco.getSheets()[0];
    sheetDados.deleteRow(rowIndex);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function generateFichasAction(selectedIndices) {
  try {
    const ss = CONFIG.SS;
    const sheetTemplate = ss.getSheetByName("TEMPLATE");
    const sheetImpressao = ss.getSheetByName("IMPRESSAO");
    
    if (!sheetTemplate || !sheetImpressao) {
       return { success: false, error: "Abas 'TEMPLATE' ou 'IMPRESSAO' não encontradas no script." };
    }
    
    sheetImpressao.clear();
    
    const ssBanco = SpreadsheetApp.openById(CONFIG.ID_BANCO_DADOS);
    const sheetDados = ssBanco.getSheetByName("FICHAS") || ssBanco.getSheets()[0];
    const rawDados = sheetDados.getDataRange().getValues();

    let archiveCache = {};
    try {
      const ssArquivo = SpreadsheetApp.openById(CONFIG.ID_BANCO_ARQUIVO);
      const dataArq = ssArquivo.getSheets()[0].getDataRange().getValues();
      dataArq.forEach((r, i) => { 
        if(i >= 3) {
          const cpfKey = String(r[4] || "").replace(/\D/g, '').padStart(11, '0');
          if (cpfKey !== "00000000000") archiveCache[cpfKey] = r[5]; 
        }
      });
    } catch(e) {}

    let rowCount = 1;
    selectedIndices.sort((a,b) => a - b).forEach(idx => {
      const row = rawDados[idx - 1];
      if (row) {
        populateTemplateLogic(row, sheetTemplate, sheetImpressao, rowCount, archiveCache);
        sheetDados.getRange(idx, 40).setValue("IMPRESSO");
        rowCount += 64;
      }
    });

    const ssId = ss.getId();
    const sheetId = sheetImpressao.getSheetId();
    const pdfUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=pdf&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&fzr=false&sheetid=${sheetId}&top_margin=0.2&bottom_margin=0.2&left_margin=0.2&right_margin=0.2`;

    return {
      success: true,
      message: `${selectedIndices.length} fichas preparadas com sucesso.`,
      printUrl: pdfUrl
    };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function populateTemplateLogic(row, template, target, startRow, archiveCache) {
  template.getRangeList(["S4", "Q4", "N1", "A15", "N5", "Q11", "A13", "H4", "G15", "F11", "Q9", "A9", "O8", "M15", "D6", "H2", "G11", "L11"]).clearContent();
  const cpf = String(row[6] || "").replace(/\D/g, '').padStart(11, '0');
  
  let dtNasc = row[4];
  if (!(dtNasc instanceof Date) && dtNasc) {
     const p = String(dtNasc).split("/");
     if (p.length === 3) dtNasc = new Date(p[2], p[1]-1, p[0]);
  }
  
  let dtPraca = row[14];
  if (!(dtPraca instanceof Date) && dtPraca) {
     const p = String(dtPraca).split("/");
     if (p.length === 3) dtPraca = new Date(p[2], p[1]-1, p[0]);
  }
  
  const postoCompleto = [row[8], row[9], row[10]].filter(Boolean).join(" ");
  
  if (dtNasc instanceof Date) {
    template.getRange("A11").setValue(calculateAge(dtNasc));
    template.getRange("B11").setValue(dtNasc);
  }
  
  if (dtPraca instanceof Date) {
    const tempoServicoStr = calcServiceFull(dtPraca);
    const dataFormatada = Utilities.formatDate(dtPraca, Session.getScriptTimeZone(), "dd/MM/yyyy");
    template.getRange("A15").setValue(dataFormatada + " - " + tempoServicoStr);
  }
  
  template.getRange("H4").setValue(String(row[1] || "")); 
  template.getRange("G15").setValue(String(row[6] || "")); 
  template.getRange("F11").setValue(String(row[7] || "")); 
  template.getRange("Q9").setValue(postoCompleto);
  template.getRange("A9").setValue(String(row[11] || "")); 
  template.getRange("O8").setValue(String(row[12] || "")); 
  template.getRange("M15").setValue(String(row[13] || "")); 
  template.getRange("D6").setValue("Letra(s) " + (row[16] || "")); 
  template.getRange("H2").setValue(String(row[21] || "")); 
  template.getRange("G11").setValue(String(row[23] || "")); 
  template.getRange("L11").setValue(String(row[5] || "")); 
  
  const isSGPO = ["BCT", "BCO", "CTA", "PTA", "OEA", "ATCO"].some(esp => ` ${postoCompleto} `.includes(` ${esp} `));
  if (row[13] === "4 BAVEX") template.getRange("N1").setValue("4 BAVEX");
  else if (isSGPO) template.getRange("N1").setValue("SGPO");
  
  if (archiveCache[cpf]) template.getRange("S4").setValue(archiveCache[cpf]);
  template.getRange("A1:T63").copyTo(target.getRange(startRow, 1));
}

function calculateAgeFromStr(str) {
  if (!str) return 0;
  const p = String(str).split("/");
  if (p.length !== 3) return 0;
  return calculateAge(new Date(p[2], p[1]-1, p[0]));
}

function calculateAge(dt) {
  if (!(dt instanceof Date)) return 0;
  const today = new Date();
  let age = today.getFullYear() - dt.getFullYear();
  if (today.getMonth() < dt.getMonth() || (today.getMonth() === dt.getMonth() && today.getDate() < dt.getDate())) age--;
  return age;
}

function calcServiceFull(dtPraca) {
  const inicio = new Date(dtPraca);
  const hoje = new Date();
  let y = hoje.getFullYear() - inicio.getFullYear();
  let m = hoje.getMonth() - inicio.getMonth();
  let d = hoje.getDate() - inicio.getDate();
  if (d < 0) { m--; d += new Date(hoje.getFullYear(), hoje.getMonth(), 0).getDate(); }
  if (m < 0) { y--; m += 12; }
  return `${y} anos ${m} meses e ${d} dias`;
}