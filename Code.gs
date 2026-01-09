
/**
 * GERADOR DE FICHAS DE INSPEÇÃO DE SAÚDE
 */

const CONFIG = {
  SS: SpreadsheetApp.getActiveSpreadsheet(),
  API_FICHAS: "https://script.google.com/macros/s/AKfycbKbu1m8zsQfaqUXyufwVqNmxQRw0pHOo6H528muJ3FIm49zonO537amN309LuRhz52Dw/exec",
  API_INSPECIONADOS: "https://script.google.com/macros/s/AKfycbzydoxvSMbWIJvCktWSdzEhP6g5dC_Wh7e5PqID-D1qJCMNVe8cnGCVrnZwWdtJH20/exec"
};

/**
 * Função Proxy para evitar erro de CORS no navegador.
 */
function proxyFetch(url) {
  try {
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      followRedirects: true,
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    const content = response.getContentText();

    if (responseCode !== 200) {
      throw new Error(`Servidor remoto retornou erro ${responseCode}: ${content.substring(0, 100)}...`);
    }
    
    // Tenta retornar o objeto parseado; se falhar, retorna como texto bruto
    try {
      return JSON.parse(content);
    } catch (e) {
      return content; 
    }
  } catch (err) {
    throw new Error("Falha ao buscar dados externos: " + err.message);
  }
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index.html')
    .setTitle('Gerador de Fichas de Saúde')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function generateFichasAction(selectedIndices) {
  const ss = CONFIG.SS;
  const sheetTemplate = ss.getSheetByName("TEMPLATE");
  const sheetImpressao = ss.getSheetByName("IMPRESSAO");
  
  if (!sheetTemplate || !sheetImpressao) {
    return { success: false, message: "Abas TEMPLATE ou IMPRESSAO não encontradas." };
  }

  try {
    sheetImpressao.clear();
    let currentRow = 1;

    // Busca dados atualizados via proxy
    const fichasApi = proxyFetch(CONFIG.API_FICHAS);
    const fichas = Array.isArray(fichasApi) ? fichasApi : (fichasApi?.data || []);
    
    const inspecionadosApi = proxyFetch(CONFIG.API_INSPECIONADOS);
    const inspecionados = Array.isArray(inspecionadosApi) ? inspecionadosApi : (inspecionadosApi?.data || []);

    const archiveCache = {};
    inspecionados.forEach(row => {
      if (!row || row.length < 6) return;
      const cpf = String(row[4] || "").replace(/\D/g, '').padStart(11, '0');
      const arquivo = row[5];
      if (cpf && cpf !== '00000000000' && arquivo) archiveCache[cpf] = arquivo;
    });

    selectedIndices.sort((a, b) => a - b).forEach(idx => {
      // O idx é a linha na planilha, os dados do array começam em 0 e podem ter cabeçalho
      // Se a API retornar com cabeçalho, idx-1 é o registro correto
      const rowData = fichas[idx - 1]; 
      if (rowData) {
        populateTemplate(rowData, sheetTemplate, sheetImpressao, currentRow, archiveCache);
        currentRow += 64;
      }
    });

    const printUrl = `${ss.getUrl()}#gid=${sheetImpressao.getSheetId()}`;

    return {
      success: true,
      message: `${selectedIndices.length} ficha(s) gerada(s) com sucesso!`,
      printUrl: printUrl
    };
  } catch (e) {
    return { success: false, message: "Erro na geração das fichas: " + e.message };
  }
}

function populateTemplate(row, template, target, startRow, archiveCache) {
  // Limpeza das células dinâmicas no Template
  template.getRange("S4").clearContent(); 
  template.getRange("Q4").clearContent(); 
  template.getRange("N1").clearContent();
  template.getRange("A15").clearContent();
  template.getRange("N4").setValue("NÃO");
  template.getRange("N5").clearContent();

  const cpf = String(row[6] || "").replace(/\D/g, '').padStart(11, '0');
  const dtNasc = row[4] ? new Date(row[4]) : null;
  const dtPraca = row[14] ? new Date(row[14]) : null;
  const today = new Date();

  // Idade
  if (dtNasc && !isNaN(dtNasc.getTime())) {
    let age = today.getFullYear() - dtNasc.getFullYear();
    if (today.getMonth() < dtNasc.getMonth() || (today.getMonth() == dtNasc.getMonth() && today.getDate() < dtNasc.getDate())) age--;
    template.getRange("A11").setValue(age);
    template.getRange("B11").setValue(dtNasc);
  }

  // Tempo de Serviço
  if (dtPraca && !isNaN(dtPraca.getTime())) {
    template.getRange("A15").setValue(calculateServiceTime(dtPraca, today));
  }

  // Dados Pessoais e Militares
  template.getRange("H4").setValue(row[1]); // COD_INSP
  template.getRange("L11").setValue(row[5]); // NATURALIDADE
  template.getRange("G15").setValue(cpf);
  template.getRange("F11").setValue(row[7]); // SEXO
  template.getRange("Q9").setValue(`${row[8]} ${row[9] || ""} ${row[10] || ""}`.trim()); // POSTO QUADRO ESP
  template.getRange("A9").setValue(row[11]); // NOME
  template.getRange("O8").setValue(row[12]); // SARAM
  template.getRange("M15").setValue(row[13]); // OM
  template.getRange("D6").setValue("Letra(s) " + (row[16] || "")); // FINALIDADE
  template.getRange("Q11").setValue(row[17]); // EMAIL
  template.getRange("A13").setValue(row[18]); // ENDERECO
  template.getRange("P15").setValue(row[19]); // TELEFONE
  template.getRange("H2").setValue(row[20]);  // GRUPO
  template.getRange("G11").setValue(row[22]); // COR

  // Restrição Médica
  if (row[24]) {
    template.getRange("N4").setValue(row[24]);
    template.getRange("N5").setValue(row[25]);
  }

  // Nacionalidade simplificada
  const brStates = ["AM", "AC", "AL", "AP", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RO", "RS", "RR", "SC", "SP", "SE", "TO"];
  const naturalidade = String(row[5]).toUpperCase();
  template.getRange("I11").setValue(brStates.includes(naturalidade) ? "BRASILEIRA" : "ESTRANGEIRA");

  // SGPO / BAVEX
  const specs = ["BCT", "BCO", "CTA", "PTA", "OEA", "ATCO"];
  const cargoInfo = `${row[8]} ${row[9] || ""} ${row[10] || ""}`.toUpperCase();
  if (row[13] === "4 BAVEX") {
    template.getRange("N1").setValue("4 BAVEX");
  } else if (specs.some(s => cargoInfo.includes(s))) {
    template.getRange("N1").setValue("SGPO");
  }

  // Arquivo (Prontuário)
  if (archiveCache[cpf]) {
    template.getRange("S4").setValue(archiveCache[cpf]);
  }

  // Copia Template formatado para aba de Impressão
  template.getRange("A1:T64").copyTo(target.getRange(startRow, 1));
}

function calculateServiceTime(start, end) {
  let y = end.getFullYear() - start.getFullYear();
  let m = end.getMonth() - start.getMonth();
  let d = end.getDate() - start.getDate();
  if (d < 0) {
    m--;
    d += new Date(end.getFullYear(), end.getMonth(), 0).getDate();
  }
  if (m < 0) {
    y--;
    m += 12;
  }
  let res = [];
  if (y > 0) res.push(`${y} ano${y > 1 ? 's' : ''}`);
  if (m > 0) res.push(`${m} mê${m > 1 ? 'ses' : 's'}`);
  if (d > 0) res.push(`${d} dia${d > 1 ? 's' : ''}`);
  return res.join(", ") || "0 dias";
}
