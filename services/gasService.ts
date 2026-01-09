
import { InspectionData, FetchResult, ProcessingResult } from '../types';

const URL_FICHAS = "https://script.google.com/macros/s/AKfycbzBwBAKyHudfz2go2Z2z6yAFiq1Yrt6GduggBbsoSAPYw5dOu7aeBG69K_RGGXzlvi4Og/exec";
const URL_INSPECIONANDO = "https://script.google.com/macros/s/AKfycbzydoxvSMbWIJvCktWSdzEhP6g5dC_Wh7e5PqID-D1qJCMNVe8cnGCVrnZwWdtJH20/exec";

// Referência global ao Google Apps Script
const run = (window as any).google?.script?.run;

/**
 * Função para buscar dados via Proxy do Apps Script para evitar CORS.
 */
async function fetchDados(url: string): Promise<any> {
  // Se não houver 'run', estamos em ambiente local/preview. Retornamos null para ativar Mocks.
  if (!run) return null;
  
  return new Promise((resolve, reject) => {
    run.withSuccessHandler((response: any) => {
      if (typeof response === 'string') {
        try {
          resolve(JSON.parse(response));
        } catch (e) {
          reject(new Error("Falha ao processar resposta do servidor como JSON."));
        }
      } else {
        resolve(response);
      }
    })
    .withFailureHandler((err: any) => {
      reject(new Error(err.message || "Erro na comunicação com o servidor Apps Script."));
    })
    .proxyFetch(url);
  });
}

const formatBRDate = (date: Date): string => {
  return date.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric' });
};

const calculateAge = (birthday: string | Date): number => {
  const birth = new Date(birthday);
  if (isNaN(birth.getTime())) return 0;
  const today = new Date();
  let age = today.getFullYear() - birth.getFullYear();
  const m = today.getMonth() - birth.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < birth.getDate())) {
    age--;
  }
  return age;
};

/**
 * Gera dados de exemplo para quando o GAS não está disponível (Ambiente Local/Preview)
 */
const getMockData = (): FetchResult => {
  const today = new Date();
  const todayStr = formatBRDate(today);
  
  const mockRows: InspectionData[] = [
    {
      id: "INSP-8A99ED-20251229",
      dtInsp: todayStr,
      codInsp: "INSP-8A99ED-20251229",
      rg: "528389",
      nome: "MARCELINO RODRIGUES BARROS JUNIOR",
      cpf: "752.177.912-68",
      sexo: "MASC",
      posto: "3S",
      quadro: "QESA",
      especialidade: "SEF",
      om: "HAMN",
      dtNascimento: "05/09/1984",
      dtPraca: "06/03/2003",
      originalIndex: 2,
      vinculo: "AERONAUTICA",
      finalidade: "G1 - Verificação de capacidade funcional...",
      grupo: "IIB - DEMAIS",
      idade: calculateAge("1984-09-05"),
      controle: "FICHC59B82-20260108"
    },
    {
      id: "INSP-X921-2026",
      dtInsp: todayStr,
      codInsp: "INSP-X921-2026",
      rg: "123456",
      nome: "ANA MARIA SILVEIRA",
      cpf: "111.222.333-44",
      sexo: "FEM",
      posto: "1T",
      quadro: "QOEA",
      especialidade: "COM",
      om: "VII COMAR",
      dtNascimento: "10/10/1990",
      dtPraca: "01/02/2010",
      originalIndex: 3,
      vinculo: "AERONAUTICA",
      finalidade: "G1 - Manutenção de Aptidão",
      grupo: "IIA - TRIPULANTES",
      idade: calculateAge("1990-10-10"),
      controle: "ARQ-44552"
    }
  ];

  return {
    inspections: mockRows,
    stats: {
      totalFichas: mockRows.length,
      uniqueInspecionandos: 2,
      homens: 1,
      mulheres: 1
    },
    printUrl: "#"
  };
};

export const gasService = {
  fetchInspections: async (): Promise<FetchResult> => {
    // Se não estiver no GAS, retorna Mock Data
    if (!run) {
      console.warn("Aviso: google.script.run não detectado. Usando Mock Mode.");
      return getMockData();
    }

    try {
      const [fichasJson, inspecionandosJson]: [any, any] = await Promise.all([
        fetchDados(URL_FICHAS),
        fetchDados(URL_INSPECIONANDO)
      ]);

      const rows = Array.isArray(fichasJson) ? fichasJson : (fichasJson?.data || []);
      if (!rows || rows.length === 0) {
        return { 
          inspections: [], 
          stats: { totalFichas: 0, uniqueInspecionandos: 0, homens: 0, mulheres: 0 }, 
          printUrl: "" 
        };
      }

      const inspecionadosRows = Array.isArray(inspecionandosJson) ? inspecionandosJson : (inspecionandosJson?.data || []);
      const archiveMap: Record<string, string> = {};
      
      inspecionadosRows.forEach((r: any[]) => {
        if (!r || r.length < 6) return;
        const cpfRaw = String(r[4] || "").replace(/\D/g, '').padStart(11, '0');
        if (cpfRaw && cpfRaw !== '00000000000') archiveMap[cpfRaw] = String(r[5] || "");
      });

      const todayStr = formatBRDate(new Date());
      const uniqueCpfs = new Set<string>();
      let homens = 0;
      let mulheres = 0;

      const startIdx = rows[0]?.[0] === "DT_INSP" ? 1 : 0;
      
      const filteredInspections = rows.slice(startIdx).map((row: any[], index: number) => {
        if (!row || row.length < 2) return null;

        const dtRaw = row[0];
        const dtParsed = new Date(dtRaw);
        const dtFormatada = !isNaN(dtParsed.getTime()) ? formatBRDate(dtParsed) : String(dtRaw);

        if (!dtFormatada.includes(todayStr)) return null;

        const cpf = String(row[6] || "").replace(/\D/g, '').padStart(11, '0');
        const sexo = String(row[7] || "").toUpperCase();
        
        if (cpf && cpf !== '00000000000') uniqueCpfs.add(cpf);
        if (sexo.includes("M") || sexo.includes("MASC")) homens++;
        else if (sexo.includes("F") || sexo.includes("FEM")) mulheres++;

        return {
          id: String(row[1] || `row-${index}`),
          dtInsp: dtFormatada,
          codInsp: row[1],
          rg: row[2],
          dtNascimento: row[4] ? formatBRDate(new Date(row[4])) : "",
          cpf: cpf,
          sexo: sexo,
          nome: row[11],
          om: row[13],
          posto: row[8],
          quadro: row[9],
          especialidade: row[10],
          dtPraca: row[14] ? formatBRDate(new Date(row[14])) : "",
          originalIndex: index + startIdx + 1,
          vinculo: row[15],
          finalidade: row[16],
          grupo: row[20],
          idade: row[4] ? calculateAge(row[4]) : 0,
          controle: archiveMap[cpf] || ""
        } as InspectionData;
      }).filter((item: any) => item !== null);

      return {
        inspections: filteredInspections,
        stats: {
          totalFichas: filteredInspections.length,
          uniqueInspecionandos: uniqueCpfs.size,
          homens,
          mulheres
        },
        printUrl: "" 
      };
    } catch (err: any) {
      console.error("Erro na camada de serviço:", err);
      throw err;
    }
  },

  processFichas: async (indices: number[]): Promise<ProcessingResult> => {
    if (!run) {
      // Simulação de processamento para ambiente de teste
      return new Promise((resolve) => {
        setTimeout(() => {
          resolve({
            success: true,
            message: `(MODO SIMULAÇÃO) ${indices.length} ficha(s) processada(s) com sucesso. No ambiente real, elas seriam enviadas para a aba IMPRESSÃO.`,
            printUrl: "#"
          });
        }, 1500);
      });
    }

    return new Promise((resolve, reject) => {
      run.withSuccessHandler((res: ProcessingResult) => resolve(res))
         .withFailureHandler((err: any) => reject(new Error(err.message || "Erro no processamento das fichas.")))
         .generateFichasAction(indices);
    });
  }
};
