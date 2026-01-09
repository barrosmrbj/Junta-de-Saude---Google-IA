
import { InspectionData, ProcessingResult, FetchResult } from '../types';

const run = (window as any).google?.script?.run;

export const gasService = {
  fetchInspections: async (): Promise<FetchResult> => {
    if (!run) {
      console.warn("Google Apps Script environment not detected. Returning mock data.");
      const mockData: InspectionData[] = [
        {
          id: '1',
          dtInsp: new Date().toLocaleDateString('pt-BR'),
          codInsp: 'INSP-20260108-MOCK',
          rg: '528389',
          nome: 'MARCELINO RODRIGUES BARROS JUNIOR',
          cpf: '752.177.912-68',
          om: 'HAMN',
          posto: '3S',
          quadro: 'QESA',
          especialidade: 'SEF',
          dtNascimento: '05/09/1984',
          dtPraca: '06/03/2003',
          originalIndex: 1,
          controle: 'F-001',
          finalidade: 'G1 - Verificação de capacidade funcional',
          grupo: 'IIB',
          vinculo: 'AERONAUTICA',
          idade: 41,
          sexo: 'MASC'
        }
      ];
      return {
        inspections: mockData,
        stats: {
          totalFichas: 1,
          uniqueInspecionandos: 1,
          homens: 1,
          mulheres: 0
        },
        printUrl: '#'
      };
    }

    return new Promise((resolve, reject) => {
      run.withSuccessHandler((data: FetchResult) => resolve(data))
         .withFailureHandler((err: any) => reject(err))
         .getPendingInspections();
    });
  },

  processFichas: async (indices: number[]): Promise<ProcessingResult> => {
    if (!run) {
      return new Promise((resolve) => {
        setTimeout(() => resolve({ 
          success: true, 
          message: 'Processamento simulado concluído!', 
          count: indices.length,
          printUrl: '#' 
        }), 1000);
      });
    }

    return new Promise((resolve, reject) => {
      run.withSuccessHandler((res: ProcessingResult) => resolve(res))
         .withFailureHandler((err: any) => reject(err))
         .generateFichasAction(indices);
    });
  }
};
