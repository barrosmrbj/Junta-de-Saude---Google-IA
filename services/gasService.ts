
import { InspectionData, ProcessingResult } from '../types';

/**
 * Mocking the google.script.run for local development.
 * In production, this will be replaced by the actual GAS global object.
 */
const run = (window as any).google?.script?.run;

export const gasService = {
  fetchInspections: async (): Promise<InspectionData[]> => {
    if (!run) {
      console.warn("Google Apps Script environment not detected. Returning mock data.");
      return [
        {
          id: '1',
          dtInsp: '08/01/2026',
          codInsp: 'INSP-20251229',
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
          // Adding missing properties to fix type error
          controle: 'F-001',
          finalidade: 'INSPEÇÃO DE SAÚDE PERIÓDICA',
          grupo: 'GRUPO 1',
          vinculo: 'ATIVO',
          idade: 41
        }
      ];
    }

    return new Promise((resolve, reject) => {
      run.withSuccessHandler((data: any) => resolve(data))
         .withFailureHandler((err: any) => reject(err))
         .getPendingInspections();
    });
  },

  processFichas: async (indices: number[]): Promise<ProcessingResult> => {
    if (!run) {
      return new Promise((resolve) => {
        setTimeout(() => resolve({ success: true, message: 'Processamento simulado concluído!', count: indices.length }), 2000);
      });
    }

    return new Promise((resolve, reject) => {
      run.withSuccessHandler((res: ProcessingResult) => resolve(res))
         .withFailureHandler((err: any) => reject(err))
         .generateFichasAction(indices);
    });
  }
};
