
import { InspectionData, ProcessingResult, FetchResult } from '../types';

const run = (window as any).google?.script?.run;

export const gasService = {
  fetchInspections: async (): Promise<FetchResult> => {
    if (!run) {
      throw new Error("Ambiente Google Apps Script não detectado. A execução local não é suportada para APIs externas devido a CORS.");
    }

    return new Promise((resolve, reject) => {
      run.withSuccessHandler((data: FetchResult) => resolve(data))
         .withFailureHandler((err: any) => reject(err))
         .getPendingInspections();
    });
  },

  processFichas: async (indices: number[]): Promise<ProcessingResult> => {
    if (!run) {
      throw new Error("Ambiente Google Apps Script não detectado.");
    }

    return new Promise((resolve, reject) => {
      run.withSuccessHandler((res: ProcessingResult) => resolve(res))
         .withFailureHandler((err: any) => reject(err))
         .generateFichasAction(indices);
    });
  }
};
