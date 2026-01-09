
export interface InspectionData {
  id: string; // Internal unique ID
  dtInsp: string;
  codInsp: string;
  rg: string;
  nome: string;
  cpf: string;
  om: string;
  posto: string;
  quadro: string;
  especialidade: string;
  dtNascimento: string;
  dtPraca: string;
  originalIndex: number;
  // New requested fields
  controle: string;      // Cod da ficha
  finalidade: string;    // Finalidade da inspeção
  grupo: string;         // Grupo
  vinculo: string;       // Vínculo
  idade: number;         // Idade calculada
}

export interface ProcessingResult {
  success: boolean;
  message: string;
  count?: number;
}
