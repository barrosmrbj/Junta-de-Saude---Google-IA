
export interface InspectionData {
  id: string;
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
  vinculo: string;
  finalidade: string;
  grupo: string;
  idade: number;
  controle: string;
  sexo: string;
}

export interface DashboardStats {
  totalFichas: number;
  uniqueInspecionandos: number;
  homens: number;
  mulheres: number;
}

export interface FetchResult {
  inspections: InspectionData[];
  stats: DashboardStats;
  printUrl: string;
}

export interface ProcessingResult {
  success: boolean;
  message: string;
  count?: number;
  printUrl?: string;
}
