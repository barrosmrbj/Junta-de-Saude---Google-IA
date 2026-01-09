
import React, { useState, useEffect, useCallback } from 'react';
import { InspectionData, DashboardStats, ProcessingResult } from './types';
import { gasService } from './services/gasService';
import { 
  ClipboardDocumentCheckIcon, 
  ArrowPathIcon, 
  PrinterIcon,
  CheckCircleIcon,
  ExclamationCircleIcon,
  IdentificationIcon,
  UsersIcon,
  UserIcon,
  DocumentTextIcon,
  MagnifyingGlassIcon
} from '@heroicons/react/24/outline';

const App: React.FC = () => {
  const [inspections, setInspections] = useState<InspectionData[]>([]);
  const [stats, setStats] = useState<DashboardStats>({ totalFichas: 0, uniqueInspecionandos: 0, homens: 0, mulheres: 0 });
  const [selectedIds, setSelectedIds] = useState<Set<number>>(new Set());
  const [loading, setLoading] = useState<boolean>(true);
  const [processing, setProcessing] = useState<boolean>(false);
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [status, setStatus] = useState<{ type: 'success' | 'error' | null, message: string, printUrl?: string }>({ type: null, message: '' });

  const loadData = useCallback(async () => {
    setLoading(true);
    setStatus({ type: null, message: '' });
    try {
      const result = await gasService.fetchInspections();
      setInspections(result.inspections);
      setStats(result.stats);
      if (result.inspections.length === 0) {
        setStatus({ type: null, message: 'Nenhuma inspeção encontrada para a data de hoje.' });
      }
    } catch (err: any) {
      console.error("Erro detalhado:", err);
      setStatus({ 
        type: 'error', 
        message: `Erro ao carregar dados: ${err.message || err || 'Erro desconhecido'}. Verifique se o Script foi publicado corretamente.` 
      });
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    loadData();
  }, [loadData]);

  const toggleSelect = (originalIndex: number) => {
    const newSelected = new Set(selectedIds);
    if (newSelected.has(originalIndex)) {
      newSelected.delete(originalIndex);
    } else {
      newSelected.add(originalIndex);
    }
    setSelectedIds(newSelected);
  };

  const toggleSelectAll = () => {
    if (selectedIds.size === filteredInspections.length && filteredInspections.length > 0) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(filteredInspections.map(i => i.originalIndex)));
    }
  };

  const handleGenerate = async () => {
    if (selectedIds.size === 0) return;
    
    setProcessing(true);
    setStatus({ type: null, message: '' });
    
    try {
      const indices: number[] = Array.from(selectedIds);
      const result = await gasService.processFichas(indices);
      if (result.success) {
        setStatus({ 
          type: 'success', 
          message: `${result.message}`,
          printUrl: result.printUrl 
        });
        setSelectedIds(new Set());
      } else {
        setStatus({ type: 'error', message: result.message });
      }
    } catch (err: any) {
      setStatus({ type: 'error', message: `Erro no processamento: ${err.message || 'Falha na comunicação com o servidor.'}` });
    } finally {
      setProcessing(false);
    }
  };

  const filteredInspections = inspections.filter(item => 
    item.nome.toLowerCase().includes(searchTerm.toLowerCase()) ||
    item.cpf.includes(searchTerm) ||
    item.rg.includes(searchTerm) ||
    (item.codInsp && item.codInsp.toLowerCase().includes(searchTerm.toLowerCase()))
  );

  return (
    <div className="min-h-screen p-4 md:p-8 max-w-[1600px] mx-auto space-y-6">
      <header className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 flex flex-col lg:flex-row lg:items-center justify-between gap-6">
        <div className="flex items-center gap-4">
          <div className="p-3 bg-blue-600 rounded-xl shadow-lg shadow-blue-200">
            <ClipboardDocumentCheckIcon className="h-8 w-8 text-white" />
          </div>
          <div>
            <h1 className="text-2xl font-bold text-slate-800">Inspeções de Hoje</h1>
            <p className="text-slate-500 text-sm">Controle de Fichas de Inspeção de Saúde (FIS)</p>
          </div>
        </div>

        <div className="grid grid-cols-2 md:grid-cols-4 gap-4 flex-1 max-w-2xl">
          <StatCard icon={<DocumentTextIcon className="h-5 w-5"/>} label="Fichas" value={stats.totalFichas} color="text-blue-600" bg="bg-blue-50" />
          <StatCard icon={<UsersIcon className="h-5 w-5"/>} label="Pessoas" value={stats.uniqueInspecionandos} color="text-indigo-600" bg="bg-indigo-50" />
          <StatCard icon={<UserIcon className="h-5 w-5"/>} label="Homens" value={stats.homens} color="text-cyan-600" bg="bg-cyan-50" />
          <StatCard icon={<UserIcon className="h-5 w-5"/>} label="Mulheres" value={stats.mulheres} color="text-pink-600" bg="bg-pink-50" />
        </div>
        
        <div className="flex items-center gap-3">
          <button 
            onClick={loadData}
            disabled={loading || processing}
            className="p-3 rounded-xl bg-white border border-slate-200 shadow-sm hover:bg-slate-50 transition-colors disabled:opacity-50"
            title="Recarregar APIs"
          >
            <ArrowPathIcon className={`h-5 w-5 text-slate-600 ${loading ? 'animate-spin' : ''}`} />
          </button>
          
          <button
            onClick={handleGenerate}
            disabled={selectedIds.size === 0 || processing}
            className="flex items-center gap-2 px-6 py-3 bg-blue-600 text-white rounded-xl font-semibold shadow-lg shadow-blue-200 hover:bg-blue-700 active:scale-95 transition-all disabled:opacity-50 whitespace-nowrap"
          >
            {processing ? (
              <div className="h-5 w-5 border-2 border-white/30 border-t-white rounded-full animate-spin" />
            ) : (
              <PrinterIcon className="h-5 w-5" />
            )}
            Gerar {selectedIds.size} Selecionada(s)
          </button>
        </div>
      </header>

      {status.type && (
        <div className={`p-4 rounded-xl flex flex-col md:flex-row items-center justify-between gap-4 border animate-in fade-in slide-in-from-top-4 ${
          status.type === 'success' ? 'bg-emerald-50 border-emerald-100 text-emerald-800' : 'bg-rose-50 border-rose-100 text-rose-800'
        }`}>
          <div className="flex items-center gap-3">
            {status.type === 'success' ? <CheckCircleIcon className="h-6 w-6 text-emerald-500" /> : <ExclamationCircleIcon className="h-6 w-6 text-rose-500" />}
            <span className="font-medium">{status.message}</span>
          </div>
          {status.printUrl && (
            <a 
              href={status.printUrl} 
              target="_blank" 
              rel="noopener noreferrer"
              className="px-4 py-2 bg-emerald-600 text-white rounded-lg text-sm font-bold shadow-sm hover:bg-emerald-700 transition-colors flex items-center gap-2"
            >
              <PrinterIcon className="h-4 w-4" />
              ABRIR ABA IMPRESSÃO
            </a>
          )}
        </div>
      )}

      <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
        <div className="p-4 border-b border-slate-100 bg-slate-50/30 flex justify-between items-center">
          <div className="relative max-w-md w-full">
            <input
              type="text"
              placeholder="Buscar por nome, CPF ou RG..."
              className="w-full pl-10 pr-4 py-2 rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 text-sm"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
            <MagnifyingGlassIcon className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-slate-400" />
          </div>
          <div className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">
            Exibindo registros de {new Date().toLocaleDateString('pt-BR')}
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse min-w-[1200px]">
            <thead>
              <tr className="bg-slate-50 text-slate-500 text-[10px] uppercase tracking-widest font-black">
                <th className="px-4 py-4 w-12 text-center">
                  <input 
                    type="checkbox" 
                    className="h-4 w-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500 cursor-pointer"
                    checked={selectedIds.size === filteredInspections.length && filteredInspections.length > 0}
                    onChange={toggleSelectAll}
                  />
                </th>
                <th className="px-4 py-4">Inspecionado</th>
                <th className="px-4 py-4">Posto / Quadro / Especialidade</th>
                <th className="px-4 py-4">CPF / Sexo</th>
                <th className="px-4 py-4">Prontuário (Arq.)</th>
                <th className="px-4 py-4">Finalidade</th>
                <th className="px-4 py-4">Idade</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {loading ? (
                Array.from({ length: 5 }).map((_, i) => (
                  <tr key={i} className="animate-pulse">
                    <td colSpan={7} className="px-4 py-6 bg-slate-50/10 border-b border-slate-50"></td>
                  </tr>
                ))
              ) : filteredInspections.length > 0 ? (
                filteredInspections.map((item) => (
                  <tr 
                    key={item.originalIndex} 
                    className={`hover:bg-slate-50 transition-colors cursor-pointer ${selectedIds.has(item.originalIndex) ? 'bg-blue-50/40' : ''}`}
                    onClick={() => toggleSelect(item.originalIndex)}
                  >
                    <td className="px-4 py-4 text-center" onClick={(e) => e.stopPropagation()}>
                      <input 
                        type="checkbox" 
                        className="h-4 w-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500 cursor-pointer"
                        checked={selectedIds.has(item.originalIndex)}
                        onChange={() => toggleSelect(item.originalIndex)}
                      />
                    </td>
                    <td className="px-4 py-4">
                      <div className="font-bold text-slate-800 uppercase text-xs leading-tight">{item.nome}</div>
                      <div className="text-[10px] text-slate-400 font-mono mt-0.5">{item.codInsp}</div>
                    </td>
                    <td className="px-4 py-4 text-[10px] font-medium">
                      <div className="text-blue-700 font-black">{item.posto}</div>
                      <div className="text-slate-500 uppercase">{item.quadro} - {item.especialidade}</div>
                    </td>
                    <td className="px-4 py-4">
                      <div className="flex flex-col gap-0.5">
                        <div className="flex items-center gap-1 text-[11px] text-slate-600 font-mono">
                          <IdentificationIcon className="h-3 w-3 text-slate-400" />
                          <span>{item.cpf}</span>
                        </div>
                        <span className="text-[9px] font-bold text-slate-400 ml-4">{item.sexo}</span>
                      </div>
                    </td>
                    <td className="px-4 py-4">
                      {item.controle ? (
                        <span className="text-[10px] font-black text-blue-600 bg-blue-50 px-2 py-1 rounded border border-blue-100">
                          {item.controle}
                        </span>
                      ) : (
                        <span className="text-[9px] text-slate-300 italic">Não vinculado</span>
                      )}
                    </td>
                    <td className="px-4 py-4 max-w-xs">
                      <div className="text-[10px] text-slate-500 line-clamp-2 italic" title={item.finalidade}>
                        {item.finalidade}
                      </div>
                    </td>
                    <td className="px-4 py-4">
                      <div className="flex flex-col">
                        <span className="text-xs font-bold text-slate-700">{item.idade} ANOS</span>
                        <span className="text-[9px] text-slate-400">{item.dtNascimento}</span>
                      </div>
                    </td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={7} className="px-4 py-20 text-center">
                    <DocumentTextIcon className="h-12 w-12 text-slate-200 mx-auto mb-4" />
                    <p className="text-slate-400 font-medium">Nenhum registro encontrado para hoje.</p>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

const StatCard: React.FC<{ icon: React.ReactNode, label: string, value: number, color: string, bg: string }> = ({ icon, label, value, color, bg }) => (
  <div className={`${bg} p-4 rounded-xl border border-white/50 flex items-center gap-3 shadow-sm`}>
    <div className={`${color} opacity-80`}>{icon}</div>
    <div>
      <div className="text-[10px] font-bold uppercase tracking-wider text-slate-500 opacity-60">{label}</div>
      <div className={`text-xl font-black ${color}`}>{value}</div>
    </div>
  </div>
);

export default App;
