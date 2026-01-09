
import React, { useState, useEffect, useCallback } from 'react';
import { InspectionData, ProcessingResult } from './types';
import { gasService } from './services/gasService';
import { 
  ClipboardDocumentCheckIcon, 
  ArrowPathIcon, 
  PrinterIcon,
  CheckCircleIcon,
  ExclamationCircleIcon,
  IdentificationIcon,
  UserGroupIcon,
  CalendarDaysIcon
} from '@heroicons/react/24/outline';

const App: React.FC = () => {
  const [inspections, setInspections] = useState<InspectionData[]>([]);
  const [selectedIds, setSelectedIds] = useState<Set<number>>(new Set());
  const [loading, setLoading] = useState<boolean>(true);
  const [processing, setProcessing] = useState<boolean>(false);
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [status, setStatus] = useState<{ type: 'success' | 'error' | null, message: string }>({ type: null, message: '' });

  const loadData = useCallback(async () => {
    setLoading(true);
    try {
      const data = await gasService.fetchInspections();
      setInspections(data);
      setStatus({ type: null, message: '' });
    } catch (err) {
      setStatus({ type: 'error', message: 'Erro ao carregar dados do Google Sheets.' });
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
        setStatus({ type: 'success', message: result.message });
        setSelectedIds(new Set());
      } else {
        setStatus({ type: 'error', message: result.message });
      }
    } catch (err) {
      setStatus({ type: 'error', message: 'Erro crítico durante o processamento.' });
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
    <div className="min-h-screen p-4 md:p-8 max-w-[1600px] mx-auto">
      {/* Header */}
      <header className="mb-6 flex flex-col lg:flex-row lg:items-center justify-between gap-4 bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
        <div className="flex items-center gap-4">
          <div className="p-3 bg-blue-600 rounded-xl shadow-lg shadow-blue-200">
            <ClipboardDocumentCheckIcon className="h-8 w-8 text-white" />
          </div>
          <div>
            <h1 className="text-2xl font-bold text-slate-800">
              Gerador de Fichas de Inspeção
            </h1>
            <p className="text-slate-500 text-sm">Controle de inspeções e geração automática para impressão.</p>
          </div>
        </div>
        
        <div className="flex items-center gap-3">
          <div className="relative group">
            <input
              type="text"
              placeholder="Buscar por nome, CPF, RG ou Código..."
              className="w-full md:w-80 pl-10 pr-4 py-2.5 rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all text-sm"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
            <div className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 group-focus-within:text-blue-500 transition-colors">
              <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
              </svg>
            </div>
          </div>

          <button 
            onClick={loadData}
            disabled={loading || processing}
            className="p-2.5 rounded-xl bg-white border border-slate-200 shadow-sm hover:bg-slate-50 transition-colors disabled:opacity-50"
            title="Recarregar dados"
          >
            <ArrowPathIcon className={`h-5 w-5 text-slate-600 ${loading ? 'animate-spin' : ''}`} />
          </button>
          
          <button
            onClick={handleGenerate}
            disabled={selectedIds.size === 0 || processing}
            className="flex items-center gap-2 px-6 py-2.5 bg-blue-600 text-white rounded-xl font-semibold shadow-lg shadow-blue-200 hover:bg-blue-700 active:scale-95 transition-all disabled:opacity-50 disabled:shadow-none disabled:active:scale-100 whitespace-nowrap"
          >
            {processing ? (
              <>
                <div className="h-4 w-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                Processando...
              </>
            ) : (
              <>
                <PrinterIcon className="h-5 w-5" />
                Gerar {selectedIds.size} Ficha(s)
              </>
            )}
          </button>
        </div>
      </header>

      {/* Status Alerts */}
      {status.type && (
        <div className={`mb-6 p-4 rounded-xl flex items-center gap-3 border animate-in fade-in slide-in-from-top-4 duration-300 ${
          status.type === 'success' ? 'bg-emerald-50 border-emerald-100 text-emerald-800' : 'bg-rose-50 border-rose-100 text-rose-800'
        }`}>
          {status.type === 'success' ? (
            <CheckCircleIcon className="h-6 w-6 text-emerald-500" />
          ) : (
            <ExclamationCircleIcon className="h-6 w-6 text-rose-500" />
          )}
          <span className="font-medium">{status.message}</span>
        </div>
      )}

      {/* Table Container */}
      <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse min-w-[1200px]">
            <thead>
              <tr className="bg-slate-50 text-slate-500 text-xs uppercase tracking-wider font-bold">
                <th className="px-4 py-4 w-12 text-center">
                  <input 
                    type="checkbox" 
                    className="h-4 w-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500 cursor-pointer"
                    checked={selectedIds.size === filteredInspections.length && filteredInspections.length > 0}
                    onChange={toggleSelectAll}
                  />
                </th>
                <th className="px-4 py-4">Inspecionado</th>
                <th className="px-4 py-4">Posto / Quadro / Espec.</th>
                <th className="px-4 py-4">Documentos</th>
                <th className="px-4 py-4">OM</th>
                <th className="px-4 py-4">Ficha (Controle)</th>
                <th className="px-4 py-4">Finalidade</th>
                <th className="px-4 py-4">Bio / Vínculo</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {loading ? (
                Array.from({ length: 5 }).map((_, i) => (
                  <tr key={i} className="animate-pulse">
                    <td colSpan={8} className="px-4 py-6 bg-slate-50/20"></td>
                  </tr>
                ))
              ) : filteredInspections.length > 0 ? (
                filteredInspections.map((item) => (
                  <tr 
                    key={item.originalIndex} 
                    className={`hover:bg-slate-50/80 transition-colors cursor-pointer group ${selectedIds.has(item.originalIndex) ? 'bg-blue-50/40' : ''}`}
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
                      <div className="font-bold text-slate-800 uppercase text-sm leading-tight">{item.nome}</div>
                      <div className="flex items-center gap-1.5 mt-0.5">
                        <span className="text-[10px] text-slate-400 font-mono tracking-tighter uppercase">{item.codInsp}</span>
                      </div>
                    </td>
                    <td className="px-4 py-4">
                      <div className="flex flex-col">
                        <span className="inline-flex items-center px-2 py-0.5 rounded-md bg-blue-50 text-blue-700 text-[10px] font-bold uppercase w-fit">
                          {item.posto}
                        </span>
                        <span className="text-xs text-slate-600 font-medium mt-1">
                          {item.quadro} {item.especialidade}
                        </span>
                      </div>
                    </td>
                    <td className="px-4 py-4">
                      <div className="flex flex-col gap-0.5">
                        <div className="flex items-center gap-1 text-xs text-slate-600 font-mono">
                          <IdentificationIcon className="h-3 w-3 text-slate-400" />
                          <span>CPF: {item.cpf}</span>
                        </div>
                        <div className="text-[10px] text-slate-400 ml-4 italic">RG: {item.rg}</div>
                      </div>
                    </td>
                    <td className="px-4 py-4">
                      <div className="text-sm font-semibold text-slate-700">{item.om}</div>
                    </td>
                    <td className="px-4 py-4">
                      <div className="text-xs font-mono font-bold text-blue-600 bg-blue-50/50 px-2 py-1 rounded border border-blue-100/50 w-fit">
                        {item.controle || '---'}
                      </div>
                    </td>
                    <td className="px-4 py-4 max-w-[280px]">
                      <div className="text-xs text-slate-600 line-clamp-2 leading-relaxed" title={item.finalidade}>
                        {item.finalidade || '---'}
                      </div>
                    </td>
                    <td className="px-4 py-4">
                      <div className="flex flex-col gap-1.5">
                        <div className="flex items-center gap-2">
                          <span className="text-[11px] font-bold text-slate-700 whitespace-nowrap">{item.idade} ANOS</span>
                          <span className="text-[10px] text-slate-400 whitespace-nowrap">({item.dtNascimento})</span>
                        </div>
                        <div className="flex items-center gap-3">
                          <div className="flex items-center gap-1 text-[10px] text-slate-500">
                            <UserGroupIcon className="h-3 w-3" />
                            <span className="uppercase truncate max-w-[80px]">{item.vinculo}</span>
                          </div>
                          <div className="flex items-center gap-1 text-[10px] text-slate-500">
                            <div className="h-2 w-2 rounded-full bg-slate-300" />
                            <span className="uppercase truncate max-w-[80px]">{item.grupo}</span>
                          </div>
                        </div>
                      </div>
                    </td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={8} className="px-4 py-16 text-center">
                    <div className="flex flex-col items-center justify-center gap-3 text-slate-400">
                      <ClipboardDocumentCheckIcon className="h-12 w-12 opacity-20" />
                      <p className="text-lg">Nenhuma inspeção encontrada.</p>
                      <p className="text-sm">Tente ajustar seus filtros de busca ou verifique a planilha.</p>
                    </div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      <div className="mt-6 flex flex-col sm:flex-row items-center justify-between gap-4 text-slate-500 text-[11px]">
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-1.5">
            <div className="w-3 h-3 rounded bg-blue-50 border border-blue-100" />
            <span>Selecionado</span>
          </div>
          <div>Total de registros: <span className="font-bold">{filteredInspections.length}</span></div>
        </div>
        <div className="text-center sm:text-right">
          <p>DICA: No Sheets, configure impressão A4, Margens Estreitas, "Ajustar à Página".</p>
          <p className="mt-0.5 opacity-60">Aba IMPRESSÃO será limpa a cada novo processamento.</p>
        </div>
      </div>
    </div>
  );
};

export default App;
