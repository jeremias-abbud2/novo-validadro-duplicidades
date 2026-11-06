import { useState, useCallback, useRef, DragEvent, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  Upload, 
  AlertCircle, 
  CheckCircle2, 
  Clock, 
  User, 
  Building2, 
  FileText,
  Trash2,
  FileDown,
  Search,
  Filter,
  X,
  BarChart3,
  TrendingUp,
  Calendar,
  AlertTriangle,
  PieChart,
  Files,
  Plus,
  Layers,
  FileSpreadsheet,
  Cpu,
  Database,
  Network,
  Terminal,
  ShieldAlert,
  Wrench
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { cn } from './lib/utils';
import { exportFormattedExcelReport, exportFormattedDashboardReport, exportFormattedITReport, ITDiagnosticsStats } from './lib/excelExport';

// Types for the report data
interface TransactionRow {
  'DATA DA TRAN'?: string | number | Date;
  'TIPO DE TRANSAÇÃO'?: string;
  'BENEFICIÁRIO'?: string;
  'MATRÍCULA'?: string | number;
  'PRESTADOR'?: string;
  'PROCEDIMENTO'?: string;
  'VALOR'?: string | number;
  'QUANT'?: string | number;
  'SITUAÇÃO'?: string;
  'Nº CONSELHO'?: string | number;
  [key: string]: any;
}

interface DuplicateGroup {
  fingerprint: string;
  rows: TransactionRow[];
  params: {
    prestador: string;
    prestador_matricula: string;
    beneficiario: string;
    beneficiario_matricula: string;
    procedimento: string;
    especialidade: string;
    situacao: string;
    tipo: string;
    quantidade: string;
    conselho: string;
    data: string;
    conselhoExecTipo: string;
    conselhoExecNum: string;
    conselhoExecUf: string;
  };
}

interface FilterState {
  prestador: string;
  beneficiario: string;
  matricula: string;
  procedimento: string;
  data: string;
  situacao: string;
  intervalCategory: string;
  conselhoExecutante: string;
}

export default function App() {
  const [data, setData] = useState<TransactionRow[] | null>(null);
  const [duplicateGroups, setDuplicateGroups] = useState<DuplicateGroup[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [transactionCount, setTransactionCount] = useState<number>(0);
  const [filters, setFilters] = useState<FilterState>({
    prestador: '',
    beneficiario: '',
    matricula: '',
    procedimento: '',
    data: '',
    situacao: '',
    intervalCategory: '',
    conselhoExecutante: ''
  });

  const [uniqueValues, setUniqueValues] = useState<{
    prestadores: string[];
    procedimentos: string[];
    situacoes: string[];
    datas: string[];
    conselhoExecutantes: string[];
  }>({
    prestadores: [],
    procedimentos: [],
    situacoes: [],
    datas: [],
    conselhoExecutantes: [],
  });
  
  const [activeTab, setActiveTab] = useState<'table' | 'dashboard' | 'itDiagnostics'>('table');
  const [uploadMode, setUploadMode] = useState<'single' | 'multi'>('single');
  const [batchFiles, setBatchFiles] = useState<File[]>([]);
  const [sourceFiles, setSourceFiles] = useState<string[]>([]);
  const [processingStatus, setProcessingStatus] = useState<string | null>(null);
  const [deduplicationMessage, setDeduplicationMessage] = useState<string | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const multiFileInputRef = useRef<HTMLInputElement>(null);
  const addMoreFileInputRef = useRef<HTMLInputElement>(null);

  function normalizeDate(val: any): { date: string; time: string } {
    if (val === undefined || val === null || val === '') return { date: '', time: '' };
    
    try {
      if (typeof val === 'number') {
        // Caso Excel: pode ser data+hora ou apenas hora (valor < 1)
        const date = XLSX.SSF.parse_date_code(val);
        const d = val >= 1 ? `${date.d.toString().padStart(2, '0')}/${date.m.toString().padStart(2, '0')}/${date.y}` : '';
        const t = `${date.H.toString().padStart(2, '0')}:${date.M.toString().padStart(2, '0')}`;
        return { date: d, time: t };
      }

      if (val instanceof Date) {
        if (isNaN(val.getTime())) return { date: '', time: '' };
        const d = `${val.getDate().toString().padStart(2, '0')}/${(val.getMonth() + 1).toString().padStart(2, '0')}/${val.getFullYear()}`;
        const t = `${val.getHours().toString().padStart(2, '0')}:${val.getMinutes().toString().padStart(2, '0')}`;
        return { date: d, time: t };
      }

      const str = String(val).trim();
      if (!str) return { date: '', time: '' };

      const parts = str.split(/[\sT]/);
      
      // Caso string contenha apenas horário (ex: "12:30" ou "00:00" ou "0:00")
      if (parts.length === 1 && parts[0].includes(':')) {
        const timeSubParts = parts[0].split(':');
        const h = (timeSubParts[0] || '0').padStart(2, '0');
        const m = (timeSubParts[1] || '0').padStart(2, '0');
        return { date: '', time: `${h}:${m}` };
      }

      const datePart = parts[0] || '';
      let timePart = '';
      if (parts[1] && parts[1].includes(':')) {
        const timeSubParts = parts[1].split(':');
        const h = (timeSubParts[0] || '0').padStart(2, '0');
        const m = (timeSubParts[1] || '0').padStart(2, '0');
        timePart = `${h}:${m}`;
      } else if (parts[1]) {
        timePart = parts[1].substring(0, 5);
      }
      
      return {
        date: datePart,
        time: timePart
      };
    } catch (e) {
      const s = String(val);
      return s.includes(':') ? { date: '', time: s.substring(0, 5) } : { date: s, time: '' };
    }
  };

  function getRowTime(row: any): string {
    const dateVal = getColValue(row, 'DATA');
    const hourVal = getColValue(row, 'HORA');
    
    // 1. Tenta extrair da coluna HORA se existir e não for vazia
    if (hourVal !== undefined && hourVal !== null && String(hourVal).trim() !== '') {
      const hNorm = normalizeDate(hourVal);
      if (hNorm.time) {
        return hNorm.time;
      }
      if (typeof hourVal === 'string' && hourVal.includes(':')) {
        const [h, m] = hourVal.trim().split(':');
        return `${(h || '0').padStart(2, '0')}:${(m || '0').padStart(2, '0')}`;
      }
      if (typeof hourVal === 'number' || (!isNaN(Number(hourVal)) && String(hourVal).trim() !== '')) {
        const num = Number(hourVal);
        // Fração decimal do Excel para horários (ex: 0.5 = 12:00, 0 = 00:00)
        if (num >= 0 && num < 1) {
          const totalMin = Math.round(num * 24 * 60);
          const h = Math.floor(totalMin / 60) % 24;
          const m = totalMin % 60;
          return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}`;
        }
        // Horário militar inteiro (ex: 0 -> 00:00, 15 -> 00:15, 130 -> 01:30, 1430 -> 14:30)
        if (num >= 0 && num <= 2400) {
          const s = String(Math.floor(num)).padStart(4, '0');
          const h = s.substring(0, 2);
          const m = s.substring(2, 4);
          if (Number(h) < 24 && Number(m) < 60) {
            return `${h}:${m}`;
          }
        }
      }
    }

    // 2. Tenta extrair horário da coluna DATA se presente
    const dNorm = normalizeDate(dateVal);
    if (dNorm.time) {
      return dNorm.time;
    }

    return '';
  }

  function getGroupTimeInterval(rows: TransactionRow[]) {
    const times: number[] = [];
    rows.forEach(row => {
      const time = getRowTime(row);
      // Permite horários em 00:00 e arredores (ex: meia-noite em internações)
      if (time && time !== '--:--' && time.includes(':')) {
        const [hh, mm] = time.split(':').map(Number);
        if (!isNaN(hh) && !isNaN(mm) && hh >= 0 && hh < 24 && mm >= 0 && mm < 60) {
          times.push(hh * 60 + mm);
        }
      }
    });

    if (times.length < 2) return { minDiff: null, label: '', level: 'normal' };

    // Calcula a menor diferença circular entre qualquer par de horários (relojoeiro de 24h = 1440 min)
    let minDiff = Infinity;
    for (let i = 0; i < times.length; i++) {
      for (let j = i + 1; j < times.length; j++) {
        const rawDiff = Math.abs(times[i] - times[j]);
        const circularDiff = Math.min(rawDiff, 1440 - rawDiff);
        if (circularDiff < minDiff) {
          minDiff = circularDiff;
        }
      }
    }

    if (minDiff === 0) {
      return { minDiff, label: `Mesmo Horário (Muito Suspeito)`, level: 'very_high' };
    } else if (minDiff <= 30) {
      return { minDiff, label: `≤ 30m (Bem Provável)`, level: 'high' };
    } else if (minDiff <= 60) {
      return { minDiff, label: `30m - 1h (Pouco Menos)`, level: 'medium' };
    } else {
      return { minDiff, label: `> 1h`, level: 'low' };
    }
  };

  function getColValue(row: any, concept: string): any {
    const keys = Object.keys(row);
    
    // Normalização agressiva para lidar com diferentes encodings e caracteres especiais
    const normalize = (s: string) => {
      if (!s) return '';
      return String(s)
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '') // Remove acentos
        .toUpperCase()
        .replace(/\s+/g, '') // Remove espaços
        .replace(/[^A-Z0-9]/g, ''); // Remove tudo que não é alfanumérico
    };

    const aliases: Record<string, string[]> = {
      'DATA': ['DATA_TRAN', 'DATA DA TRAN', 'DT_TRAN', 'DATA', 'DIA', 'DATA_TRANSICAO', 'DATA_ATENDIMENTO', 'DATA DA TRANSACAO'],
      'HORA': ['HORARIO_', 'HORA', 'HORARIO', 'TIME', 'HORA_TRAN', 'HORA_ATENDIMENTO'],
      'BENE_NOME': ['NOME_BENEFICIARIO', 'BENEFICIÁRIO', 'NOME_BE', 'NM_BENEFICIARIO', 'PACIENTE', 'NOME', 'BENEFICIARIO', 'NOME DO BENEFICIARIO'],
      'BENE_MATRICULA': ['MATRICULA_BENEFICIARIO', 'MATRICULA_BENE', 'MATRÍCULA', 'MATRICUL', 'MATRICULA', 'COD_BENE', 'CARTEIRA', 'MATRICULA_BENEFICIARIO_DV', 'MATRICULA_PACIENTE', 'MATRICULA DO BENEFICIARIO'],
      'PREST_NOME': ['NOME_PRESTADOR', 'NM_PRESTADOR', 'PRESTADOR', 'NOME_PR', 'UNIDADE', 'NOME_PRESTADOR_EXECUTANTE', 'NOME DO PRESTADOR'],
      'PREST_MATRICULA': ['MATRICULA_PRESTADOR', 'COD_PRESTADOR', 'CNPJ_PRESTADOR', 'MATRICULA DO PRESTADOR'],
      'PROCEDIMENTO': ['PROCEDIMI', 'PROCEDIMENTO', 'COD_PROC', 'PROCED', 'CODIGO', 'PROCE', 'ITEM', 'PROC'],
      'QUANTIDADE': ['QUANTIDADE', 'QUANT', 'QTD', 'QUANT_SOLIC', 'QTDE'],
      'SITUACAO': ['SITUAÇÃO', 'SITUACAO', 'STATUS', 'RESULTADO', 'SIT'],
      'CONSELHO': ['Nº CONSELHO', 'NUMERO_CONSELHO', 'CONSELHO', 'CRM', 'NUMERO_', 'REGISTRO_CONSELHO', 'CONSELHO PROFISSIONAL'],
      'TIPO': ['TIPO_TRA', 'TIPO DE TRANSAÇÃO', 'TIPO', 'TIPO_TRANSACAO', 'TIPO DE TRANSACAO', 'TIPO DE SOLICITACAO', 'TIPO_SOLICITACAO'],
      'ESPECIALIDADE': ['ESPECIALIDADE', 'ESPEC', 'AREA_MEDICA', 'NOME_ESPECIALIDADE', 'ESPECIALIDADE_PRESTADOR'],
      'TIPO_CONSELHO_EXECUTANTE': ['TIPO_CONSELHO_EXECUTANTE', 'TIPO_CONSELHO_EXEC', 'CONSELHO_EXECUTANTE_TIPO', 'TIPO_CONSELHO', 'CONSELHO_TIPO'],
      'NUMERO_CONSELHO_EXECUTANTE': ['NUMERO_CONSELHO_EXECUTANTE', 'NUM_CONSELHO_EXECUTANTE', 'NUMERO_CONSELHO_EXEC', 'CRM_EXECUTANTE', 'NUM_CONSELHO_EXEC'],
      'UF_CONSELHO_EXECUTANTE': ['UF_CONSELHO_EXECUTANTE', 'UF_CONSELHO_EXEC', 'CONSELHO_EXECUTANTE_UF', 'UF_CRM_EXECUTANTE'],
      'REMESSA': [
        'REMESSA', 'NUMERO_REMESSA', 'NUMERO DA REMESSA', 'NUMERO DE REMESSA',
        'Nº REMESSA', 'NO_REMESSA', 'NR_REMESSA', 'NR REMESSA', 'N_REMESSA',
        'NUM_REMESSA', 'REMESSA_NUMERO', 'REMESSA_NUM', 'NUM_REM', 'COD_REMESSA',
        'CODIGO_REMESSA', 'LOTE_REMESSA', 'NUMERO_LOTE', 'NUM_LOTE', 'NR_LOTE'
      ],
      'PROTOCOLO': [
        'PROTOCOLO', 'NUMERO_PROTOCOLO', 'NUMERO DO PROTOCOLO', 'NUMERO DE PROTOCOLO',
        'Nº PROTOCOLO', 'NO_PROTOCOLO', 'NR_PROTOCOLO', 'NR PROTOCOLO', 'N_PROTOCOLO',
        'NUM_PROTOCOLO', 'PROTOCOLO_NUMERO', 'PROTOCOLO_NUM', 'NUM_PROT', 'COD_PROTOCOLO',
        'CODIGO_PROTOCOLO', 'PROTOCOLO_TRANSACAO', 'PROTOCOLO_ATENDIMENTO', 'NUM_PROTOCOLO_TRANSACAO'
      ],
      'ID_TRANSACAO': [
        'TRANSACAO', 'NUM_TRANSACAO', 'NUMERO_TRANSACAO', 'NR_TRANSACAO', 'NO_TRANSACAO',
        'COD_TRANSACAO', 'CODIGO_TRANSACAO', 'ID_TRANSACAO', 'GUIA', 'NUMERO_GUIA',
        'NUM_GUIA', 'NR_GUIA', 'SENHA', 'NUMERO_AUTORIZACAO', 'NUM_AUTORIZACAO',
        'AUTORIZACAO', 'ATENDIMENTO', 'NUM_ATENDIMENTO', 'NUMERO_ATENDIMENTO'
      ]
    };

    const targets = aliases[concept] || [concept];
    const normalizedTargets = targets.map(normalize);

    // 1. Prioridade Máxima: Match Exato (sem espaços e case-insensitive)
    for (const target of targets) {
      const found = keys.find(k => k.trim().toUpperCase() === target.toUpperCase());
      if (found) return row[found];
    }

    // 2. Prioridade Média: Match Normalizado
    const candidates = keys.map(k => {
      const normalizedK = normalize(k);
      let bestScore = 0;
      
      normalizedTargets.forEach(t => {
        if (normalizedK === t) bestScore = 100;
        else if (normalizedK.startsWith(t)) bestScore = Math.max(bestScore, 80);
        else if (normalizedK.includes(t)) bestScore = Math.max(bestScore, 50);
      });

      // Penalização para evitar confusão entre Beneficiário e Prestador
      const upperK = k.toUpperCase();
      if ((concept === 'BENE_NOME' || concept === 'BENE_MATRICULA') && upperK.includes('PREST')) bestScore -= 60;
      if ((concept === 'PREST_NOME' || concept === 'PREST_MATRICULA') && upperK.includes('BENE')) bestScore -= 60;

      return { key: k, score: bestScore };
    }).filter(c => c.score > 40);

    candidates.sort((a, b) => b.score - a.score);
    
    return candidates.length > 0 ? row[candidates[0].key] : undefined;
  };

  const normalizeText = (v: any) => String(v ?? '').trim().toLowerCase();
  const normalizeRaw = (v: any) => String(v ?? '').trim();
  const normalizeNum = (v: any) => {
    if (v === undefined || v === null || v === '') return '';
    const n = parseFloat(String(v));
    return isNaN(n) ? normalizeText(v) : String(n);
  };

  const robustNormalize = (s: any) => {
    if (s === undefined || s === null || s === '') return '';
    return String(s)
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '') // Remove acentos
      .toUpperCase()
      .replace(/\s+/g, '') // Remove espaços
      .replace(/[^A-Z0-9]/g, ''); // Remove tudo que não é alfanumérico
  };

  const hasRemessaValue = (row: TransactionRow): boolean => {
    const remVal = getColValue(row, 'REMESSA');
    if (remVal === undefined || remVal === null) return false;
    const str = String(remVal).trim();
    return str !== '' && str !== '0' && str !== '-' && str !== '--' && str.toLowerCase() !== 'null' && str.toLowerCase() !== 'undefined';
  };

  const isSameAttendance = (rowA: TransactionRow, rowB: TransactionRow): boolean => {
    // 1. Se ambos possuírem Protocolo / Guia / ID de Transação, compara diretamente
    const protA = robustNormalize(getColValue(rowA, 'PROTOCOLO') || getColValue(rowA, 'ID_TRANSACAO'));
    const protB = robustNormalize(getColValue(rowB, 'PROTOCOLO') || getColValue(rowB, 'ID_TRANSACAO'));
    
    const bMatA = robustNormalize(getColValue(rowA, 'BENE_MATRICULA'));
    const bMatB = robustNormalize(getColValue(rowB, 'BENE_MATRICULA'));
    const bNomeA = robustNormalize(getColValue(rowA, 'BENE_NOME'));
    const bNomeB = robustNormalize(getColValue(rowB, 'BENE_NOME'));

    // Beneficiário deve coincidir (por matrícula ou por nome)
    const matchBene = (bMatA && bMatB && bMatA === bMatB) || 
                      (bNomeA && bNomeB && bNomeA === bNomeB) || 
                      (!bMatA && !bMatB && bNomeA === bNomeB) || 
                      (!bNomeA && !bNomeB && bMatA === bMatB);
    if (!matchBene) return false;

    const procA = robustNormalize(getColValue(rowA, 'PROCEDIMENTO'));
    const procB = robustNormalize(getColValue(rowB, 'PROCEDIMENTO'));
    if (procA !== procB) return false;

    const dateA = normalizeDate(getColValue(rowA, 'DATA')).date;
    const dateB = normalizeDate(getColValue(rowB, 'DATA')).date;
    if (dateA && dateB && dateA !== dateB) return false;

    // Se ambos tiverem número de protocolo/guia preenchido, é o critério decisivo
    if (protA && protB) {
      return protA === protB;
    }

    // 2. Prestador (Matrícula ou Nome)
    const pMatA = robustNormalize(getColValue(rowA, 'PREST_MATRICULA'));
    const pMatB = robustNormalize(getColValue(rowB, 'PREST_MATRICULA'));
    const pNomeA = robustNormalize(getColValue(rowA, 'PREST_NOME'));
    const pNomeB = robustNormalize(getColValue(rowB, 'PREST_NOME'));
    if (pMatA && pMatB && pMatA !== pMatB) return false;
    if (pNomeA && pNomeB && pNomeA !== pNomeB && !pMatA && !pMatB) return false;

    // 3. Quantidade
    const qtdA = normalizeNum(getColValue(rowA, 'QUANTIDADE'));
    const qtdB = normalizeNum(getColValue(rowB, 'QUANTIDADE'));
    if (qtdA && qtdB && qtdA !== qtdB) return false;

    // 4. CRM / Conselho Executante (se presente em ambos)
    const crmA = robustNormalize(getColValue(rowA, 'NUMERO_CONSELHO_EXECUTANTE') || getColValue(rowA, 'CONSELHO'));
    const crmB = robustNormalize(getColValue(rowB, 'NUMERO_CONSELHO_EXECUTANTE') || getColValue(rowB, 'CONSELHO'));
    if (crmA && crmB && crmA !== crmB) return false;

    // 5. Horário
    const timeA = getRowTime(rowA);
    const timeB = getRowTime(rowB);
    const hasTimeA = timeA && timeA !== '--:--' && timeA.includes(':');
    const hasTimeB = timeB && timeB !== '--:--' && timeB.includes(':');

    // Se ambos os relatórios têm horário informado, os horários DEVEM coincidir para ser o mesmo atendimento
    if (hasTimeA && hasTimeB) {
      return timeA === timeB;
    }

    return true;
  };

  const consolidateMultipleFileRows = (rows: TransactionRow[]): { consolidatedRows: TransactionRow[]; deduplicatedCount: number } => {
    // 1. Identifica quais linhas possuem valor de remessa válido
    const rowsWithRemessa: TransactionRow[] = [];
    const rowsWithoutRemessa: TransactionRow[] = [];

    rows.forEach(r => {
      if (hasRemessaValue(r)) {
        rowsWithRemessa.push(r);
      } else {
        rowsWithoutRemessa.push(r);
      }
    });

    // Se não há linhas com remessa ou não há linhas sem remessa, mantém todas
    if (rowsWithRemessa.length === 0 || rowsWithoutRemessa.length === 0) {
      return { consolidatedRows: rows, deduplicatedCount: 0 };
    }

    // 2. Indexa as linhas com remessa em baldes (buckets) por Beneficiário + Procedimento + Data
    const remessaBuckets = new Map<string, { row: TransactionRow; matched: boolean }[]>();

    rowsWithRemessa.forEach(row => {
      const bMat = robustNormalize(getColValue(row, 'BENE_MATRICULA'));
      const bNome = robustNormalize(getColValue(row, 'BENE_NOME'));
      const proc = robustNormalize(getColValue(row, 'PROCEDIMENTO'));
      const { date } = normalizeDate(getColValue(row, 'DATA'));
      
      const key = `${bMat || bNome}|${proc}|${date}`;
      if (!remessaBuckets.has(key)) {
        remessaBuckets.set(key, []);
      }
      remessaBuckets.get(key)!.push({ row, matched: false });
    });

    // 3. Avalia as linhas sem remessa: se coincidirem com um atendimento que já possui remessa,
    // descarta a versão sem remessa ("apenas traga o resultado com remessa")
    const keptRowsWithoutRemessa: TransactionRow[] = [];
    let deduplicatedCount = 0;

    rowsWithoutRemessa.forEach(rowNoRem => {
      const bMat = robustNormalize(getColValue(rowNoRem, 'BENE_MATRICULA'));
      const bNome = robustNormalize(getColValue(rowNoRem, 'BENE_NOME'));
      const proc = robustNormalize(getColValue(rowNoRem, 'PROCEDIMENTO'));
      const { date } = normalizeDate(getColValue(rowNoRem, 'DATA'));
      const key = `${bMat || bNome}|${proc}|${date}`;

      const candidates = remessaBuckets.get(key);
      let matched = false;

      if (candidates && candidates.length > 0) {
        for (const cand of candidates) {
          if (!cand.matched && isSameAttendance(rowNoRem, cand.row)) {
            matched = true;
            cand.matched = true;
            deduplicatedCount++;

            // Se o registro com remessa não possuía hora e o sem remessa possuía, complementa a hora
            const candTime = getRowTime(cand.row);
            const noRemTime = getRowTime(rowNoRem);
            if ((!candTime || candTime === '--:--') && noRemTime && noRemTime !== '--:--') {
              const rawDate = cand.row['DATA'] || cand.row['DATA DA TRAN'] || cand.row['DATA_TRAN'];
              if (typeof rawDate === 'string' && !rawDate.includes(':')) {
                cand.row['HORA'] = noRemTime;
              }
            }
            break;
          }
        }
      }

      if (!matched) {
        keptRowsWithoutRemessa.push(rowNoRem);
      }
    });

    return {
      consolidatedRows: [...rowsWithRemessa, ...keptRowsWithoutRemessa],
      deduplicatedCount
    };
  };

  const parseFileToRows = async (file: File): Promise<TransactionRow[]> => {
    let jsonData: TransactionRow[] = [];
    
    // 1. Tentar leitura binária primeiro (funciona para XLSX, XLS e CSVs em UTF-8)
    const arrayBuffer = await file.arrayBuffer();
    const binaryData = new Uint8Array(arrayBuffer);
    
    try {
      const workbook = XLSX.read(binaryData, { 
        type: 'array', 
        cellDates: true,
        dateNF: 'dd/mm/yyyy hh:mm:ss',
        cellText: false,
        cellNF: false,
        cellStyles: false
      });
      
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const initialData = XLSX.utils.sheet_to_json(worksheet, { defval: '' }) as TransactionRow[];
      
      // Se conseguimos ler e parece ter colunas suficientes, usamos isso
      if (initialData.length > 0 && Object.keys(initialData[0]).length > 2) {
        jsonData = initialData;
      }
    } catch (readErr) {
      console.warn(`Binary read failed for ${file.name}, trying text-based fallbacks...`, readErr);
    }

    // 2. Se não funcionou ou parece ser um CSV com delimitador/encoding especial
    if (jsonData.length === 0) {
      const encodings = ['Windows-1252', 'ISO-8859-1', 'UTF-8'];
      let maxCols = 0;
      let bestData: TransactionRow[] = [];

      for (const encoding of encodings) {
        try {
          const text = await new Promise<string>((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result as string);
            reader.onerror = reject;
            reader.readAsText(file, encoding);
          });

          // Pula arquivos que parecem ser binários (zips/xlsx) lidos como texto
          if (text.startsWith('PK\x03\x04') || text.includes('\x00')) {
            continue;
          }

          const firstLine = text.split('\n')[0];
          const hasSemicolon = firstLine.includes(';');
          
          // Tenta ler com SheetJS tratando como string
          const workbook = XLSX.read(text, { 
            type: 'string', 
            FS: hasSemicolon ? ';' : ',',
            cellDates: true 
          });
          
          const firstSheetName = workbook.SheetNames[0];
          const currentData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { defval: '' }) as TransactionRow[];
          
          if (currentData.length > 0) {
            const cols = Object.keys(currentData[0]).length;
            if (cols > maxCols) {
              maxCols = cols;
              bestData = currentData;
              if (cols > 5) break; 
            }
          }
        } catch (e) {
          console.warn(`Failed reading text with ${encoding} for ${file.name}`, e);
        }
      }
      jsonData = bestData;
    }

    if (jsonData.length === 0) {
      throw new Error(`O arquivo "${file.name}" está vazio ou não pôde ser interpretado.`);
    }

    return jsonData;
  };

  const processFile = async (file: File) => {
    setIsProcessing(true);
    setProcessingStatus(`Lendo ${file.name}...`);
    setFileName(file.name);
    setSourceFiles([file.name]);
    setError(null);
    setData(null);
    setDuplicateGroups([]);

    try {
      const rawData = await parseFileToRows(file);
      const taggedData = rawData.map(row => ({
        ...row,
        ARQUIVO_ORIGEM: file.name
      }));

      setData(taggedData);
      setTransactionCount(taggedData.length);
      analyzeDuplicates(taggedData);
    } catch (err: any) {
      setError(err.message || 'Erro ao processar o arquivo.');
    } finally {
      setIsProcessing(false);
      setProcessingStatus(null);
    }
  };

  const processMultipleFiles = async (filesToProcess: File[]) => {
    if (!filesToProcess || filesToProcess.length === 0) {
      setError('Por favor, adicione pelo menos um arquivo para iniciar a análise.');
      return;
    }

    if (filesToProcess.length === 1) {
      return processFile(filesToProcess[0]);
    }

    setIsProcessing(true);
    setError(null);
    setData(null);
    setDuplicateGroups([]);
    const names = filesToProcess.map(f => f.name);
    setFileName(`Lote Consolidado (${filesToProcess.length} arquivos)`);
    setSourceFiles(names);

    try {
      const allRows: TransactionRow[] = [];
      const failedFiles: string[] = [];

      for (let i = 0; i < filesToProcess.length; i++) {
        const file = filesToProcess[i];
        setProcessingStatus(`Lendo arquivo ${i + 1} de ${filesToProcess.length}: ${file.name}...`);
        try {
          const rows = await parseFileToRows(file);
          rows.forEach(r => {
            allRows.push({
              ...r,
              ARQUIVO_ORIGEM: file.name
            });
          });
        } catch (err: any) {
          console.error(`Falha no arquivo ${file.name}:`, err);
          failedFiles.push(file.name);
        }
      }

      if (allRows.length === 0) {
        throw new Error('Nenhum dado válido pôde ser extraído dos arquivos selecionados.');
      }

      if (failedFiles.length > 0) {
        setError(`Aviso: ${failedFiles.length} arquivo(s) não puderam ser lidos (${failedFiles.join(', ')}). Os demais foram processados com sucesso.`);
      }

      // Se houver mais de um relatório, unifica atendimentos coincidentes priorizando a versão com remessa
      const { consolidatedRows, deduplicatedCount } = consolidateMultipleFileRows(allRows);

      if (deduplicatedCount > 0) {
        setDeduplicationMessage(`${deduplicatedCount} atendimento(s) coincidente(s) entre relatórios foram unificados, trazendo apenas o resultado com remessa.`);
      } else {
        setDeduplicationMessage(null);
      }

      setProcessingStatus(`Cruzando ${consolidatedRows.length} registros entre todos os arquivos...`);
      setData(consolidatedRows);
      setTransactionCount(consolidatedRows.length);
      analyzeDuplicates(consolidatedRows);
    } catch (err: any) {
      setError(err.message || 'Erro ao processar lote consolidado de arquivos.');
    } finally {
      setIsProcessing(false);
      setProcessingStatus(null);
    }
  };

  const addFilesToBatch = (newFiles: FileList | File[]) => {
    const validFiles: File[] = [];
    const existingKeys = new Set(batchFiles.map(f => `${f.name}-${f.size}`));

    Array.from(newFiles).forEach(file => {
      const lower = file.name.toLowerCase();
      if (lower.endsWith('.xlsx') || lower.endsWith('.xls') || lower.endsWith('.csv')) {
        const key = `${file.name}-${file.size}`;
        if (!existingKeys.has(key)) {
          validFiles.push(file);
          existingKeys.add(key);
        }
      }
    });

    if (validFiles.length === 0) {
      if (newFiles.length > 0) {
        setError('Nenhum arquivo novo com extensão compatível (.xlsx, .xls, .csv) foi encontrado.');
      }
      return;
    }

    setBatchFiles(prev => [...prev, ...validFiles]);
    setError(null);
  };

  const removeFileFromBatch = (index: number) => {
    setBatchFiles(prev => prev.filter((_, i) => i !== index));
  };

  const handleMultiDrop = (e: DragEvent) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      addFilesToBatch(e.dataTransfer.files);
    }
  };

  const analyzeDuplicates = (rows: TransactionRow[]) => {
    const groupsMap: Record<string, TransactionRow[]> = {};
    const prestadoresSet = new Set<string>();
    const procedimentosSet = new Set<string>();
    const situacoesSet = new Set<string>();
    const datasSet = new Set<string>();
    const conselhoExecSet = new Set<string>();
    
    const normalizeText = (v: any) => String(v ?? '').trim().toLowerCase();
    const normalizeRaw = (v: any) => String(v ?? '').trim();
    const normalizeNum = (v: any) => {
      if (v === undefined || v === null || v === '') return '';
      const n = parseFloat(String(v));
      return isNaN(n) ? normalizeText(v) : String(n);
    };

    const robustNormalize = (s: any) => {
      if (s === undefined || s === null || s === '') return '';
      return String(s)
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '') // Remove acentos
        .toUpperCase()
        .replace(/\s+/g, '') // Remove espaços
        .replace(/[^A-Z0-9]/g, ''); // Remove tudo que não é alfanumérico
    };

    rows.forEach(row => {
      const keys = Object.keys(row);
      const rawDateVal = getColValue(row, 'DATA');
      const { date } = normalizeDate(rawDateVal);

      const prestadorNome = normalizeRaw(getColValue(row, 'PREST_NOME'));
      const prestadorMatricula = robustNormalize(getColValue(row, 'PREST_MATRICULA'));
      const beneficiarioNome = normalizeRaw(getColValue(row, 'BENE_NOME'));
      const beneficiarioMatriculaOriginal = robustNormalize(getColValue(row, 'BENE_MATRICULA'));
      
      let beneficiarioMatricula = beneficiarioMatriculaOriginal;
      // Secondary check for DV if it wasn't already part of the matricula - stricter check
      const dvCol = keys.find(k => {
        const uk = k.toUpperCase();
        return (uk === 'DV' || uk.includes('_DV') || uk.includes(' DV')) && 
               (uk.includes('MATRICULA') || uk.includes('BENE') || uk.includes('PACIENTE'));
      });

      if (dvCol && row[dvCol]) {
        const dv = String(row[dvCol]).trim();
        if (dv && !beneficiarioMatricula.endsWith(dv)) {
          beneficiarioMatricula = `${beneficiarioMatricula}${dv}`;
        }
      }

      const procedimento = normalizeRaw(getColValue(row, 'PROCEDIMENTO'));
      const especialidade = normalizeRaw(getColValue(row, 'ESPECIALIDADE'));
      const situacao = normalizeRaw(getColValue(row, 'SITUACAO'));
      const tipo = normalizeRaw(getColValue(row, 'TIPO'));
      const quantidade = normalizeNum(getColValue(row, 'QUANTIDADE'));
      const conselho = robustNormalize(getColValue(row, 'CONSELHO'));
      
      const cExecTipo = robustNormalize(getColValue(row, 'TIPO_CONSELHO_EXECUTANTE'));
      const cExecNum = robustNormalize(getColValue(row, 'NUMERO_CONSELHO_EXECUTANTE'));
      const cExecUf = robustNormalize(getColValue(row, 'UF_CONSELHO_EXECUTANTE'));

      const params = [
        robustNormalize(prestadorNome),
        prestadorMatricula,
        date,
        robustNormalize(beneficiarioNome),
        beneficiarioMatricula,
        robustNormalize(procedimento),
        robustNormalize(especialidade),
        robustNormalize(situacao),
        robustNormalize(tipo),
        quantidade,
        conselho,
        cExecTipo,
        cExecNum,
        cExecUf,
      ];

      const fingerprint = params.join('|');
      
      if (!groupsMap[fingerprint]) {
        groupsMap[fingerprint] = [];
      }
      groupsMap[fingerprint].push(row);
    });

    const calculatedGroups: DuplicateGroup[] = Object.entries(groupsMap)
      .filter(([_, groupRows]) => groupRows.length > 1)
      .map(([fingerprint, groupRows]) => {
        const first = groupRows[0];
        const { date } = normalizeDate(getColValue(first, 'DATA'));
        
        const bMat = normalizeText(getColValue(first, 'BENE_MATRICULA'));
        const pMat = normalizeText(getColValue(first, 'PREST_MATRICULA'));

        const prestador = String(getColValue(first, 'PREST_NOME') ?? '');
        const procedimento = String(getColValue(first, 'PROCEDIMENTO') ?? '');
        const situacao = String(getColValue(first, 'SITUACAO') ?? '');

        const cExecTipoVal = String(getColValue(first, 'TIPO_CONSELHO_EXECUTANTE') ?? '');
        const cExecNumVal = String(getColValue(first, 'NUMERO_CONSELHO_EXECUTANTE') ?? '');
        const cExecUfVal = String(getColValue(first, 'UF_CONSELHO_EXECUTANTE') ?? '');
        const cExecStr = [cExecTipoVal, cExecNumVal, cExecUfVal].filter(Boolean).join(' ');
        if (cExecStr) conselhoExecSet.add(cExecStr);

        if (prestador) prestadoresSet.add(prestador);
        if (procedimento) procedimentosSet.add(procedimento);
        if (situacao) situacoesSet.add(situacao);
        if (date) datasSet.add(date);

        return {
          fingerprint,
          rows: groupRows,
          params: {
            prestador,
            prestador_matricula: pMat,
            beneficiario: String(getColValue(first, 'BENE_NOME') ?? ''),
            beneficiario_matricula: bMat,
            procedimento,
            especialidade: String(getColValue(first, 'ESPECIALIDADE') ?? ''),
            situacao,
            tipo: String(getColValue(first, 'TIPO') ?? ''),
            quantidade: String(getColValue(first, 'QUANTIDADE') ?? ''),
            conselho: String(getColValue(first, 'CONSELHO') ?? ''),
            data: date,
            conselhoExecTipo: cExecTipoVal,
            conselhoExecNum: cExecNumVal,
            conselhoExecUf: cExecUfVal
          }
        };
      });

    setUniqueValues({
      prestadores: Array.from(prestadoresSet).sort(),
      procedimentos: Array.from(procedimentosSet).sort(),
      situacoes: Array.from(situacoesSet).sort(),
      datas: Array.from(datasSet).sort((a, b) => {
        const [da, ma, ya] = a.split('/').map(Number);
        const [db, mb, yb] = b.split('/').map(Number);
        return new Date(ya, ma - 1, da).getTime() - new Date(yb, mb - 1, db).getTime();
      }),
      conselhoExecutantes: Array.from(conselhoExecSet).sort(),
    });

    setDuplicateGroups(calculatedGroups);
  };

  const filteredDuplicates = useMemo(() => {
    return duplicateGroups.filter(group => {
      const matchPrestador = !filters.prestador || group.params.prestador.toLowerCase().includes(filters.prestador.toLowerCase());
      const matchBeneficiario = !filters.beneficiario || group.params.beneficiario.toLowerCase().includes(filters.beneficiario.toLowerCase());
      const matchMatricula = !filters.matricula || 
        group.params.beneficiario_matricula.toLowerCase().includes(filters.matricula.toLowerCase()) ||
        group.params.prestador_matricula.toLowerCase().includes(filters.matricula.toLowerCase());
      const matchProcedimento = !filters.procedimento || group.params.procedimento.toLowerCase().includes(filters.procedimento.toLowerCase());
      const matchData = !filters.data || group.params.data.includes(filters.data);
      const matchSituacao = !filters.situacao || group.params.situacao.toLowerCase().includes(filters.situacao.toLowerCase());
      
      const cExecFull = `${group.params.conselhoExecTipo} ${group.params.conselhoExecNum} ${group.params.conselhoExecUf}`.toLowerCase();
      const matchConselhoExec = !filters.conselhoExecutante || 
        cExecFull.includes(filters.conselhoExecutante.toLowerCase()) ||
        group.params.conselhoExecNum.toLowerCase().includes(filters.conselhoExecutante.toLowerCase());

      const intervalInfo = getGroupTimeInterval(group.rows);
      const matchInterval = !filters.intervalCategory || intervalInfo.level === filters.intervalCategory;

      return matchPrestador && matchBeneficiario && matchMatricula && matchProcedimento && matchData && matchSituacao && matchConselhoExec && matchInterval;
    });
  }, [duplicateGroups, filters]);

  const isFiltered = useMemo(() => {
    return Object.values(filters).some(v => v !== '');
  }, [filters]);

  const filteredRowCount = useMemo(() => {
    return filteredDuplicates.reduce((sum, g) => sum + g.rows.length, 0);
  }, [filteredDuplicates]);

  const categoryStats = useMemo(() => {
    let veryHigh = 0;
    let high = 0;
    let medium = 0;
    let low = 0;
    filteredDuplicates.forEach(g => {
      const info = getGroupTimeInterval(g.rows);
      if (info.level === 'very_high') veryHigh++;
      else if (info.level === 'high') high++;
      else if (info.level === 'medium') medium++;
      else low++;
    });
    return { veryHigh, high, medium, low, total: filteredDuplicates.length };
  }, [filteredDuplicates]);

  const dashboardStats = useMemo(() => {
    const daysMap: Record<string, number> = {};
    const hoursMap: Record<string, number> = {};
    const providersMap: Record<string, number> = {};
    const proceduresMap: Record<string, number> = {};

    filteredDuplicates.forEach(group => {
      const d = group.params.data || 'Data não informada';
      daysMap[d] = (daysMap[d] || 0) + 1;

      const prov = group.params.prestador || 'Não informado';
      providersMap[prov] = (providersMap[prov] || 0) + 1;

      const proc = group.params.procedimento || 'Não informado';
      proceduresMap[proc] = (proceduresMap[proc] || 0) + 1;

      group.rows.forEach(row => {
        const time = getRowTime(row);
        if (time && time !== '--:--' && time.includes(':')) {
          const hourGroup = time.split(':')[0].padStart(2, '0') + ':00';
          hoursMap[hourGroup] = (hoursMap[hourGroup] || 0) + 1;
        }
      });
    });

    const topDays = Object.entries(daysMap)
      .map(([day, count]) => ({ day, count }))
      .sort((a, b) => b.count - a.count);

    const topHours = Object.entries(hoursMap)
      .map(([hour, count]) => ({ hour, count }))
      .sort((a, b) => b.count - a.count);

    const topProviders = Object.entries(providersMap)
      .map(([provider, count]) => ({ provider, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);

    const topProcedures = Object.entries(proceduresMap)
      .map(([procedure, count]) => ({ procedure, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);

    return { topDays, topHours, topProviders, topProcedures };
  }, [filteredDuplicates]);

  // Diagnóstico Estatístico de Padrões e Comparativos para Auditoria
  const itStats = useMemo<ITDiagnosticsStats>(() => {
    let deltaZeroCount = 0;
    let delta1to5Count = 0;
    let delta6to30Count = 0;
    let deltaOver30Count = 0;

    let sameProtocolCount = 0;
    let differentProtocolCount = 0;
    let noProtocolCount = 0;

    let sameRemessaCount = 0;
    let differentRemessaCount = 0;
    let noRemessaCount = 0;

    const providerCounts: Record<string, number> = {};
    const procedureCounts: Record<string, number> = {};
    const beneficiaryCounts: Record<string, number> = {};
    const situationCounts: Record<string, number> = {};
    const dayOfWeekCounts: Record<number, number> = { 0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0 };
    const monthCounts: Record<string, number> = {};
    const dayOfMonthCounts: Record<number, number> = {};
    for (let i = 1; i <= 31; i++) dayOfMonthCounts[i] = 0;

    const monthCycleCounts: Record<string, number> = {
      'Início do Mês (Dias 01 a 10)': 0,
      'Meio do Mês (Dias 11 a 20)': 0,
      'Fechamento de Mês (Dias 21 a 31)': 0
    };

    const shiftCounts: Record<string, number> = {
      'Madrugada (00h-06h)': 0,
      'Manhã (06h-12h)': 0,
      'Tarde (12h-18h)': 0,
      'Noite (18h-24h)': 0
    };

    filteredDuplicates.forEach(group => {
      // Delta analysis
      const info = getGroupTimeInterval(group.rows);
      if (info.minDiff === 0) {
        deltaZeroCount++;
      } else if (info.minDiff !== null && info.minDiff > 0 && info.minDiff <= 5) {
        delta1to5Count++;
      } else if (info.minDiff !== null && info.minDiff > 5 && info.minDiff <= 30) {
        delta6to30Count++;
      } else {
        deltaOver30Count++;
      }

      // Protocol analysis
      const protocols = new Set<string>();
      group.rows.forEach(r => {
        const p = getColValue(r, 'PROTOCOLO');
        if (p !== undefined && p !== null && String(p).trim() !== '') {
          protocols.add(String(p).trim());
        }
      });
      if (protocols.size === 0) {
        noProtocolCount++;
      } else if (protocols.size === 1) {
        sameProtocolCount++;
      } else {
        differentProtocolCount++;
      }

      // Remessa analysis
      const remessas = new Set<string>();
      group.rows.forEach(r => {
        const rem = getColValue(r, 'REMESSA');
        if (rem !== undefined && rem !== null && String(rem).trim() !== '') {
          remessas.add(String(rem).trim());
        }
      });
      if (remessas.size === 0) {
        noRemessaCount++;
      } else if (remessas.size === 1) {
        sameRemessaCount++;
      } else {
        differentRemessaCount++;
      }

      // Provider concentration
      const prov = group.params.prestador || 'Não informado';
      providerCounts[prov] = (providerCounts[prov] || 0) + 1;

      // Procedure frequency
      const proc = group.params.procedimento || 'Não informado';
      procedureCounts[proc] = (procedureCounts[proc] || 0) + 1;

      // Beneficiary frequency
      const bene = group.params.beneficiario || 'Não informado';
      beneficiaryCounts[bene] = (beneficiaryCounts[bene] || 0) + 1;

      // Situation frequency
      const sit = group.params.situacao || 'Não informada';
      situationCounts[sit] = (situationCounts[sit] || 0) + 1;

      // Date analysis (Day of week, Month & Day of month/cycle)
      const dStr = group.params.data || '';
      if (dStr) {
        const parts = dStr.split('/');
        if (parts.length === 3) {
          const dd = parseInt(parts[0], 10);
          const mm = parseInt(parts[1], 10) - 1;
          const yyyy = parseInt(parts[2], 10);
          if (!isNaN(dd) && !isNaN(mm) && !isNaN(yyyy)) {
            const dt = new Date(yyyy, mm, dd);
            const dow = dt.getDay();
            if (!isNaN(dow)) {
              dayOfWeekCounts[dow] = (dayOfWeekCounts[dow] || 0) + 1;
            }
            const monthNames = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
            const mKey = `${monthNames[mm] || 'Mês'}/${yyyy}`;
            monthCounts[mKey] = (monthCounts[mKey] || 0) + 1;

            if (dd >= 1 && dd <= 31) {
              dayOfMonthCounts[dd] = (dayOfMonthCounts[dd] || 0) + 1;
              if (dd <= 10) monthCycleCounts['Início do Mês (Dias 01 a 10)']++;
              else if (dd <= 20) monthCycleCounts['Meio do Mês (Dias 11 a 20)']++;
              else monthCycleCounts['Fechamento de Mês (Dias 21 a 31)']++;
            }
          }
        }
      }

      // Shift analysis
      group.rows.forEach(r => {
        const t = getRowTime(r);
        if (t && t.includes(':')) {
          const hh = parseInt(t.split(':')[0], 10);
          if (!isNaN(hh)) {
            if (hh >= 0 && hh < 6) shiftCounts['Madrugada (00h-06h)']++;
            else if (hh >= 6 && hh < 12) shiftCounts['Manhã (06h-12h)']++;
            else if (hh >= 12 && hh < 18) shiftCounts['Tarde (12h-18h)']++;
            else if (hh >= 18 && hh < 24) shiftCounts['Noite (18h-24h)']++;
          }
        }
      });
    });

    const total = filteredDuplicates.length;
    const sortedProviders = Object.entries(providerCounts).sort((a, b) => b[1] - a[1]);
    const topProviderName = sortedProviders.length > 0 ? sortedProviders[0][0] : 'Nenhum';
    const topProviderCount = sortedProviders.length > 0 ? sortedProviders[0][1] : 0;
    const topProviderConcentrationPercent = total > 0 ? Math.round((topProviderCount / total) * 100) : 0;

    const top3Sum = sortedProviders.slice(0, 3).reduce((acc, [, c]) => acc + c, 0);
    const top3ProvidersSharePercent = total > 0 ? Math.round((top3Sum / total) * 100) : 0;

    const sortedProcedures = Object.entries(procedureCounts).sort((a, b) => b[1] - a[1]);
    const sortedBeneficiaries = Object.entries(beneficiaryCounts).sort((a, b) => b[1] - a[1]);
    const sortedSituations = Object.entries(situationCounts).sort((a, b) => b[1] - a[1]);

    const dayLabels = ['Domingo', 'Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado'];
    const dayOfWeekDistribution = [1, 2, 3, 4, 5, 6, 0].map(idx => ({
      dayName: dayLabels[idx],
      count: dayOfWeekCounts[idx] || 0,
      percent: total > 0 ? Math.round(((dayOfWeekCounts[idx] || 0) / total) * 100) : 0
    }));

    const dayOfMonthDistribution = Array.from({ length: 31 }, (_, i) => i + 1).map(d => ({
      dayNum: d,
      dayLabel: `Dia ${String(d).padStart(2, '0')}`,
      count: dayOfMonthCounts[d] || 0,
      percent: total > 0 ? Math.round(((dayOfMonthCounts[d] || 0) / total) * 100) : 0
    }));

    const monthCycleDistribution = Object.entries(monthCycleCounts).map(([cycleName, count]) => ({
      cycleName,
      count,
      percent: total > 0 ? Math.round((count / total) * 100) : 0
    }));

    const monthlyDistribution = Object.entries(monthCounts).map(([monthLabel, count]) => ({
      monthLabel,
      count,
      percent: total > 0 ? Math.round((count / total) * 100) : 0
    }));

    const totalShiftOccurrences = Object.values(shiftCounts).reduce((acc, c) => acc + c, 0);
    const shiftDistribution = Object.entries(shiftCounts).map(([shiftName, count]) => ({
      shiftName,
      count,
      percent: totalShiftOccurrences > 0 ? Math.round((count / totalShiftOccurrences) * 100) : 0
    }));

    const topProvidersList = sortedProviders.slice(0, 10).map(([provider, count]) => ({
      provider,
      count,
      percent: total > 0 ? Math.round((count / total) * 100) : 0
    }));

    const topProceduresList = sortedProcedures.slice(0, 10).map(([procedure, count]) => ({
      procedure,
      count,
      percent: total > 0 ? Math.round((count / total) * 100) : 0
    }));

    const topBeneficiariesList = sortedBeneficiaries.slice(0, 10).map(([beneficiary, count]) => ({
      beneficiary,
      count,
      percent: total > 0 ? Math.round((count / total) * 100) : 0
    }));

    const situationsDistribution = sortedSituations.slice(0, 10).map(([situation, count]) => ({
      situation,
      count,
      percent: total > 0 ? Math.round((count / total) * 100) : 0
    }));

    return {
      totalGroups: total,
      deltaZeroCount,
      instantPercent: total > 0 ? Math.round((deltaZeroCount / total) * 100) : 0,
      delta1to5Count,
      microPercent: total > 0 ? Math.round((delta1to5Count / total) * 100) : 0,
      delta6to30Count,
      shortPercent: total > 0 ? Math.round((delta6to30Count / total) * 100) : 0,
      deltaOver30Count,
      extendedPercent: total > 0 ? Math.round((deltaOver30Count / total) * 100) : 0,
      sameProtocolCount,
      sameProtocolPercent: total > 0 ? Math.round((sameProtocolCount / total) * 100) : 0,
      differentProtocolCount,
      differentProtocolPercent: total > 0 ? Math.round((differentProtocolCount / total) * 100) : 0,
      noProtocolCount,
      sameRemessaCount,
      sameRemessaPercent: total > 0 ? Math.round((sameRemessaCount / total) * 100) : 0,
      differentRemessaCount,
      differentRemessaPercent: total > 0 ? Math.round((differentRemessaCount / total) * 100) : 0,
      noRemessaCount,
      topProviderConcentrationPercent,
      top3ProvidersSharePercent,
      topProviderName,
      topProviderCount,
      topProvidersList,
      topProceduresList,
      dayOfWeekDistribution,
      dayOfMonthDistribution,
      monthCycleDistribution,
      monthlyDistribution,
      shiftDistribution,
      topBeneficiariesList,
      situationsDistribution
    };
  }, [filteredDuplicates]);

  const exportDashboardReport = () => {
    if (filteredDuplicates.length === 0) return;

    const wb = XLSX.utils.book_new();

    const summaryData = [
      { Métrica: isFiltered ? 'Total de Registros nos Filtros Atuais' : 'Total de Registros Analisados', Valor: isFiltered ? filteredRowCount : transactionCount },
      { Métrica: isFiltered ? 'Grupos de Duplicidade Filtrados' : 'Total de Grupos de Duplicidade', Valor: filteredDuplicates.length },
      { Métrica: 'Duplicidades Muito Suspeitas (Mesmo Horário - 0m)', Valor: categoryStats.veryHigh },
      { Métrica: 'Duplicidades Bem Prováveis (≤ 30min)', Valor: categoryStats.high },
      { Métrica: 'Duplicidades Pouco Menos Prováveis (30m - 1h)', Valor: categoryStats.medium },
      { Métrica: 'Outras Duplicidades (> 1h)', Valor: categoryStats.low },
    ];
    const wsSummary = XLSX.utils.json_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, wsSummary, "Resumo Executivo");

    const wsDays = XLSX.utils.json_to_sheet(dashboardStats.topDays.map(d => ({ 'Data': d.day, 'Quantidade de Duplicidades': d.count })));
    XLSX.utils.book_append_sheet(wb, wsDays, "Picos por Dia");

    const wsHours = XLSX.utils.json_to_sheet(dashboardStats.topHours.map(h => ({ 'Faixa Horária': h.hour, 'Ocorrências': h.count })));
    XLSX.utils.book_append_sheet(wb, wsHours, "Picos por Horário");

    const wsProv = XLSX.utils.json_to_sheet(dashboardStats.topProviders.map(p => ({ 'Prestador': p.provider, 'Duplicidades': p.count })));
    XLSX.utils.book_append_sheet(wb, wsProv, "Top Prestadores");

    const wsDow = XLSX.utils.json_to_sheet(itStats.dayOfWeekDistribution.map(d => ({ 'Dia da Semana': d.dayName, 'Ocorrências': d.count, 'Percentual': `${d.percent}%` })));
    XLSX.utils.book_append_sheet(wb, wsDow, "Dia da Semana");

    const wsDom = XLSX.utils.json_to_sheet(itStats.dayOfMonthDistribution.map(d => ({ 'Dia do Mês': d.dayLabel, 'Ocorrências': d.count, 'Percentual': `${d.percent}%` })));
    XLSX.utils.book_append_sheet(wb, wsDom, "Dias do Mes (01-31)");

    const wsCycle = XLSX.utils.json_to_sheet(itStats.monthCycleDistribution.map(c => ({ 'Etapa do Mês': c.cycleName, 'Ocorrências': c.count, 'Percentual': `${c.percent}%` })));
    XLSX.utils.book_append_sheet(wb, wsCycle, "Ciclo Mensal");

    const wsShift = XLSX.utils.json_to_sheet(itStats.shiftDistribution.map(s => ({ 'Turno': s.shiftName, 'Ocorrências': s.count, 'Percentual': `${s.percent}%` })));
    XLSX.utils.book_append_sheet(wb, wsShift, "Turnos do Dia");

    const itSheetData = [
      { Métrica: 'Total de Duplicidades Filtradas', Valor: filteredDuplicates.length },
      { Métrica: 'Duplicidades Simultâneas (Δt = 0m)', Valor: itStats.deltaZeroCount },
      { Métrica: 'Micro-intervalo (1 a 5m)', Valor: itStats.delta1to5Count },
      { Métrica: 'Mesma Sessão de Atendimento (6 a 30m)', Valor: itStats.delta6to30Count },
      { Métrica: 'Reemissão Tardia (> 30m)', Valor: itStats.deltaOver30Count },
      { Métrica: 'Mesmo Número de Protocolo', Valor: itStats.sameProtocolCount },
      { Métrica: 'Protocolos Diferentes', Valor: itStats.differentProtocolCount },
      { Métrica: 'Protocolo Não Informado', Valor: itStats.noProtocolCount },
      { Métrica: 'Concentração Prestador Líder', Valor: `${itStats.topProviderConcentrationPercent}% (${itStats.topProviderName})` },
      { Métrica: 'Concentração Top 3 Prestadores', Valor: `${itStats.top3ProvidersSharePercent}%` }
    ];
    const wsIT = XLSX.utils.json_to_sheet(itSheetData);
    XLSX.utils.book_append_sheet(wb, wsIT, "Padroes_Comparativos");

    XLSX.writeFile(wb, isFiltered ? "Dashboard_Duplicidades_Filtradas.xlsx" : "Dashboard_Auditoria_Duplicidades.xlsx");
  };

  const handleDrop = (e: DragEvent) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls') || file.name.endsWith('.csv'))) {
      processFile(file);
    } else {
      setError('Por favor, envie um arquivo Excel (.xlsx, .xls) ou CSV.');
    }
  };

  const fileHasRemessa = useMemo(() => {
    if (!data || data.length === 0) return false;
    const hasHeader = getColValue(data[0], 'REMESSA') !== undefined;
    const hasValue = data.some(row => {
      const v = getColValue(row, 'REMESSA');
      return v !== undefined && v !== null && String(v).trim() !== '';
    });
    return hasHeader || hasValue;
  }, [data]);

  const fileHasProtocolo = useMemo(() => {
    if (!data || data.length === 0) return false;
    const hasHeader = getColValue(data[0], 'PROTOCOLO') !== undefined;
    const hasValue = data.some(row => {
      const v = getColValue(row, 'PROTOCOLO');
      return v !== undefined && v !== null && String(v).trim() !== '';
    });
    return hasHeader || hasValue;
  }, [data]);

  const exportResults = () => {
    if (filteredDuplicates.length === 0) return;

    const exportData = filteredDuplicates.flatMap(group => {
      const intervalInfo = getGroupTimeInterval(group.rows);
      return group.rows.map(row => {
        const { date } = normalizeDate(getColValue(row, 'DATA'));
        const time = getRowTime(row);
        const remessaVal = getColValue(row, 'REMESSA');
        const protocoloVal = getColValue(row, 'PROTOCOLO');

        const exportRow: Record<string, any> = {};

        // Se o arquivo possuir número de remessa ou protocolo, insere no início para fácil visualização
        if (fileHasRemessa) {
          exportRow['REMESSA'] = remessaVal !== undefined && remessaVal !== null ? remessaVal : '';
        }
        if (fileHasProtocolo) {
          exportRow['PROTOCOLO'] = protocoloVal !== undefined && protocoloVal !== null ? protocoloVal : '';
        }

        // Incorpora todas as colunas originais do relatório
        Object.assign(exportRow, row);

        // Garante que as colunas padronizadas 'REMESSA' e 'PROTOCOLO' fiquem preenchidas
        if (fileHasRemessa) {
          exportRow['REMESSA'] = remessaVal !== undefined && remessaVal !== null ? remessaVal : '';
        }
        if (fileHasProtocolo) {
          exportRow['PROTOCOLO'] = protocoloVal !== undefined && protocoloVal !== null ? protocoloVal : '';
        }

        exportRow['DATA_NORMALIZADA'] = date;
        exportRow['HORARIO'] = time || '--:--';
        exportRow['ID_BENEFICIARIO'] = getColValue(row, 'BENE_MATRICULA') ?? '';
        exportRow['ID_PRESTADOR'] = getColValue(row, 'PREST_MATRICULA') ?? '';
        exportRow['CATEGORIA_INTERVALO'] = intervalInfo.label || 'Outros';
        exportRow['STATUS_AUDITORIA'] = 'Suspeita de Duplicidade';
        if (sourceFiles.length > 1 || row['ARQUIVO_ORIGEM']) {
          exportRow['ARQUIVO_ORIGEM'] = row['ARQUIVO_ORIGEM'] || fileName || '';
        }

        return exportRow;
      });
    });

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Duplicidades");

    const itSheetData = [
      { Métrica: 'Total de Duplicidades Filtradas', Valor: filteredDuplicates.length },
      { Métrica: 'Duplicidades Simultâneas (Δt = 0m)', Valor: itStats.deltaZeroCount },
      { Métrica: 'Micro-intervalo (1 a 5m)', Valor: itStats.delta1to5Count },
      { Métrica: 'Mesma Sessão de Atendimento (6 a 30m)', Valor: itStats.delta6to30Count },
      { Métrica: 'Reemissão Tardia (> 30m)', Valor: itStats.deltaOver30Count },
      { Métrica: 'Mesmo Número de Protocolo', Valor: itStats.sameProtocolCount },
      { Métrica: 'Protocolos Diferentes', Valor: itStats.differentProtocolCount },
      { Métrica: 'Protocolo Não Informado', Valor: itStats.noProtocolCount },
      { Métrica: 'Concentração Prestador Líder', Valor: `${itStats.topProviderConcentrationPercent}% (${itStats.topProviderName})` },
      { Métrica: 'Concentração Top 3 Prestadores', Valor: `${itStats.top3ProvidersSharePercent}%` }
    ];
    const wsIT = XLSX.utils.json_to_sheet(itSheetData);
    XLSX.utils.book_append_sheet(wb, wsIT, "Padroes_Comparativos");

    XLSX.writeFile(wb, "Relatorio_Duplicidades_Filtradas.xlsx");
  };

  const handleExportFormatted = async () => {
    if (filteredDuplicates.length === 0) return;

    // Resumo dos filtros para personalização no cabeçalho do arquivo Excel
    const activeFiltersList: string[] = [];
    if (filters.prestador) activeFiltersList.push(`Prestador: ${filters.prestador}`);
    if (filters.beneficiario) activeFiltersList.push(`Beneficiário: ${filters.beneficiario}`);
    if (filters.matricula) activeFiltersList.push(`Matrícula: ${filters.matricula}`);
    if (filters.procedimento) activeFiltersList.push(`Procedimento: ${filters.procedimento}`);
    if (filters.data) activeFiltersList.push(`Data: ${filters.data}`);
    if (filters.situacao) activeFiltersList.push(`Situação: ${filters.situacao}`);
    if (filters.conselhoExecutante) activeFiltersList.push(`CRM: ${filters.conselhoExecutante}`);
    if (filters.intervalCategory) {
      const catLabel = filters.intervalCategory === 'very_high'
        ? 'Mesmo Horário (Muito Suspeito)'
        : filters.intervalCategory === 'high'
        ? '≤ 30min (Bem Provável)'
        : filters.intervalCategory === 'medium'
        ? '30m - 1h (Pouco Menos)'
        : '> 1h';
      activeFiltersList.push(`Intervalo: ${catLabel}`);
    }

    await exportFormattedExcelReport({
      fileName: fileName || 'Relatório de Auditoria',
      transactionCount: isFiltered ? filteredRowCount : transactionCount,
      duplicateGroups: filteredDuplicates,
      fileHasRemessa,
      fileHasProtocolo,
      selectedPrestadorFilter: activeFiltersList.length > 0 ? activeFiltersList.join(' | ') : undefined,
      getColValue,
      normalizeDate,
      getGroupTimeInterval,
      getRowTime,
      hasMultiOriginFiles: sourceFiles.length > 1,
      itStats
    });
  };

  const handleExportFormattedDashboard = async () => {
    if (filteredDuplicates.length === 0) return;
    await exportFormattedDashboardReport({
      transactionCount: isFiltered ? filteredRowCount : transactionCount,
      duplicateGroups: filteredDuplicates,
      categoryStats,
      dashboardStats,
      itStats
    });
  };

  const handleExportFormattedIT = async () => {
    if (filteredDuplicates.length === 0) return;
    await exportFormattedITReport({
      transactionCount: isFiltered ? filteredRowCount : transactionCount,
      itStats,
      fileName: fileName || 'Padrões e Comparativos'
    });
  };

  const resetFilters = () => {
    setFilters({
      prestador: '',
      beneficiario: '',
      matricula: '',
      procedimento: '',
      data: '',
      situacao: '',
      intervalCategory: '',
      conselhoExecutante: ''
    });
  };

  const resetToHome = () => {
    setData(null);
    setDuplicateGroups([]);
    setFileName(null);
    setSourceFiles([]);
    setBatchFiles([]);
    setProcessingStatus(null);
    setDeduplicationMessage(null);
    setTransactionCount(0);
    resetFilters();
  };

  return (
    <div className="min-h-screen bg-[#0A0B0D] text-[#E0E0E0] font-sans flex flex-col overflow-x-hidden">
      {/* Header Section */}
      <header className="flex flex-col md:flex-row items-center justify-between px-8 py-6 border-b border-white/10 gap-6">
        <div 
          className="cursor-pointer group"
          onClick={resetToHome}
        >
          <h1 className="text-2xl font-light tracking-widest text-white uppercase group-hover:text-emerald-400 transition-colors">Auditoria de Duplicidade</h1>
          <p className="text-xs text-gray-500 mt-1 uppercase tracking-tighter italic">Sistema de Verificação de Integridade de Relatórios{fileName ? ` • ${fileName}` : ''}</p>
        </div>
        <div className="flex items-center gap-6">
          {sourceFiles.length > 1 && (
            <div className="hidden sm:flex items-center gap-2 px-3 py-1.5 bg-emerald-500/10 border border-emerald-500/30 rounded text-emerald-300 text-xs font-mono" title={sourceFiles.join(', ')}>
              <Files size={14} className="text-emerald-400" />
              <span>Lote: {sourceFiles.length} arquivos</span>
            </div>
          )}
          {data && (
            <button 
              onClick={resetToHome}
              className="flex items-center gap-2 px-4 py-2 text-[10px] uppercase tracking-widest text-[#E0E0E0]/60 hover:text-white border border-white/10 rounded hover:bg-white/5 transition-all cursor-pointer"
            >
              <Trash2 size={14} /> Novo Relatório
            </button>
          )}
          <div className="flex gap-4">
            <div className="px-4 py-2 bg-white/5 border border-white/10 rounded flex flex-col min-w-[140px]">
              <span className="text-[10px] text-gray-500 uppercase">
                {isFiltered ? 'Registros Filtrados' : 'Registros Analisados'}
              </span>
              <div className="flex items-baseline gap-1.5">
                <span className="text-xl font-mono text-white leading-none">
                  {isFiltered ? filteredRowCount : transactionCount}
                </span>
                {isFiltered && (
                  <span className="text-[10px] font-mono text-gray-400" title={`Total geral do arquivo: ${transactionCount}`}>
                    / {transactionCount}
                  </span>
                )}
              </div>
            </div>
            <div className="px-4 py-2 bg-red-900/20 border border-red-500/30 rounded flex flex-col min-w-[140px]">
              <span className="text-[10px] text-red-400 uppercase">
                {isFiltered ? 'Duplicidades Filtradas' : 'Possíveis Duplicidades'}
              </span>
              <div className="flex items-baseline gap-1.5">
                <span className="text-xl font-mono text-red-500 leading-none">
                  {filteredDuplicates.length}
                </span>
                {isFiltered && (
                  <span className="text-[10px] font-mono text-red-400/70" title={`Total geral de grupos de duplicidade: ${duplicateGroups.length}`}>
                    / {duplicateGroups.length}
                  </span>
                )}
              </div>
            </div>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 p-8 flex flex-col gap-6">
        {!data && !isProcessing && (
          <motion.div 
            initial={{ opacity: 0, scale: 0.98 }}
            animate={{ opacity: 1, scale: 1 }}
            className="flex-1 flex flex-col items-center justify-center border border-white/5 rounded-2xl bg-white/[0.01] p-6 md:p-12"
          >
            {/* Mode Switcher Tabs */}
            <div className="flex items-center gap-2 p-1.5 bg-white/5 border border-white/10 rounded-xl mb-8">
              <button
                type="button"
                onClick={() => setUploadMode('single')}
                className={cn(
                  "flex items-center gap-2 px-5 py-2.5 rounded-lg text-xs uppercase tracking-wider transition-all cursor-pointer",
                  uploadMode === 'single'
                    ? "bg-white/15 text-white font-semibold shadow-md border border-white/20"
                    : "text-gray-400 hover:text-white hover:bg-white/5"
                )}
              >
                <FileSpreadsheet size={16} /> Arquivo Único
              </button>
              <button
                type="button"
                onClick={() => setUploadMode('multi')}
                className={cn(
                  "flex items-center gap-2 px-5 py-2.5 rounded-lg text-xs uppercase tracking-wider transition-all cursor-pointer",
                  uploadMode === 'multi'
                    ? "bg-emerald-500/20 text-emerald-300 font-semibold shadow-md border border-emerald-500/30"
                    : "text-gray-400 hover:text-white hover:bg-white/5"
                )}
              >
                <Files size={16} /> Múltiplos Arquivos (Lote Consolidado)
                <span className="text-[9px] bg-emerald-500/30 text-emerald-300 px-1.5 py-0.5 rounded font-mono font-bold">NOVO</span>
              </button>
            </div>

            {/* Mode 1: Arquivo Único (Inalterado) */}
            {uploadMode === 'single' ? (
              <div 
                onDragOver={(e) => e.preventDefault()}
                onDrop={handleDrop}
                onClick={() => fileInputRef.current?.click()}
                className="group cursor-pointer flex flex-col items-center p-14 border-2 border-dashed border-white/10 hover:border-white/30 rounded-xl transition-all max-w-xl w-full"
              >
                <div className="w-20 h-20 bg-white/5 rounded-full flex items-center justify-center mb-6 group-hover:scale-110 transition-transform">
                  <Upload size={32} className="text-white/60 group-hover:text-white" />
                </div>
                <h2 className="text-xl font-light text-white mb-2 uppercase tracking-wide">Importar Relatório Único</h2>
                <p className="text-xs text-gray-500 font-mono mb-8">FORMATOS: .XLSX, .XLS, .CSV</p>
                <input 
                  type="file" 
                  ref={fileInputRef}
                  className="hidden" 
                  accept=".xlsx, .xls, .csv"
                  onChange={(e) => e.target.files?.[0] && processFile(e.target.files[0])}
                />
                <button className="px-8 py-3 bg-white/10 border border-white/10 text-xs uppercase tracking-widest text-white hover:bg-white/20 transition-all cursor-pointer">
                  Explorar Arquivo
                </button>
              </div>
            ) : (
              /* Mode 2: Múltiplos Arquivos (Lote Consolidado) */
              <div className="w-full max-w-3xl flex flex-col gap-6">
                <div className="bg-emerald-500/5 border border-emerald-500/20 rounded-xl p-5 text-left flex items-start gap-3">
                  <div className="p-2 bg-emerald-500/10 rounded-lg text-emerald-400 shrink-0 mt-0.5">
                    <Layers size={20} />
                  </div>
                  <div>
                    <h3 className="text-sm font-semibold text-emerald-400 uppercase tracking-wider">Auditoria Consolidada de Lote de Arquivos</h3>
                    <p className="text-xs text-gray-400 font-light mt-1 leading-relaxed">
                      Envie múltiplos arquivos ao mesmo tempo para cruzar dados de diferentes meses, prestadores ou faturamentos. O sistema identificará duplicidades tanto dentro de cada arquivo quanto re-cobranças entre arquivos distintos.
                    </p>
                  </div>
                </div>

                <input 
                  type="file" 
                  ref={multiFileInputRef}
                  className="hidden" 
                  multiple
                  accept=".xlsx, .xls, .csv"
                  onChange={(e) => {
                    if (e.target.files && e.target.files.length > 0) {
                      addFilesToBatch(e.target.files);
                    }
                  }}
                />
                <input 
                  type="file" 
                  ref={addMoreFileInputRef}
                  className="hidden" 
                  multiple
                  accept=".xlsx, .xls, .csv"
                  onChange={(e) => {
                    if (e.target.files && e.target.files.length > 0) {
                      addFilesToBatch(e.target.files);
                    }
                  }}
                />

                {batchFiles.length === 0 ? (
                  <div 
                    onDragOver={(e) => e.preventDefault()}
                    onDrop={handleMultiDrop}
                    onClick={() => multiFileInputRef.current?.click()}
                    className="group cursor-pointer flex flex-col items-center p-14 border-2 border-dashed border-emerald-500/30 hover:border-emerald-500/60 rounded-xl bg-emerald-500/[0.02] hover:bg-emerald-500/[0.04] transition-all text-center"
                  >
                    <div className="w-20 h-20 bg-emerald-500/10 rounded-full flex items-center justify-center mb-6 group-hover:scale-110 transition-transform">
                      <Files size={32} className="text-emerald-400 group-hover:text-emerald-300" />
                    </div>
                    <h2 className="text-xl font-light text-white mb-2 uppercase tracking-wide">Selecionar Múltiplos Arquivos</h2>
                    <p className="text-xs text-gray-400 font-light max-w-md mb-2">
                      Arraste e solte vários relatórios aqui (.xlsx, .xls, .csv) de uma só vez ou clique para explorar.
                    </p>
                    <p className="text-[10px] text-gray-500 font-mono mb-8">SUPORTE A SELEÇÃO MÚLTIPLA DE ARQUIVOS</p>
                    <button className="px-8 py-3 bg-emerald-600/30 border border-emerald-500/50 text-xs uppercase tracking-widest text-emerald-300 hover:bg-emerald-600/40 transition-all cursor-pointer font-semibold shadow-lg shadow-emerald-950/50">
                      Selecionar Arquivos do Lote
                    </button>
                  </div>
                ) : (
                  <div className="bg-white/[0.02] border border-white/10 rounded-xl p-6 flex flex-col gap-5">
                    <div className="flex items-center justify-between pb-4 border-b border-white/10">
                      <div className="flex items-center gap-2">
                        <Files size={18} className="text-emerald-400" />
                        <span className="text-sm font-semibold text-white uppercase tracking-wider">
                          Fila do Lote ({batchFiles.length} {batchFiles.length === 1 ? 'arquivo' : 'arquivos'})
                        </span>
                      </div>
                      <div className="flex items-center gap-3">
                        <button
                          type="button"
                          onClick={() => addMoreFileInputRef.current?.click()}
                          className="flex items-center gap-1.5 px-3 py-1.5 text-xs text-emerald-400 hover:text-emerald-300 bg-emerald-500/10 hover:bg-emerald-500/20 border border-emerald-500/30 rounded transition-all cursor-pointer"
                        >
                          <Plus size={14} /> Adicionar Mais
                        </button>
                        <button
                          type="button"
                          onClick={() => setBatchFiles([])}
                          className="flex items-center gap-1.5 px-3 py-1.5 text-xs text-red-400 hover:text-red-300 bg-red-500/10 hover:bg-red-500/20 border border-red-500/20 rounded transition-all cursor-pointer"
                        >
                          <Trash2 size={13} /> Limpar
                        </button>
                      </div>
                    </div>

                    {/* Files List */}
                    <div className="max-h-64 overflow-y-auto divide-y divide-white/5 pr-1 custom-scrollbar">
                      {batchFiles.map((f, i) => (
                        <div key={`${f.name}-${i}`} className="py-2.5 px-3 flex items-center justify-between hover:bg-white/[0.02] rounded transition-colors">
                          <div className="flex items-center gap-3 min-w-0 pr-4">
                            <span className="text-[10px] font-mono text-gray-500 w-5 text-right">{i + 1}.</span>
                            <FileSpreadsheet size={16} className="text-emerald-400 shrink-0" />
                            <div className="min-w-0">
                              <p className="text-xs font-medium text-gray-200 truncate" title={f.name}>{f.name}</p>
                              <span className="text-[9px] font-mono text-gray-500">
                                {f.size < 1024 * 1024 ? `${(f.size / 1024).toFixed(1)} KB` : `${(f.size / (1024 * 1024)).toFixed(2)} MB`}
                              </span>
                            </div>
                          </div>
                          <button
                            type="button"
                            onClick={() => removeFileFromBatch(i)}
                            className="text-gray-500 hover:text-red-400 p-1 transition-colors"
                            title="Remover arquivo da fila"
                          >
                            <X size={14} />
                          </button>
                        </div>
                      ))}
                    </div>

                    {/* Drop more area */}
                    <div 
                      onDragOver={(e) => e.preventDefault()}
                      onDrop={handleMultiDrop}
                      className="border border-dashed border-white/10 hover:border-white/20 p-3 rounded-lg text-center text-xs text-gray-500 cursor-pointer"
                      onClick={() => addMoreFileInputRef.current?.click()}
                    >
                      + Arraste mais arquivos aqui ou clique para complementar o lote
                    </div>

                    {/* Process Action CTA */}
                    <button
                      type="button"
                      onClick={() => processMultipleFiles(batchFiles)}
                      className="w-full py-4 bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs uppercase tracking-widest rounded-lg transition-all shadow-xl shadow-emerald-950/60 flex items-center justify-center gap-2 cursor-pointer"
                    >
                      <Layers size={16} /> Analisar Duplicidades no Lote Consolidado ({batchFiles.length} {batchFiles.length === 1 ? 'arquivo' : 'arquivos'})
                    </button>
                  </div>
                )}
              </div>
            )}

            <div className="mt-16 w-full max-w-4xl grid grid-cols-1 md:grid-cols-2 gap-8 text-left border-t border-white/5 pt-12">
              <div className="space-y-4">
                <h3 className="text-xs font-bold uppercase tracking-widest text-emerald-500 flex items-center gap-2">
                  <div className="w-1 h-1 bg-emerald-500 rounded-full" /> Formato de Arquivo Sugerido
                </h3>
                <p className="text-sm text-gray-400 font-light leading-relaxed">
                  Para máxima precisão, recomendamos arquivos <span className="text-white font-medium">.XLSX</span> ou <span className="text-white font-medium">.CSV</span> com as seguintes colunas:
                </p>
                <div className="grid grid-cols-2 gap-2">
                  {[
                    'DATA', 'HORA', 'NOME_BENEFICIARIO', 'MATRICULA_BENEFICIARIO', 
                    'NOME_PRESTADOR', 'MATRICULA_PRESTADOR', 'ESPECIALIDADE',
                    'PROCEDIMENTO', 'QUANTIDADE', 'SITUACAO', 'CONSELHO', 'TIPO'
                  ].map(col => (
                    <div key={col} className="flex items-center gap-2 text-[10px] font-mono text-gray-500 bg-white/5 px-2 py-1 rounded">
                      <CheckCircle2 size={10} className="text-emerald-500/50" /> {col}
                    </div>
                  ))}
                </div>
              </div>
              <div className="space-y-4">
                <h3 className="text-xs font-bold uppercase tracking-widest text-amber-500 flex items-center gap-2">
                  <div className="w-1 h-1 bg-amber-500 rounded-full" /> Dicas de Processamento
                </h3>
                <ul className="space-y-3 text-sm text-gray-400 font-light">
                  <li className="flex items-start gap-2">
                    <span className="text-amber-500/50 mt-1">●</span>
                    <span>No modo Lote Consolidado, o sistema identifica cobranças idênticas faturadas em meses ou arquivos distintos.</span>
                  </li>
                  <li className="flex items-start gap-2">
                    <span className="text-amber-500/50 mt-1">●</span>
                    <span>Diferenciamos automaticamente Matrículas de Beneficiários e Prestadores.</span>
                  </li>
                  <li className="flex items-start gap-2">
                    <span className="text-amber-500/50 mt-1">●</span>
                    <span>O sistema possui algoritmos de aliasing para tentar identificar colunas com nomes variados, mas o padrão acima é o ideal.</span>
                  </li>
                  <li className="flex items-start gap-2">
                    <span className="text-amber-500/50 mt-1">●</span>
                    <span>Campos de Matrícula e DV podem estar separados ou unificados.</span>
                  </li>
                </ul>
              </div>
            </div>

            <div className="mt-12 text-center text-[10px] text-gray-600 uppercase tracking-widest flex gap-8">
              <span>● Processamento Local</span>
              <span>● Analisador Heurístico</span>
              <span>● Exportação Direta</span>
            </div>
          </motion.div>
        )}

        {isProcessing && (
          <div className="flex-1 flex flex-col items-center justify-center py-20">
            <div className="relative">
              <div className="w-14 h-14 border-2 border-white/10 border-t-emerald-400 rounded-full animate-spin"></div>
              <Clock className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 text-white/30" size={18} />
            </div>
            <p className="mt-5 text-xs uppercase font-mono tracking-[0.25em] text-emerald-400 animate-pulse text-center max-w-lg px-4">
              {processingStatus || 'Sincronizando e Processando Dados'}
            </p>
            <p className="mt-1 text-[10px] text-gray-500 font-mono">Por favor, aguarde enquanto os registros são auditados...</p>
          </div>
        )}

        {error && (
          <div className="p-4 bg-red-900/10 border border-red-500/20 rounded text-red-400 text-xs font-mono mb-6 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <AlertCircle size={14} />
              <span>{error}</span>
            </div>
            <button onClick={() => setError(null)} className="hover:text-red-300">
              <X size={14} />
            </button>
          </div>
        )}

        {data && !isProcessing && (
          <div className="flex flex-col gap-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
            {deduplicationMessage && (
              <div className="bg-emerald-500/10 border border-emerald-500/30 rounded-lg px-4 py-2.5 text-xs text-emerald-300 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                <div className="flex items-center gap-2">
                  <CheckCircle2 size={16} className="text-emerald-400 shrink-0" />
                  <span>{deduplicationMessage}</span>
                </div>
                <span className="text-[9px] uppercase font-mono text-emerald-400 bg-emerald-500/20 border border-emerald-500/30 px-2 py-0.5 rounded self-start sm:self-auto font-semibold">
                  Remessa Priorizada
                </span>
              </div>
            )}

            {/* Visual Quantification & Tab Switcher Bar */}
            <div className="flex flex-col md:flex-row justify-between items-stretch md:items-center gap-4 bg-white/[0.02] border border-white/10 p-4 rounded-lg">
              <div className="flex items-center gap-3 overflow-x-auto">
                <div className="px-3 py-2 bg-white/5 border border-white/10 rounded flex items-center gap-2">
                  <span className="text-[10px] text-gray-400 uppercase">Total:</span>
                  <span className="text-sm font-mono font-bold text-white">{categoryStats.total}</span>
                </div>
                <div className="px-3 py-2 bg-purple-950/40 border border-purple-500/40 rounded flex items-center gap-2" title="Transações registradas exatamente no mesmo horário (0 min)">
                  <span className="w-2 h-2 rounded-full bg-purple-400 animate-pulse"></span>
                  <span className="text-[10px] text-purple-300 uppercase font-semibold">Mesmo Horário (Muito Suspeito):</span>
                  <span className="text-sm font-mono font-bold text-purple-300">{categoryStats.veryHigh}</span>
                </div>
                <div className="px-3 py-2 bg-red-950/40 border border-red-500/30 rounded flex items-center gap-2" title="Intervalo até 30 minutos">
                  <span className="w-2 h-2 rounded-full bg-red-500"></span>
                  <span className="text-[10px] text-red-300 uppercase">≤ 30m (Provável):</span>
                  <span className="text-sm font-mono font-bold text-red-400">{categoryStats.high}</span>
                </div>
                <div className="px-3 py-2 bg-amber-950/30 border border-amber-500/30 rounded flex items-center gap-2" title="Intervalo entre 30 minutos e 1 hora">
                  <span className="w-2 h-2 rounded-full bg-amber-500"></span>
                  <span className="text-[10px] text-amber-300 uppercase">30m-1h (Menos):</span>
                  <span className="text-sm font-mono font-bold text-amber-400">{categoryStats.medium}</span>
                </div>
                <div className="px-3 py-2 bg-white/5 border border-white/10 rounded flex items-center gap-2" title="Intervalo superior a 1 hora">
                  <span className="w-2 h-2 rounded-full bg-gray-400"></span>
                  <span className="text-[10px] text-gray-400 uppercase">&gt; 1h:</span>
                  <span className="text-sm font-mono font-bold text-gray-300">{categoryStats.low}</span>
                </div>
              </div>

              <div className="flex items-center gap-2 self-end md:self-auto">
                <button
                  onClick={() => setActiveTab('table')}
                  className={cn(
                    "px-4 py-2 text-[10px] uppercase font-mono tracking-widest rounded border transition-all flex items-center gap-2 cursor-pointer",
                    activeTab === 'table' ? "bg-white/15 border-white/30 text-white font-bold" : "bg-white/5 border-white/10 text-gray-400 hover:text-white"
                  )}
                >
                  <Search size={14} /> Listagem ({filteredDuplicates.length})
                </button>
                <button
                  onClick={() => setActiveTab('dashboard')}
                  className={cn(
                    "px-4 py-2 text-[10px] uppercase font-mono tracking-widest rounded border transition-all flex items-center gap-2 cursor-pointer",
                    activeTab !== 'table' ? "bg-emerald-500/20 border-emerald-400/50 text-emerald-300 font-bold shadow-lg shadow-emerald-950/40" : "bg-white/5 border-white/10 text-gray-400 hover:text-white"
                  )}
                >
                  <BarChart3 size={14} /> Dashboard & Comparativos
                </button>
              </div>
            </div>

            {activeTab !== 'table' ? (
              <div className="flex flex-col gap-6 animate-in fade-in duration-300">
                {/* Dashboard Header & Export */}
                <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4 bg-white/[0.02] border border-white/10 px-6 py-5 rounded-lg">
                  <div>
                    <div className="flex items-center gap-2 mb-1">
                      <span className="px-2 py-0.5 rounded text-[9px] uppercase font-mono tracking-widest bg-emerald-500/20 text-emerald-400 border border-emerald-500/30 font-bold">
                        Painel Consolidado
                      </span>
                      <span className="text-[10px] text-gray-400 font-mono">
                        Volumetria: {filteredDuplicates.length} grupos ({isFiltered ? filteredRowCount : transactionCount} registros)
                      </span>
                    </div>
                    <h3 className="text-base font-bold uppercase tracking-widest text-white flex items-center gap-2">
                      <BarChart3 size={18} className="text-emerald-400" /> Dashboard Executivo de Duplicidades, Picos e Padrões Comparativos
                    </h3>
                    <p className="text-xs text-gray-400 mt-1 max-w-3xl leading-relaxed">
                      Análise integrada de incidência temporal, comparativos entre dias da semana, evolução mensal, turnos do dia, prestadores, procedimentos e beneficiários recorrentes.
                    </p>
                  </div>
                  <div className="flex items-center gap-3">
                    <button 
                      onClick={handleExportFormattedDashboard}
                      className="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white text-[10px] uppercase font-bold tracking-widest rounded transition-all flex items-center gap-2 shadow-lg shadow-emerald-900/30 cursor-pointer"
                      title="Gera dashboard executivo completo em Excel (.XLSX) formatado com cores, gráficos e planilhas comparativas"
                    >
                      <FileDown size={14} /> Exportar Dashboard & Padrões (.XLSX)
                    </button>
                    <button 
                      onClick={exportDashboardReport}
                      className="px-3 py-2.5 bg-white/5 border border-white/10 hover:bg-white/10 text-gray-300 hover:text-white text-[10px] uppercase tracking-widest rounded transition-all flex items-center gap-2 cursor-pointer"
                      title="Exportação simples do resumo analítico"
                    >
                      <FileDown size={12} /> Dados Simples
                    </button>
                  </div>
                </div>

                {/* 4 Indicadores Executivos Rápidos (KPIs) */}
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                  <div className="bg-white/[0.02] border border-white/10 p-5 rounded-lg flex flex-col justify-between hover:border-emerald-500/40 transition-all">
                    <div>
                      <span className="text-[10px] uppercase tracking-widest text-gray-400 font-mono block mb-1">Total de Grupos Duplicados</span>
                      <div className="flex items-baseline gap-2">
                        <span className="text-2xl font-bold font-mono text-white">{itStats.totalGroups}</span>
                        <span className="text-xs text-emerald-400 font-mono font-semibold">({isFiltered ? filteredRowCount : transactionCount} reg.)</span>
                      </div>
                    </div>
                    <p className="text-[10px] text-gray-400 mt-3 pt-3 border-t border-white/5">
                      Volume total de grupos em duplicidade analisados.
                    </p>
                  </div>

                  <div className="bg-white/[0.02] border border-white/10 p-5 rounded-lg flex flex-col justify-between hover:border-red-500/40 transition-all">
                    <div>
                      <span className="text-[10px] uppercase tracking-widest text-gray-400 font-mono block mb-1">Simultâneas (0 min / Provável)</span>
                      <div className="flex items-baseline gap-2">
                        <span className="text-2xl font-bold font-mono text-red-400">{itStats.deltaZeroCount}</span>
                        <span className="text-xs text-red-400 font-mono font-semibold">({itStats.instantPercent}%)</span>
                      </div>
                    </div>
                    <p className="text-[10px] text-gray-400 mt-3 pt-3 border-t border-white/5">
                      Ocorrências gravadas no mesmo minuto exato.
                    </p>
                  </div>

                  <div className="bg-white/[0.02] border border-white/10 p-5 rounded-lg flex flex-col justify-between hover:border-amber-500/40 transition-all">
                    <div>
                      <span className="text-[10px] uppercase tracking-widest text-gray-400 font-mono block mb-1">Micro-intervalo (1 a 5 min)</span>
                      <div className="flex items-baseline gap-2">
                        <span className="text-2xl font-bold font-mono text-amber-400">{itStats.delta1to5Count}</span>
                        <span className="text-xs text-amber-400 font-mono font-semibold">({itStats.microPercent}%)</span>
                      </div>
                    </div>
                    <p className="text-[10px] text-gray-400 mt-3 pt-3 border-t border-white/5">
                      Duplicidades ocorridas em até 5 minutos.
                    </p>
                  </div>

                  <div className="bg-white/[0.02] border border-white/10 p-5 rounded-lg flex flex-col justify-between hover:border-purple-500/40 transition-all">
                    <div>
                      <span className="text-[10px] uppercase tracking-widest text-gray-400 font-mono block mb-1">Concentração Prestador Líder</span>
                      <div className="flex items-baseline gap-2">
                        <span className="text-2xl font-bold font-mono text-purple-400">{itStats.topProviderConcentrationPercent}%</span>
                        <span className="text-xs text-gray-400 font-mono truncate max-w-[120px]" title={itStats.topProviderName}>{itStats.topProviderName}</span>
                      </div>
                    </div>
                    <p className="text-[10px] text-gray-400 mt-3 pt-3 border-t border-white/5">
                      Top 3 prestadores somam {itStats.top3ProvidersSharePercent}% das ocorrências.
                    </p>
                  </div>
                </div>

                {/* BLOCO 1: COMPARATIVOS TEMPORAIS (DIAS DA SEMANA E DIAS DO MÊS) */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  {/* Painel 1.1: Comparativo por Dia da Semana */}
                  <div className="bg-white/[0.02] border border-white/10 p-6 rounded-lg flex flex-col gap-4">
                    <div className="flex justify-between items-center border-b border-white/5 pb-3">
                      <h4 className="text-xs font-bold uppercase tracking-widest text-sky-300 flex items-center gap-2">
                        <Calendar size={15} /> Comparativo por Dia da Semana (Seg a Dom)
                      </h4>
                      <span className="text-[10px] font-mono text-gray-400">Distribuição Semanal</span>
                    </div>

                    <div className="space-y-3">
                      {itStats.dayOfWeekDistribution.map((item, idx) => (
                        <div key={idx} className="bg-white/5 p-3 rounded border border-white/5 flex flex-col gap-1.5">
                          <div className="flex justify-between items-center text-xs">
                            <span className="text-white font-medium">{item.dayName}</span>
                            <span className="font-mono font-bold text-sky-300">
                              {item.count} grupo(s) ({item.percent}%)
                            </span>
                          </div>
                          <div className="w-full bg-white/10 h-2 rounded-full overflow-hidden">
                            <div 
                              className="bg-sky-500 h-full rounded-full transition-all duration-500" 
                              style={{ width: `${Math.min(100, Math.max(2, item.percent))}%` }}
                            />
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Painel 1.2: Dias do Mês com Maior Incidência (Datas de Pico) */}
                  <div className="bg-white/[0.02] border border-white/10 p-6 rounded-lg flex flex-col gap-4">
                    <div className="flex justify-between items-center border-b border-white/5 pb-3">
                      <h4 className="text-xs font-bold uppercase tracking-widest text-emerald-400 flex items-center gap-2">
                        <Calendar size={15} /> Datas de Pico no Mês (Maiores Incidências)
                      </h4>
                      <span className="text-[10px] font-mono text-gray-400">Picos por Data</span>
                    </div>

                    <div className="space-y-3 max-h-[340px] overflow-y-auto custom-scrollbar pr-2">
                      {dashboardStats.topDays.length === 0 ? (
                        <p className="text-xs text-gray-500 italic">Nenhum dado de data disponível.</p>
                      ) : (
                        dashboardStats.topDays.map((item) => {
                          const maxCount = dashboardStats.topDays[0]?.count || 1;
                          const pct = Math.round((item.count / maxCount) * 100);
                          return (
                            <div key={item.day} className="flex flex-col gap-1.5 bg-white/5 p-3 rounded border border-white/5">
                              <div className="flex justify-between items-center text-xs">
                                <span className="font-mono text-white font-medium">{item.day || 'Data Indefinida'}</span>
                                <span className="font-mono font-bold text-emerald-400">{item.count} grupo(s)</span>
                              </div>
                              <div className="w-full bg-white/10 h-2 rounded-full overflow-hidden">
                                <div className="bg-emerald-500 h-full rounded-full transition-all duration-500" style={{ width: `${pct}%` }} />
                              </div>
                            </div>
                          );
                        })
                      )}
                    </div>
                  </div>
                </div>

                {/* BLOCO 1.5: ANÁLISE COMPARATIVA DOS DIAS DO MÊS (01 A 31) E CICLO MENSAL */}
                <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                  {/* Painel: Comparativo por Dias do Mês (Dias 01 a 31) */}
                  <div className="lg:col-span-2 bg-white/[0.02] border border-white/10 p-6 rounded-lg flex flex-col gap-4">
                    <div className="flex justify-between items-center border-b border-white/5 pb-3">
                      <h4 className="text-xs font-bold uppercase tracking-widest text-emerald-300 flex items-center gap-2">
                        <Calendar size={15} /> Comparativo por Dias do Mês (Dias 01 a 31)
                      </h4>
                      <span className="text-[10px] font-mono text-gray-400">Padrão Diário</span>
                    </div>

                    <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6 gap-2.5 max-h-[320px] overflow-y-auto custom-scrollbar pr-2">
                      {itStats.dayOfMonthDistribution.map((dom) => (
                        <div key={dom.dayNum} className="bg-white/5 p-2.5 rounded border border-white/5 flex flex-col justify-between gap-1 hover:border-emerald-500/30 transition-all">
                          <div className="flex justify-between items-center">
                            <span className="text-[11px] font-mono font-bold text-white">{dom.dayLabel}</span>
                            <span className="text-[10px] font-mono font-bold text-emerald-400">{dom.count}</span>
                          </div>
                          <div className="w-full bg-white/10 h-1.5 rounded-full overflow-hidden mt-1">
                            <div 
                              className="bg-emerald-500 h-full rounded-full transition-all duration-500" 
                              style={{ width: `${Math.min(100, Math.max(2, dom.percent))}%` }}
                            />
                          </div>
                          <span className="text-[9px] font-mono text-gray-400 text-right mt-0.5">{dom.percent}%</span>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Painel: Etapas do Ciclo Mensal de Faturamento */}
                  <div className="bg-white/[0.02] border border-white/10 p-6 rounded-lg flex flex-col gap-4">
                    <div className="flex justify-between items-center border-b border-white/5 pb-3">
                      <h4 className="text-xs font-bold uppercase tracking-widest text-amber-300 flex items-center gap-2">
                        <TrendingUp size={15} /> Etapas do Ciclo Mensal
                      </h4>
                      <span className="text-[10px] font-mono text-gray-400">Início, Meio e Fim</span>
                    </div>

                    <div className="space-y-3">
                      {itStats.monthCycleDistribution.map((cy, idx) => (
                        <div key={idx} className="bg-white/5 p-3 rounded border border-white/5 flex flex-col gap-1.5">
                          <div className="flex justify-between items-center text-xs">
                            <span className="text-white font-medium">{cy.cycleName}</span>
                            <span className="font-mono font-bold text-amber-300">
                              {cy.count} grupo(s) ({cy.percent}%)
                            </span>
                          </div>
                          <div className="w-full bg-white/10 h-2 rounded-full overflow-hidden">
                            <div 
                              className="bg-amber-500 h-full rounded-full transition-all duration-500" 
                              style={{ width: `${Math.min(100, Math.max(2, cy.percent))}%` }}
                            />
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>

                {/* BLOCO 2: ANÁLISE POR HORÁRIO E TURNOS DO DIA */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  {/* Painel 2.1: Distribuição por Turnos do Dia (Soma 100%) */}
                  <div className="bg-white/[0.02] border border-white/10 p-6 rounded-lg flex flex-col gap-4">
                    <div className="flex justify-between items-center border-b border-white/5 pb-3">
                      <h4 className="text-xs font-bold uppercase tracking-widest text-amber-300 flex items-center gap-2">
                        <Clock size={15} /> Distribuição por Turnos do Dia
                      </h4>
                      <span className="text-[10px] font-mono text-gray-400">Total: 100%</span>
                    </div>

                    <div className="space-y-3">
                      {itStats.shiftDistribution.map((item, idx) => (
                        <div key={idx} className="bg-white/5 p-3 rounded border border-white/5 flex flex-col gap-1.5">
                          <div className="flex justify-between items-center text-xs">
                            <span className="text-white font-medium">{item.shiftName}</span>
                            <span className="font-mono font-bold text-amber-300">
                              {item.count} ocorrência(s) ({item.percent}%)
                            </span>
                          </div>
                          <div className="w-full bg-white/10 h-2 rounded-full overflow-hidden">
                            <div 
                              className="bg-amber-500 h-full rounded-full transition-all duration-500" 
                              style={{ width: `${Math.min(100, Math.max(2, item.percent))}%` }}
                            />
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Painel 2.2: Horários de Pico (Faixas Horárias Específicas) */}
                  <div className="bg-white/[0.02] border border-white/10 p-6 rounded-lg flex flex-col gap-4">
                    <div className="flex justify-between items-center border-b border-white/5 pb-3">
                      <h4 className="text-xs font-bold uppercase tracking-widest text-amber-400 flex items-center gap-2">
                        <Clock size={15} /> Faixas Horárias de Pico (HH:00)
                      </h4>
                      <span className="text-[10px] font-mono text-gray-400">Horários Mais Críticos</span>
                    </div>

                    <div className="space-y-3 max-h-[340px] overflow-y-auto custom-scrollbar pr-2">
                      {dashboardStats.topHours.length === 0 ? (
                        <p className="text-xs text-gray-500 italic">Nenhum horário registrado.</p>
                      ) : (
                        dashboardStats.topHours.map((item) => {
                          const totalH = dashboardStats.topHours.reduce((acc, h) => acc + h.count, 0) || 1;
                          const pct = Math.round((item.count / totalH) * 100);
                          return (
                            <div key={item.hour} className="flex flex-col gap-1.5 bg-white/5 p-3 rounded border border-white/5">
                              <div className="flex justify-between items-center text-xs">
                                <span className="font-mono text-white font-medium">{item.hour}</span>
                                <span className="font-mono font-bold text-amber-400">{item.count} ocorrência(s) ({pct}%)</span>
                              </div>
                              <div className="w-full bg-white/10 h-2 rounded-full overflow-hidden">
                                <div className="bg-amber-500 h-full rounded-full transition-all duration-500" style={{ width: `${Math.min(100, Math.max(2, pct))}%` }} />
                              </div>
                            </div>
                          );
                        })
                      )}
                    </div>
                  </div>
                </div>

                {/* BLOCO 3: EVOLUÇÃO MENSAL E FAIXAS DE INTERVALO TEMPORAL */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  {/* Painel 3.1: Evolução Temporal entre Meses */}
                  <div className="bg-white/[0.02] border border-white/10 p-6 rounded-lg flex flex-col gap-4">
                    <div className="flex justify-between items-center border-b border-white/5 pb-3">
                      <h4 className="text-xs font-bold uppercase tracking-widest text-blue-300 flex items-center gap-2">
                        <TrendingUp size={15} /> Evolução Temporal entre Meses
                      </h4>
                      <span className="text-[10px] font-mono text-gray-400">Comparativo Mensal</span>
                    </div>

                    <div className="space-y-2.5">
                      {itStats.monthlyDistribution.length === 0 ? (
                        <p className="text-xs text-gray-500 italic">Nenhuma data mensal identificada.</p>
                      ) : (
                        itStats.monthlyDistribution.map((m, idx) => (
                          <div key={idx} className="bg-white/5 p-2.5 rounded border border-white/5 flex flex-col gap-1.5">
                            <div className="flex justify-between items-center text-xs">
                              <span className="text-white font-medium font-mono">{m.monthLabel}</span>
                              <span className="font-mono font-bold text-blue-300 shrink-0 ml-2">
                                {m.count} caso(s) ({m.percent}%)
                              </span>
                            </div>
                            <div className="w-full bg-white/10 h-1.5 rounded-full overflow-hidden">
                              <div 
                                className="bg-blue-500 h-full rounded-full transition-all duration-500" 
                                style={{ width: `${Math.min(100, Math.max(2, m.percent))}%` }}
                              />
                            </div>
                          </div>
                        ))
                      )}
                    </div>
                  </div>

                  {/* Painel 3.2: Frequência por Intervalo Temporal (Δt) */}
                  <div className="bg-white/[0.02] border border-white/10 p-6 rounded-lg flex flex-col gap-4">
                    <div className="flex justify-between items-center border-b border-white/5 pb-3">
                      <h4 className="text-xs font-bold uppercase tracking-widest text-sky-300 flex items-center gap-2">
                        <Clock size={15} /> Frequência por Intervalo Temporal (Δt)
                      </h4>
                      <span className="text-[10px] font-mono text-gray-400">Janela de Tempo</span>
                    </div>

                    <div className="space-y-3">
                      <div className="bg-white/5 p-3 rounded border border-white/5">
                        <div className="flex justify-between items-center text-xs">
                          <span className="font-mono text-red-300 font-bold">Δt = 0 min (Mesmo Minuto / Simultâneo)</span>
                          <span className="font-mono font-bold text-white">{itStats.deltaZeroCount} casos ({itStats.instantPercent}%)</span>
                        </div>
                        <div className="w-full bg-white/10 h-2 rounded-full mt-2 overflow-hidden">
                          <div className="bg-red-500 h-full rounded-full transition-all duration-500" style={{ width: `${itStats.instantPercent}%` }} />
                        </div>
                      </div>

                      <div className="bg-white/5 p-3 rounded border border-white/5">
                        <div className="flex justify-between items-center text-xs">
                          <span className="font-mono text-amber-300 font-bold">Δt = 1 a 5 min (Micro-intervalo)</span>
                          <span className="font-mono font-bold text-white">{itStats.delta1to5Count} casos ({itStats.microPercent}%)</span>
                        </div>
                        <div className="w-full bg-white/10 h-2 rounded-full mt-2 overflow-hidden">
                          <div className="bg-amber-500 h-full rounded-full transition-all duration-500" style={{ width: `${itStats.microPercent}%` }} />
                        </div>
                      </div>

                      <div className="bg-white/5 p-3 rounded border border-white/5">
                        <div className="flex justify-between items-center text-xs">
                          <span className="font-mono text-blue-300 font-bold">Δt = 6 a 30 min (Mesmo Turno)</span>
                          <span className="font-mono font-bold text-white">{itStats.delta6to30Count} casos ({itStats.shortPercent}%)</span>
                        </div>
                        <div className="w-full bg-white/10 h-2 rounded-full mt-2 overflow-hidden">
                          <div className="bg-blue-500 h-full rounded-full transition-all duration-500" style={{ width: `${itStats.shortPercent}%` }} />
                        </div>
                      </div>

                      <div className="bg-white/5 p-3 rounded border border-white/5">
                        <div className="flex justify-between items-center text-xs">
                          <span className="font-mono text-gray-300 font-bold">Δt &gt; 30 min (Reemissão Tardia)</span>
                          <span className="font-mono font-bold text-white">{itStats.deltaOver30Count} casos ({itStats.extendedPercent}%)</span>
                        </div>
                        <div className="w-full bg-white/10 h-2 rounded-full mt-2 overflow-hidden">
                          <div className="bg-gray-400 h-full rounded-full transition-all duration-500" style={{ width: `${itStats.extendedPercent}%` }} />
                        </div>
                      </div>
                    </div>
                  </div>
                </div>

                {/* BLOCO 4: RANKINGS E CRUZAMENTO DE PRESTADORES, PROCEDIMENTOS, BENEFICIÁRIOS E SITUAÇÃO */}
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                  {/* Painel 4.1: Prestadores com Mais Duplicidades (Top 10) */}
                  <div className="bg-white/[0.02] border border-white/10 p-6 rounded-lg flex flex-col gap-4">
                    <div className="flex justify-between items-center border-b border-white/5 pb-3">
                      <h4 className="text-xs font-bold uppercase tracking-widest text-purple-300 flex items-center gap-2">
                        <Building2 size={15} /> Prestadores com Mais Duplicidades (Top 10)
                      </h4>
                      <span className="text-[10px] font-mono text-gray-400">Share %</span>
                    </div>

                    <div className="space-y-2.5">
                      {itStats.topProvidersList.map((p, idx) => (
                        <div key={idx} className="bg-white/5 p-2.5 rounded border border-white/5 flex flex-col gap-1.5">
                          <div className="flex justify-between items-center text-xs">
                            <span className="text-white font-medium truncate max-w-[280px]" title={p.provider}>
                              {idx + 1}. {p.provider}
                            </span>
                            <span className="font-mono font-bold text-purple-300 shrink-0 ml-2">
                              {p.count} casos ({p.percent}%)
                            </span>
                          </div>
                          <div className="w-full bg-white/10 h-1.5 rounded-full overflow-hidden">
                            <div 
                              className="bg-purple-500 h-full rounded-full transition-all duration-500" 
                              style={{ width: `${Math.min(100, Math.max(2, p.percent))}%` }}
                            />
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Painel 4.2: Procedimentos Mais Frequentes em Duplicidade (Top 10) */}
                  <div className="bg-white/[0.02] border border-white/10 p-6 rounded-lg flex flex-col gap-4">
                    <div className="flex justify-between items-center border-b border-white/5 pb-3">
                      <h4 className="text-xs font-bold uppercase tracking-widest text-indigo-300 flex items-center gap-2">
                        <FileText size={15} /> Procedimentos Mais Frequentes em Duplicidade (Top 10)
                      </h4>
                      <span className="text-[10px] font-mono text-gray-400">Volume</span>
                    </div>

                    <div className="space-y-2.5">
                      {itStats.topProceduresList.map((pr, idx) => (
                        <div key={idx} className="bg-white/5 p-2.5 rounded border border-white/5 flex flex-col gap-1.5">
                          <div className="flex justify-between items-center text-xs">
                            <span className="text-white font-medium truncate max-w-[280px]" title={pr.procedure}>
                              {idx + 1}. {pr.procedure}
                            </span>
                            <span className="font-mono font-bold text-indigo-300 shrink-0 ml-2">
                              {pr.count} casos ({pr.percent}%)
                            </span>
                          </div>
                          <div className="w-full bg-white/10 h-1.5 rounded-full overflow-hidden">
                            <div 
                              className="bg-indigo-500 h-full rounded-full transition-all duration-500" 
                              style={{ width: `${Math.min(100, Math.max(2, pr.percent))}%` }}
                            />
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Painel 4.3: Beneficiários / Pacientes com Maior Recorrência (Top 10) */}
                  <div className="bg-white/[0.02] border border-white/10 p-6 rounded-lg flex flex-col gap-4">
                    <div className="flex justify-between items-center border-b border-white/5 pb-3">
                      <h4 className="text-xs font-bold uppercase tracking-widest text-emerald-300 flex items-center gap-2">
                        <User size={15} /> Beneficiários / Pacientes com Maior Recorrência (Top 10)
                      </h4>
                      <span className="text-[10px] font-mono text-gray-400">Recorrência</span>
                    </div>

                    <div className="space-y-2.5">
                      {itStats.topBeneficiariesList.length === 0 ? (
                        <p className="text-xs text-gray-500 italic">Nenhum beneficiário identificado.</p>
                      ) : (
                        itStats.topBeneficiariesList.map((b, idx) => (
                          <div key={idx} className="bg-white/5 p-2.5 rounded border border-white/5 flex flex-col gap-1.5">
                            <div className="flex justify-between items-center text-xs">
                              <span className="text-white font-medium truncate max-w-[280px]" title={b.beneficiary}>
                                {idx + 1}. {b.beneficiary}
                              </span>
                              <span className="font-mono font-bold text-emerald-400 shrink-0 ml-2">
                                {b.count} caso(s) ({b.percent}%)
                              </span>
                            </div>
                            <div className="w-full bg-white/10 h-1.5 rounded-full overflow-hidden">
                              <div 
                                className="bg-emerald-500 h-full rounded-full transition-all duration-500" 
                                style={{ width: `${Math.min(100, Math.max(2, b.percent))}%` }}
                              />
                            </div>
                          </div>
                        ))
                      )}
                    </div>
                  </div>

                  {/* Painel 4.4: Distribuição por Situação / Status da Transação */}
                  <div className="bg-white/[0.02] border border-white/10 p-6 rounded-lg flex flex-col gap-4">
                    <div className="flex justify-between items-center border-b border-white/5 pb-3">
                      <h4 className="text-xs font-bold uppercase tracking-widest text-purple-300 flex items-center gap-2">
                        <CheckCircle2 size={15} /> Distribuição por Situação / Status da Transação
                      </h4>
                      <span className="text-[10px] font-mono text-gray-400">Status</span>
                    </div>

                    <div className="space-y-2.5">
                      {itStats.situationsDistribution.length === 0 ? (
                        <p className="text-xs text-gray-500 italic">Nenhuma situação informada.</p>
                      ) : (
                        itStats.situationsDistribution.map((st, idx) => (
                          <div key={idx} className="bg-white/5 p-2.5 rounded border border-white/5 flex flex-col gap-1.5">
                            <div className="flex justify-between items-center text-xs">
                              <span className="text-white font-medium truncate max-w-[280px]" title={st.situation}>
                                {st.situation}
                              </span>
                              <span className="font-mono font-bold text-purple-300 shrink-0 ml-2">
                                {st.count} caso(s) ({st.percent}%)
                              </span>
                            </div>
                            <div className="w-full bg-white/10 h-1.5 rounded-full overflow-hidden">
                              <div 
                                className="bg-purple-500 h-full rounded-full transition-all duration-500" 
                                style={{ width: `${Math.min(100, Math.max(2, st.percent))}%` }}
                              />
                            </div>
                          </div>
                        ))
                      )}
                    </div>
                  </div>
                </div>

                {/* BLOCO 5: COMPORTAMENTO DE IDENTIFICADORES (PROTOCOLOS E REMESSAS) */}
                <div className="bg-white/[0.02] border border-white/10 p-6 rounded-lg flex flex-col gap-4">
                  <div className="flex justify-between items-center border-b border-white/5 pb-3">
                    <h4 className="text-xs font-bold uppercase tracking-widest text-emerald-300 flex items-center gap-2">
                      <Database size={15} /> Comportamento de Identificadores (Protocolos e Remessas)
                    </h4>
                    <span className="text-[10px] font-mono text-gray-400">Proporção Estatística</span>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {/* Mesmo Protocolo */}
                    <div className="bg-white/5 p-3.5 rounded border border-white/5">
                      <div className="flex justify-between items-center text-xs">
                        <span className="font-mono text-emerald-300 font-bold">Mesmo Número de Protocolo</span>
                        <span className="font-mono font-bold text-white">{itStats.sameProtocolCount} casos ({itStats.sameProtocolPercent}%)</span>
                      </div>
                      <div className="w-full bg-white/10 h-2 rounded-full mt-2.5 overflow-hidden">
                        <div 
                          className="bg-emerald-500 h-full rounded-full transition-all duration-500" 
                          style={{ width: `${itStats.sameProtocolPercent}%` }}
                        />
                      </div>
                    </div>

                    {/* Protocolos Distintos */}
                    <div className="bg-white/5 p-3.5 rounded border border-white/5">
                      <div className="flex justify-between items-center text-xs">
                        <span className="font-mono text-amber-300 font-bold">Protocolos Distintos Gerados</span>
                        <span className="font-mono font-bold text-white">{itStats.differentProtocolCount} casos ({itStats.differentProtocolPercent}%)</span>
                      </div>
                      <div className="w-full bg-white/10 h-2 rounded-full mt-2.5 overflow-hidden">
                        <div 
                          className="bg-amber-500 h-full rounded-full transition-all duration-500" 
                          style={{ width: `${itStats.differentProtocolPercent}%` }}
                        />
                      </div>
                    </div>

                    {/* Mesma Remessa */}
                    <div className="bg-white/5 p-3.5 rounded border border-white/5">
                      <div className="flex justify-between items-center text-xs">
                        <span className="font-mono text-purple-300 font-bold">Mesma Remessa de Faturamento</span>
                        <span className="font-mono font-bold text-white">{itStats.sameRemessaCount} casos ({itStats.sameRemessaPercent}%)</span>
                      </div>
                      <div className="w-full bg-white/10 h-2 rounded-full mt-2.5 overflow-hidden">
                        <div 
                          className="bg-purple-500 h-full rounded-full transition-all duration-500" 
                          style={{ width: `${itStats.sameRemessaPercent}%` }}
                        />
                      </div>
                    </div>

                    {/* Remessas Diferentes */}
                    <div className="bg-white/5 p-3.5 rounded border border-white/5">
                      <div className="flex justify-between items-center text-xs">
                        <span className="font-mono text-sky-300 font-bold">Remessas Diferentes / Lotes Distintos</span>
                        <span className="font-mono font-bold text-white">{itStats.differentRemessaCount} casos ({itStats.differentRemessaPercent}%)</span>
                      </div>
                      <div className="w-full bg-white/10 h-2 rounded-full mt-2.5 overflow-hidden">
                        <div 
                          className="bg-sky-500 h-full rounded-full transition-all duration-500" 
                          style={{ width: `${itStats.differentRemessaPercent}%` }}
                        />
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            ) : (
              <>
                {/* Filter Section */}
                <section className="bg-white/5 border border-white/10 p-6 rounded-lg flex flex-col gap-6">
                  <div className="flex items-center justify-between border-b border-white/5 pb-4">
                    <div className="flex items-center gap-3">
                      <Filter size={16} className="text-gray-500" />
                      <h3 className="text-xs uppercase font-bold tracking-widest text-gray-400">Filtros de Auditoria</h3>
                    </div>
                    {Object.values(filters).some(v => v !== '') && (
                      <button 
                        onClick={resetFilters}
                        className="text-[10px] uppercase text-gray-500 hover:text-white flex items-center gap-1 transition-colors"
                      >
                        <Trash2 size={12} /> Limpar Filtros
                      </button>
                    )}
                  </div>
                  
                  <div className="grid grid-cols-1 md:grid-cols-8 gap-4">
                    <div className="space-y-2">
                      <label className="text-[10px] uppercase text-gray-500 font-mono">Beneficiário</label>
                      <div className="relative">
                        <User className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-600" size={14} />
                        <input 
                          value={filters.beneficiario}
                          onChange={e => setFilters(prev => ({ ...prev, beneficiario: e.target.value }))}
                          className="w-full bg-[#0A0B0D] border border-white/10 rounded px-9 py-2 text-xs focus:border-white/30 outline-none transition-all"
                          placeholder="Nome..."
                        />
                      </div>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] uppercase text-gray-500 font-mono">Matrícula (Bene/Prest)</label>
                      <div className="relative">
                        <CheckCircle2 className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-600" size={14} />
                        <input 
                          value={filters.matricula}
                          onChange={e => setFilters(prev => ({ ...prev, matricula: e.target.value }))}
                          className="w-full bg-[#0A0B0D] border border-white/10 rounded px-9 py-2 text-xs focus:border-white/30 outline-none transition-all"
                          placeholder="Nº..."
                        />
                      </div>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] uppercase text-gray-500 font-mono">Prestador</label>
                      <select 
                        value={filters.prestador}
                        onChange={e => setFilters(prev => ({ ...prev, prestador: e.target.value }))}
                        className="w-full bg-[#0A0B0D] border border-white/10 rounded px-4 py-2 text-xs focus:border-white/30 outline-none transition-all cursor-pointer"
                      >
                        <option value="">Todos</option>
                        {uniqueValues.prestadores.map(p => (
                          <option key={p} value={p}>{p}</option>
                        ))}
                      </select>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] uppercase text-gray-500 font-mono">CRM Executante</label>
                      <select 
                        value={filters.conselhoExecutante}
                        onChange={e => setFilters(prev => ({ ...prev, conselhoExecutante: e.target.value }))}
                        className="w-full bg-[#0A0B0D] border border-white/10 rounded px-4 py-2 text-xs focus:border-white/30 outline-none transition-all cursor-pointer"
                      >
                        <option value="">Todos</option>
                        {uniqueValues.conselhoExecutantes.map(c => (
                          <option key={c} value={c}>{c}</option>
                        ))}
                      </select>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] uppercase text-gray-500 font-mono">Procedimento</label>
                      <select 
                        value={filters.procedimento}
                        onChange={e => setFilters(prev => ({ ...prev, procedimento: e.target.value }))}
                        className="w-full bg-[#0A0B0D] border border-white/10 rounded px-4 py-2 text-xs focus:border-white/30 outline-none transition-all cursor-pointer"
                      >
                        <option value="">Todos</option>
                        {uniqueValues.procedimentos.map(p => (
                          <option key={p} value={p}>{p}</option>
                        ))}
                      </select>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] uppercase text-gray-500 font-mono">Data</label>
                      <select 
                        value={filters.data}
                        onChange={e => setFilters(prev => ({ ...prev, data: e.target.value }))}
                        className="w-full bg-[#0A0B0D] border border-white/10 rounded px-4 py-2 text-xs focus:border-white/30 outline-none transition-all cursor-pointer"
                      >
                        <option value="">Todas</option>
                        {uniqueValues.datas.map(d => (
                          <option key={d} value={d}>{d}</option>
                        ))}
                      </select>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] uppercase text-gray-500 font-mono">Situação</label>
                      <select 
                        value={filters.situacao}
                        onChange={e => setFilters(prev => ({ ...prev, situacao: e.target.value }))}
                        className="w-full bg-[#0A0B0D] border border-white/10 rounded px-4 py-2 text-xs focus:border-white/30 outline-none transition-all cursor-pointer"
                      >
                        <option value="">Todas</option>
                        {uniqueValues.situacoes.map(s => (
                          <option key={s} value={s}>{s}</option>
                        ))}
                      </select>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] uppercase text-gray-500 font-mono">Intervalo (Probabilidade)</label>
                      <select 
                        value={filters.intervalCategory}
                        onChange={e => setFilters(prev => ({ ...prev, intervalCategory: e.target.value }))}
                        className="w-full bg-[#0A0B0D] border border-white/10 rounded px-4 py-2 text-xs focus:border-white/30 outline-none transition-all cursor-pointer"
                      >
                        <option value="">Todos</option>
                        <option value="very_high">Mesmo Horário (Muito Suspeito - 0m)</option>
                        <option value="high">≤ 30min (Bem Provável)</option>
                        <option value="medium">30min - 1h (Pouco Menos)</option>
                        <option value="low">&gt; 1h</option>
                      </select>
                    </div>
                  </div>
                </section>

            {/* Actions Bar */}
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 bg-white/[0.02] border border-white/10 px-6 py-3 rounded-lg">
              <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">
                {filteredDuplicates.length} grupos de duplicidades encontrados com os filtros atuais
              </span>
              <div className="flex items-center gap-3 flex-wrap">
                <button 
                   disabled={filteredDuplicates.length === 0}
                   onClick={handleExportFormatted}
                   className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 border border-emerald-400/30 rounded text-[10px] uppercase font-bold tracking-widest text-white transition-all flex items-center gap-2 shadow-lg shadow-emerald-900/30 disabled:opacity-20 disabled:cursor-not-allowed cursor-pointer"
                   title="Gera relatório formal com cores de criticidade, cabeçalho de notificação, espaço para justificativa do prestador e aba com resumo de contatos"
                >
                  <FileDown size={14} /> Exportar Notificação Formatada (com Cores)
                </button>
                <button 
                   disabled={filteredDuplicates.length === 0}
                   onClick={exportResults}
                   className="px-3 py-2 bg-white/5 border border-white/10 rounded text-[10px] uppercase tracking-widest text-gray-300 hover:text-white hover:bg-white/10 transition-all flex items-center gap-2 disabled:opacity-20 disabled:cursor-not-allowed cursor-pointer"
                   title="Exportação direta dos dados tabulares brutos"
                >
                  <FileDown size={12} /> Dados Brutos (XLSX)
                </button>
              </div>
            </div>

            {/* Results Table */}
            <div className="bg-white/[0.02] border border-white/10 rounded-lg overflow-hidden flex flex-col shadow-2xl">
              <div className="grid grid-cols-12 bg-white/5 py-3 px-4 text-[10px] uppercase font-bold tracking-widest text-gray-400 border-b border-white/10">
                <div className="col-span-2">Beneficiário</div>
                <div className="col-span-2">Prestador</div>
                <div className="col-span-1">Especialidade</div>
                <div className="col-span-2">Procedimento</div>
                <div className="col-span-1 text-center">Tipo</div>
                <div className="col-span-1 text-center">Data</div>
                <div className="col-span-1 text-center">Qtd</div>
                <div className="col-span-1 text-center">Situação</div>
                <div className="col-span-1 text-center text-red-400 bg-red-400/5 rounded">Horários</div>
              </div>

              <div className="divide-y divide-white/5 max-h-[600px] overflow-y-auto custom-scrollbar">
                {filteredDuplicates.length === 0 ? (
                  <div className="py-24 flex flex-col items-center justify-center opacity-40">
                    <Search size={32} className="mb-4" />
                    <p className="text-sm font-light uppercase tracking-[0.2em]">Nenhum registro corresponde aos filtros</p>
                  </div>
                ) : (
                  filteredDuplicates.map((group, idx) => (
                    <div 
                      key={group.fingerprint}
                      className="grid grid-cols-12 items-center py-4 px-4 hover:bg-white/[0.03] transition-colors"
                    >
                      <div className="col-span-2 pr-4 border-r border-white/5">
                        <div className="flex flex-col gap-1 overflow-hidden">
                          <p className="text-xs text-white font-medium truncate uppercase" title={group.params.beneficiario}>{group.params.beneficiario}</p>
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <span className="text-[8px] text-emerald-500/60 font-mono self-start bg-emerald-500/5 px-1 rounded border border-emerald-500/10" title="Matrícula Beneficiário">
                              {group.params.beneficiario_matricula || '---'}
                            </span>
                            {sourceFiles.length > 1 && (
                              (() => {
                                const originSet = new Set(group.rows.map(r => r['ARQUIVO_ORIGEM']).filter(Boolean));
                                const isCross = originSet.size > 1;
                                return (
                                  <span 
                                    className={cn(
                                      "text-[7.5px] font-mono self-start px-1 py-0.5 rounded border inline-flex items-center gap-1",
                                      isCross 
                                        ? "bg-purple-500/20 border-purple-500/40 text-purple-300 font-bold"
                                        : "bg-cyan-500/10 border-cyan-500/20 text-cyan-400"
                                    )}
                                    title={isCross ? `Duplicidade identificada entre ${originSet.size} arquivos distintos` : 'Duplicidade dentro do mesmo arquivo'}
                                  >
                                    <Layers size={8} /> {isCross ? `Cruzado (${originSet.size} arq.)` : 'Mesmo arquivo'}
                                  </span>
                                );
                              })()
                            )}
                          </div>
                        </div>
                      </div>
                      <div className="col-span-2 px-4 border-r border-white/5">
                        <div className="flex flex-col gap-1 overflow-hidden">
                          <p className="text-xs text-gray-300 truncate uppercase" title={group.params.prestador}>{group.params.prestador}</p>
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <span className="text-[8px] text-blue-500/60 font-mono self-start bg-blue-500/5 px-1 rounded border border-blue-500/10" title="Matrícula Prestador">
                              {group.params.prestador_matricula || '---'}
                            </span>
                            {group.params.conselhoExecNum && (
                              <span className="text-[8px] text-purple-400/80 font-mono self-start bg-purple-500/5 px-1 rounded border border-purple-500/10" title="CRM Executante">
                                {group.params.conselhoExecTipo || 'CRM'} {group.params.conselhoExecNum}{group.params.conselhoExecUf ? `/${group.params.conselhoExecUf}` : ''}
                              </span>
                            )}
                          </div>
                        </div>
                      </div>
                      <div className="col-span-1 text-[9px] text-gray-400 text-center truncate italic px-1" title={group.params.especialidade}>
                        {group.params.especialidade || ''}
                      </div>
                      <div className="col-span-2 text-xs truncate font-mono text-white/90 px-2" title={group.params.procedimento}>
                        {group.params.procedimento}
                      </div>
                      <div className="col-span-1 text-[9px] text-blue-400 text-center truncate uppercase font-mono px-1" title={group.params.tipo}>
                        {group.params.tipo || ''}
                      </div>
                      <div className="col-span-1 text-xs text-gray-400 text-center">
                        {group.params.data}
                      </div>
                      <div className="col-span-1 text-xs font-mono text-white/80 text-center font-bold">
                        {group.params.quantidade.padStart(2, '0')}
                      </div>
                      <div className="col-span-1">
                        <span className={cn(
                          "text-[9px] uppercase px-2 py-0.5 rounded-full inline-block font-bold",
                          group.params.situacao.toLowerCase() === 'aprovada' 
                            ? "bg-emerald-500/10 text-emerald-500" 
                            : "bg-white/10 text-gray-400"
                        )}>
                          {group.params.situacao}
                        </span>
                      </div>
                      <div className="col-span-1 flex flex-col items-center gap-1.5">
                        {(() => {
                          const intervalInfo = getGroupTimeInterval(group.rows);
                          return (
                            <>
                              {group.rows.map((row, rIdx) => {
                                const time = getRowTime(row);

                                const badgeStyle = 
                                  intervalInfo.level === 'very_high'
                                    ? "bg-purple-600/30 border-purple-400 text-purple-200 font-extrabold shadow-md shadow-purple-900/40 ring-1 ring-purple-400/50"
                                    : intervalInfo.level === 'high' 
                                    ? "bg-red-500/25 border-red-500/50 text-red-300 font-bold shadow-sm shadow-red-500/20"
                                    : intervalInfo.level === 'medium'
                                    ? "bg-amber-500/20 border-amber-500/40 text-amber-300 font-semibold"
                                    : "bg-white/10 border-white/20 text-gray-300";
                                
                                return (
                                  <div key={rIdx} className="w-full flex flex-col items-center gap-0.5">
                                    <span 
                                      className={cn("px-1.5 py-0.5 border rounded font-mono text-[10px] w-full text-center", badgeStyle)}
                                      title={
                                        intervalInfo.level === 'very_high'
                                          ? 'Mesmo Horário Exato (0 min - Muito Suspeito)'
                                          : intervalInfo.level === 'high' 
                                          ? 'Intervalo ≤ 30 min (Bem Provável)' 
                                          : intervalInfo.level === 'medium' 
                                          ? 'Intervalo entre 30 min e 1h (Um pouco menos provável)' 
                                          : 'Intervalo > 1h'
                                      }
                                    >
                                      {time || '--:--'}
                                    </span>
                                    {sourceFiles.length > 1 && row['ARQUIVO_ORIGEM'] && (
                                      <span 
                                        className="text-[7.5px] font-mono text-cyan-300/90 bg-cyan-950/60 border border-cyan-500/30 px-1 py-0.2 rounded truncate w-full text-center"
                                        title={`Arquivo de Origem: ${row['ARQUIVO_ORIGEM']}`}
                                      >
                                        {row['ARQUIVO_ORIGEM'].length > 14 ? `${row['ARQUIVO_ORIGEM'].substring(0, 12)}…` : row['ARQUIVO_ORIGEM']}
                                      </span>
                                    )}
                                  </div>
                                );
                              })}
                              {intervalInfo.minDiff !== null && (
                                <span className={cn(
                                  "text-[8px] font-mono mt-0.5 px-1 rounded uppercase tracking-wider text-center w-full truncate",
                                  intervalInfo.level === 'very_high' ? "text-purple-300 bg-purple-500/20 font-bold border border-purple-500/30" :
                                  intervalInfo.level === 'high' ? "text-red-400 bg-red-400/10 font-bold" :
                                  intervalInfo.level === 'medium' ? "text-amber-400 bg-amber-400/10" : "text-gray-400 bg-white/5"
                                )} title={intervalInfo.label}>
                                  {intervalInfo.level === 'very_high' ? 'Mesmo Horário (Muito Suspeito)' : intervalInfo.level === 'high' ? '≤30m (Provável)' : intervalInfo.level === 'medium' ? '30m-1h' : '>1h'}
                                </span>
                              )}
                            </>
                          );
                        })()}
                      </div>
                    </div>
                  ))
                )}
              </div>

              {/* Info Bar at Bottom */}
              <div className="p-4 bg-white/5 text-[10px] text-gray-400 border-t border-white/10 flex flex-col md:flex-row justify-between items-center gap-4">
                <div className="flex items-center gap-3 flex-wrap">
                  <span className="text-amber-500">●</span>
                  <span>Critério: Mesma Data, Prestador, Especialidade, Beneficiário, Procedimento, Tipo, Situação, Qtd e Conselho.</span>
                  <div className="flex items-center gap-3 ml-2 pl-3 border-l border-white/10">
                    <span className="flex items-center gap-1 text-purple-300 font-semibold"><span className="w-2 h-2 rounded-full bg-purple-400"></span> Mesmo Horário (Muito Suspeito)</span>
                    <span className="flex items-center gap-1 text-red-400 font-medium"><span className="w-2 h-2 rounded-full bg-red-500"></span> ≤ 30min (Bem Provável)</span>
                    <span className="flex items-center gap-1 text-amber-400 font-medium"><span className="w-2 h-2 rounded-full bg-amber-500"></span> 30min-1h (Pouco Menos)</span>
                    <span className="flex items-center gap-1 text-gray-400 font-medium"><span className="w-2 h-2 rounded-full bg-gray-500"></span> &gt; 1h</span>
                  </div>
                </div>
                <div className="italic opacity-60">
                  Total de {filteredDuplicates.length} incidentes identificados para conferência.
                </div>
              </div>
            </div>
          </>
        )}
      </div>
    )}
  </main>

      {/* Internal Scrollbar Styling */}
      <style dangerouslySetInnerHTML={{ __html: `
        .custom-scrollbar::-webkit-scrollbar {
          width: 6px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
          background: rgba(255, 255, 255, 0.02);
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
          background: rgba(255, 255, 255, 0.1);
          border-radius: 3px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
          background: rgba(255, 255, 255, 0.2);
        }
      `}} />

      {/* Footer Details */}
      <footer className="px-8 py-6 border-t border-white/10 flex flex-col md:flex-row justify-between items-center bg-[#07080A] gap-4">
        <div className="flex items-center gap-2">
          <div className={cn("w-2 h-2 rounded-full", data ? "bg-emerald-500" : "bg-gray-700")}></div>
          <span className="text-[10px] text-gray-500 uppercase tracking-widest">
            AuditCheck System — Status: {data ? "Processamento Concluído" : "Pronto para Importação"}
          </span>
        </div>
        <div className="text-[10px] text-gray-500 font-mono uppercase tracking-tighter">
          Integridade e Rigor na Auditoria de Transações HealthTech
        </div>
      </footer>
    </div>
  );
}

