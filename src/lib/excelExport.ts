import ExcelJS from 'exceljs';

export interface ITDiagnosticsStats {
  totalGroups: number;
  deltaZeroCount: number;
  instantPercent: number;
  delta1to5Count: number;
  microPercent: number;
  delta6to30Count: number;
  shortPercent: number;
  deltaOver30Count: number;
  extendedPercent: number;
  sameProtocolCount: number;
  sameProtocolPercent: number;
  differentProtocolCount: number;
  differentProtocolPercent: number;
  noProtocolCount: number;
  sameRemessaCount: number;
  sameRemessaPercent: number;
  differentRemessaCount: number;
  differentRemessaPercent: number;
  noRemessaCount: number;
  topProviderConcentrationPercent: number;
  top3ProvidersSharePercent: number;
  topProviderName: string;
  topProviderCount: number;
  topProvidersList: { provider: string; count: number; percent: number }[];
  topProceduresList: { procedure: string; count: number; percent: number }[];
  dayOfWeekDistribution: { dayName: string; count: number; percent: number }[];
  monthlyDistribution: { monthLabel: string; count: number; percent: number }[];
  shiftDistribution: { shiftName: string; count: number; percent: number }[];
  topBeneficiariesList: { beneficiary: string; count: number; percent: number }[];
  situationsDistribution: { situation: string; count: number; percent: number }[];
  dayOfMonthDistribution: { dayNum: number; dayLabel: string; count: number; percent: number }[];
  monthCycleDistribution: { cycleName: string; count: number; percent: number }[];
}

export interface ExportReportOptions {
  fileName?: string;
  transactionCount: number;
  duplicateGroups: any[];
  fileHasRemessa: boolean;
  fileHasProtocolo: boolean;
  selectedPrestadorFilter?: string;
  getColValue: (row: any, concept: string) => any;
  normalizeDate: (val: any) => { date: string; time: string };
  getGroupTimeInterval: (rows: any[]) => { minDiff: number | null; label: string; level: string };
  getRowTime?: (row: any) => string;
  hasMultiOriginFiles?: boolean;
  itStats?: ITDiagnosticsStats;
}

export async function exportFormattedExcelReport(options: ExportReportOptions) {
  const {
    fileName,
    transactionCount,
    duplicateGroups,
    fileHasRemessa,
    fileHasProtocolo,
    selectedPrestadorFilter,
    getColValue,
    normalizeDate,
    getGroupTimeInterval,
    getRowTime,
    hasMultiOriginFiles,
    itStats
  } = options;

  if (duplicateGroups.length === 0) return;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'AuditCheck HealthTech';
  workbook.lastModifiedBy = 'AuditCheck HealthTech';
  workbook.created = new Date();
  workbook.modified = new Date();

  // -------------------------------------------------------------
  // ABA 1: RELATÓRIO DE NOTIFICAÇÃO E AUDITORIA DE DUPLICIDADES
  // -------------------------------------------------------------
  const ws = workbook.addWorksheet('Notificação de Duplicidades', {
    views: [{ showGridLines: true }]
  });

  // Título e Cabeçalho do Ofício
  const headerBgColor = '0F172A'; // Slate 900
  const subHeaderBg = '1E293B'; // Slate 800
  const accentTeal = '0D9488'; // Teal 600
  const lightGrayBg = 'F8FAFC';
  const borderThinColor = 'CBD5E1';

  // Row 1: Título Principal
  ws.mergeCells('A1:P1');
  const titleRow = ws.getCell('A1');
  titleRow.value = 'AUDITORIA DE CONTAS MÉDICAS — NOTIFICAÇÃO DE DUPLICIDADE DE EXECUÇÕES';
  titleRow.font = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  titleRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: headerBgColor }
  };
  titleRow.alignment = { vertical: 'middle', horizontal: 'center' };
  ws.getRow(1).height = 36;

  // Row 2: Subtítulo e Instruções de Comunicação ao Prestador
  ws.mergeCells('A2:P2');
  const subtitleRow = ws.getCell('A2');
  subtitleRow.value = 'Prezado Prestador: Identificamos apontamentos com indícios de cobrança em duplicidade no faturamento apresentado. Solicitamos a conferência dos itens destacados e o retorno com a devida justificativa ou aceite de glosa na coluna designada.';
  subtitleRow.font = { name: 'Calibri', size: 10, italic: true, color: { argb: 'FFE2E8F0' } };
  subtitleRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: subHeaderBg }
  };
  subtitleRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  ws.getRow(2).height = 30;

  // Row 3: Cards Informativos / Metadados
  const now = new Date();
  const dateFormatted = `${now.getDate().toString().padStart(2, '0')}/${(now.getMonth() + 1).toString().padStart(2, '0')}/${now.getFullYear()} às ${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
  
  ws.mergeCells('A3:D3');
  const metaCell1 = ws.getCell('A3');
  metaCell1.value = `📅 Data da Emissão: ${dateFormatted}`;
  metaCell1.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF334155' } };
  metaCell1.alignment = { vertical: 'middle', horizontal: 'left' };

  ws.mergeCells('E3:H3');
  const metaCell2 = ws.getCell('E3');
  metaCell2.value = `📊 Total: ${duplicateGroups.length} grupo(s) (${transactionCount} reg.)`;
  metaCell2.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF0F766E' } };
  metaCell2.alignment = { vertical: 'middle', horizontal: 'left' };

  ws.mergeCells('I3:L3');
  const metaCell3 = ws.getCell('I3');
  metaCell3.value = selectedPrestadorFilter ? `🔍 Filtros: ${selectedPrestadorFilter}` : '🏥 Prestadores: Todos';
  metaCell3.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF1D4ED8' } };
  metaCell3.alignment = { vertical: 'middle', horizontal: 'left' };

  ws.mergeCells('M3:P3');
  const metaCell4 = ws.getCell('M3');
  metaCell4.value = `📁 Arquivo: ${fileName || 'Relatório de Auditoria'}`;
  metaCell4.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF64748B' } };
  metaCell4.alignment = { vertical: 'middle', horizontal: 'left' };

  ws.getRow(3).height = 22;
  ['A3', 'E3', 'I3', 'M3'].forEach(pos => {
    const c = ws.getCell(pos);
    c.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'F1F5F9' }
    };
    c.border = {
      top: { style: 'thin', color: { argb: borderThinColor } },
      bottom: { style: 'thin', color: { argb: borderThinColor } },
      left: { style: 'thin', color: { argb: borderThinColor } },
      right: { style: 'thin', color: { argb: borderThinColor } }
    };
  });

  // Espaçamento
  ws.getRow(4).height = 10;

  // Definição das Colunas da Tabela
  const columns: { header: string; key: string; width: number; align?: 'left' | 'center' | 'right'; bg?: string }[] = [
    { header: 'GRUPO', key: 'grupo', width: 10, align: 'center', bg: '1E293B' }
  ];

  if (fileHasRemessa) {
    columns.push({ header: 'Nº REMESSA', key: 'remessa', width: 16, align: 'center', bg: '1E3A8A' });
  }
  if (fileHasProtocolo) {
    columns.push({ header: 'Nº PROTOCOLO', key: 'protocolo', width: 18, align: 'center', bg: '1E3A8A' });
  }

  columns.push(
    { header: 'BENEFICIÁRIO', key: 'beneficiario', width: 30, align: 'left' },
    { header: 'MATRÍCULA BENEF.', key: 'matricula_bene', width: 18, align: 'center' },
    { header: 'PRESTADOR', key: 'prestador', width: 28, align: 'left' },
    { header: 'MATRÍCULA PREST.', key: 'matricula_prest', width: 18, align: 'center' },
    { header: 'ESPECIALIDADE', key: 'especialidade', width: 18, align: 'left' },
    { header: 'PROCEDIMENTO', key: 'procedimento', width: 34, align: 'left' },
    { header: 'TIPO TRANSAÇÃO', key: 'tipo', width: 16, align: 'center' },
    { header: 'DATA EXECUÇÃO', key: 'data', width: 15, align: 'center' },
    { header: 'HORÁRIO', key: 'horario', width: 12, align: 'center' },
    { header: 'QTD', key: 'quantidade', width: 8, align: 'center' },
    { header: 'SITUAÇÃO', key: 'situacao', width: 14, align: 'center' },
    { header: 'CRM / CONSELHO', key: 'conselho', width: 16, align: 'center' },
    { header: 'PROBABILIDADE / INTERVALO', key: 'probabilidade', width: 26, align: 'center' },
    { header: 'STATUS AUDITORIA', key: 'status_auditoria', width: 22, align: 'center' },
    ...(hasMultiOriginFiles ? [{ header: 'ARQUIVO DE ORIGEM', key: 'arquivo_origem', width: 28, align: 'left' as const, bg: '334155' }] : []),
    { header: 'PARECER / RESPOSTA DO PRESTADOR', key: 'resposta_prestador', width: 38, align: 'left' }
  );

  const headerRowIdx = 5;
  const headerRow = ws.getRow(headerRowIdx);
  headerRow.height = 28;

  columns.forEach((col, idx) => {
    const cell = headerRow.getCell(idx + 1);
    cell.value = col.header;
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: col.bg || '1E293B' }
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = {
      top: { style: 'medium', color: { argb: '0F172A' } },
      bottom: { style: 'medium', color: { argb: '0F172A' } },
      left: { style: 'thin', color: { argb: '475569' } },
      right: { style: 'thin', color: { argb: '475569' } }
    };
    ws.getColumn(idx + 1).width = col.width;
  });

  // Preenchimento dos Dados com Linhas e Cores
  let currentRowIdx = 6;
  const thinBorder = {
    top: { style: 'thin' as const, color: { argb: 'CBD5E1' } },
    bottom: { style: 'thin' as const, color: { argb: 'CBD5E1' } },
    left: { style: 'thin' as const, color: { argb: 'CBD5E1' } },
    right: { style: 'thin' as const, color: { argb: 'CBD5E1' } }
  };

  const groupSeparatorBorder = {
    top: { style: 'medium' as const, color: { argb: '64748B' } },
    bottom: { style: 'thin' as const, color: { argb: 'CBD5E1' } },
    left: { style: 'thin' as const, color: { argb: 'CBD5E1' } },
    right: { style: 'thin' as const, color: { argb: 'CBD5E1' } }
  };

  duplicateGroups.forEach((group, groupIdx) => {
    const intervalInfo = getGroupTimeInterval(group.rows);
    const isHigh = intervalInfo.level === 'high';
    const isMedium = intervalInfo.level === 'medium';

    // Determina a cor de destaque da probabilidade
    let probBg = 'F1F5F9';
    let probFontColor = '475569';
    let probLabel = intervalInfo.label || 'Outros (> 1h)';

    if (isHigh) {
      probBg = 'FEE2E2'; // Light Red
      probFontColor = '991B1B'; // Dark Red
    } else if (isMedium) {
      probBg = 'FEF3C7'; // Light Amber
      probFontColor = '92400E'; // Dark Amber
    }

    const groupNumberStr = `#${(groupIdx + 1).toString().padStart(3, '0')}`;

    group.rows.forEach((row: any, rIdx: number) => {
      const { date, time } = normalizeDate(getColValue(row, 'DATA'));
      const hourVal = getColValue(row, 'HORA');
      let effectiveTime = getRowTime ? getRowTime(row) : time;
      if (!effectiveTime && hourVal !== undefined && hourVal !== null && hourVal !== '') {
        const hNorm = normalizeDate(hourVal);
        if (hNorm.time) effectiveTime = hNorm.time;
        else if (typeof hourVal === 'string' && hourVal.includes(':')) effectiveTime = hourVal.trim().substring(0, 5);
      }

      const remessaVal = getColValue(row, 'REMESSA') ?? '';
      const protocoloVal = getColValue(row, 'PROTOCOLO') ?? '';

      const beneNome = getColValue(row, 'BENE_NOME') ?? group.params.beneficiario ?? '';
      const beneMat = getColValue(row, 'BENE_MATRICULA') ?? group.params.beneficiario_matricula ?? '';
      const prestNome = getColValue(row, 'PREST_NOME') ?? group.params.prestador ?? '';
      const prestMat = getColValue(row, 'PREST_MATRICULA') ?? group.params.prestador_matricula ?? '';
      const procNome = getColValue(row, 'PROCEDIMENTO') ?? group.params.procedimento ?? '';
      const espNome = getColValue(row, 'ESPECIALIDADE') ?? group.params.especialidade ?? '';
      const tipoVal = getColValue(row, 'TIPO') ?? group.params.tipo ?? '';
      const qtdVal = getColValue(row, 'QUANTIDADE') ?? group.params.quantidade ?? '';
      const sitVal = getColValue(row, 'SITUACAO') ?? group.params.situacao ?? '';
      
      const cExecNum = getColValue(row, 'NUMERO_CONSELHO_EXECUTANTE') ?? group.params.conselhoExecNum ?? '';
      const cExecTipo = getColValue(row, 'TIPO_CONSELHO_EXECUTANTE') ?? group.params.conselhoExecTipo ?? 'CRM';
      const cExecUf = getColValue(row, 'UF_CONSELHO_EXECUTANTE') ?? group.params.conselhoExecUf ?? '';
      const conselhoStr = cExecNum ? `${cExecTipo} ${cExecNum}${cExecUf ? `/${cExecUf}` : ''}` : (getColValue(row, 'CONSELHO') ?? group.params.conselho ?? '');

      const dataRow = ws.getRow(currentRowIdx);
      dataRow.height = 24;

      const isFirstInGroup = rIdx === 0;
      const rowBorder = isFirstInGroup ? groupSeparatorBorder : thinBorder;
      const isAltRow = groupIdx % 2 === 1;
      const baseRowBg = isAltRow ? 'F8FAFC' : 'FFFFFF';

      let colPtr = 1;

      // Coluna Grupo
      const cGrupo = dataRow.getCell(colPtr++);
      cGrupo.value = groupNumberStr;
      cGrupo.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF1E293B' } };
      cGrupo.alignment = { vertical: 'middle', horizontal: 'center' };
      cGrupo.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isAltRow ? 'F1F5F9' : 'E2E8F0' } };
      cGrupo.border = rowBorder;

      // Remessa se aplicável
      if (fileHasRemessa) {
        const cRem = dataRow.getCell(colPtr++);
        cRem.value = String(remessaVal);
        cRem.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF1E3A8A' } };
        cRem.alignment = { vertical: 'middle', horizontal: 'center' };
        cRem.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EFF6FF' } }; // Azul claro
        cRem.border = rowBorder;
      }

      // Protocolo se aplicável
      if (fileHasProtocolo) {
        const cProt = dataRow.getCell(colPtr++);
        cProt.value = String(protocoloVal);
        cProt.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF1E3A8A' } };
        cProt.alignment = { vertical: 'middle', horizontal: 'center' };
        cProt.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EFF6FF' } }; // Azul claro
        cProt.border = rowBorder;
      }

      // Beneficiário
      const cBene = dataRow.getCell(colPtr++);
      cBene.value = String(beneNome).toUpperCase();
      cBene.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF0F172A' } };
      cBene.alignment = { vertical: 'middle', horizontal: 'left' };
      cBene.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: baseRowBg } };
      cBene.border = rowBorder;

      // Matrícula Beneficiário
      const cBeneMat = dataRow.getCell(colPtr++);
      cBeneMat.value = String(beneMat);
      cBeneMat.font = { name: 'Calibri', size: 9, color: { argb: 'FF047857' } }; // Emerald
      cBeneMat.alignment = { vertical: 'middle', horizontal: 'center' };
      cBeneMat.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: baseRowBg } };
      cBeneMat.border = rowBorder;

      // Prestador
      const cPrest = dataRow.getCell(colPtr++);
      cPrest.value = String(prestNome).toUpperCase();
      cPrest.font = { name: 'Calibri', size: 9, color: { argb: 'FF1E293B' } };
      cPrest.alignment = { vertical: 'middle', horizontal: 'left' };
      cPrest.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: baseRowBg } };
      cPrest.border = rowBorder;

      // Matrícula Prestador
      const cPrestMat = dataRow.getCell(colPtr++);
      cPrestMat.value = String(prestMat);
      cPrestMat.font = { name: 'Calibri', size: 9, color: { argb: 'FF475569' } };
      cPrestMat.alignment = { vertical: 'middle', horizontal: 'center' };
      cPrestMat.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: baseRowBg } };
      cPrestMat.border = rowBorder;

      // Especialidade
      const cEsp = dataRow.getCell(colPtr++);
      cEsp.value = String(espNome);
      cEsp.font = { name: 'Calibri', size: 9, italic: true, color: { argb: 'FF475569' } };
      cEsp.alignment = { vertical: 'middle', horizontal: 'left' };
      cEsp.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: baseRowBg } };
      cEsp.border = rowBorder;

      // Procedimento
      const cProc = dataRow.getCell(colPtr++);
      cProc.value = String(procNome);
      cProc.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF0F172A' } };
      cProc.alignment = { vertical: 'middle', horizontal: 'left' };
      cProc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: baseRowBg } };
      cProc.border = rowBorder;

      // Tipo Transação
      const cTipo = dataRow.getCell(colPtr++);
      cTipo.value = String(tipoVal);
      cTipo.font = { name: 'Calibri', size: 9, color: { argb: 'FF2563EB' } };
      cTipo.alignment = { vertical: 'middle', horizontal: 'center' };
      cTipo.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: baseRowBg } };
      cTipo.border = rowBorder;

      // Data
      const cData = dataRow.getCell(colPtr++);
      cData.value = String(date || group.params.data);
      cData.font = { name: 'Calibri', size: 9, color: { argb: 'FF1E293B' } };
      cData.alignment = { vertical: 'middle', horizontal: 'center' };
      cData.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: baseRowBg } };
      cData.border = rowBorder;

      // Horário (Destaque)
      const cHora = dataRow.getCell(colPtr++);
      cHora.value = (effectiveTime && effectiveTime !== '0') ? effectiveTime : '--:--';
      cHora.font = { 
        name: 'Calibri', 
        size: 9, 
        bold: true, 
        color: { argb: isHigh ? 'FF991B1B' : isMedium ? 'FF92400E' : 'FF334155' } 
      };
      cHora.alignment = { vertical: 'middle', horizontal: 'center' };
      cHora.fill = { 
        type: 'pattern', 
        pattern: 'solid', 
        fgColor: { argb: isHigh ? 'FEE2E2' : isMedium ? 'FEF3C7' : baseRowBg } 
      };
      cHora.border = rowBorder;

      // Qtd
      const cQtd = dataRow.getCell(colPtr++);
      cQtd.value = String(qtdVal);
      cQtd.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF0F172A' } };
      cQtd.alignment = { vertical: 'middle', horizontal: 'center' };
      cQtd.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: baseRowBg } };
      cQtd.border = rowBorder;

      // Situação
      const cSit = dataRow.getCell(colPtr++);
      cSit.value = String(sitVal);
      const isAprovada = String(sitVal).toLowerCase() === 'aprovada';
      cSit.font = { 
        name: 'Calibri', 
        size: 9, 
        bold: true, 
        color: { argb: isAprovada ? 'FF065F46' : 'FF475569' } 
      };
      cSit.alignment = { vertical: 'middle', horizontal: 'center' };
      cSit.fill = { 
        type: 'pattern', 
        pattern: 'solid', 
        fgColor: { argb: isAprovada ? 'ECFDF5' : baseRowBg } 
      };
      cSit.border = rowBorder;

      // Conselho
      const cConselho = dataRow.getCell(colPtr++);
      cConselho.value = String(conselhoStr);
      cConselho.font = { name: 'Calibri', size: 9, color: { argb: 'FF6B21A8' } };
      cConselho.alignment = { vertical: 'middle', horizontal: 'center' };
      cConselho.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: baseRowBg } };
      cConselho.border = rowBorder;

      // Probabilidade / Intervalo
      const cProb = dataRow.getCell(colPtr++);
      cProb.value = probLabel;
      cProb.font = { name: 'Calibri', size: 9, bold: true, color: { argb: `FF${probFontColor}` } };
      cProb.alignment = { vertical: 'middle', horizontal: 'center' };
      cProb.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: probBg } };
      cProb.border = rowBorder;

      // Status Auditoria
      const cStatus = dataRow.getCell(colPtr++);
      cStatus.value = 'SUSPEITA DE DUPLICIDADE';
      cStatus.font = { name: 'Calibri', size: 8, bold: true, color: { argb: 'FFB91C1C' } };
      cStatus.alignment = { vertical: 'middle', horizontal: 'center' };
      cStatus.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FEF2F2' } };
      cStatus.border = rowBorder;

      // Arquivo de Origem se lote consolidado
      if (hasMultiOriginFiles) {
        const cOrig = dataRow.getCell(colPtr++);
        cOrig.value = String(row['ARQUIVO_ORIGEM'] || '');
        cOrig.font = { name: 'Calibri', size: 8, italic: true, color: { argb: 'FF475569' } };
        cOrig.alignment = { vertical: 'middle', horizontal: 'left' };
        cOrig.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: baseRowBg } };
        cOrig.border = rowBorder;
      }

      // Coluna Para Resposta do Prestador (Espaço em branco com borda pontilhada)
      const cResposta = dataRow.getCell(colPtr++);
      cResposta.value = '';
      cResposta.font = { name: 'Calibri', size: 9, color: { argb: 'FF0F172A' } };
      cResposta.alignment = { vertical: 'middle', horizontal: 'left' };
      cResposta.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF' } };
      cResposta.border = {
        top: { style: 'dashed', color: { argb: '94A3B8' } },
        bottom: { style: 'dashed', color: { argb: '94A3B8' } },
        left: { style: 'thin', color: { argb: '94A3B8' } },
        right: { style: 'thin', color: { argb: '94A3B8' } }
      };

      currentRowIdx++;
    });
  });

  // Linha final de Totalizador
  const totalRow = ws.getRow(currentRowIdx);
  totalRow.height = 24;
  ws.mergeCells(`A${currentRowIdx}:E${currentRowIdx}`);
  const totCell = ws.getCell(`A${currentRowIdx}`);
  totCell.value = `TOTAL: ${duplicateGroups.length} incidentes de duplicidade agrupados`;
  totCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  totCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0F172A' } };
  totCell.alignment = { vertical: 'middle', horizontal: 'left' };

  // -------------------------------------------------------------
  // ABA 2: RESUMO EXECUTIVO PARA CONTATO COM PRESTADORES
  // -------------------------------------------------------------
  const wsPrestadores = workbook.addWorksheet('Resumo por Prestador', {
    views: [{ showGridLines: true }]
  });

  wsPrestadores.mergeCells('A1:G1');
  const titP = wsPrestadores.getCell('A1');
  titP.value = 'RESUMO CONSOLIDADO POR PRESTADOR — GESTÃO DE CONTATOS E REGULARIZAÇÃO';
  titP.font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
  titP.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0F172A' } };
  titP.alignment = { vertical: 'middle', horizontal: 'center' };
  wsPrestadores.getRow(1).height = 32;

  // Cabeçalho da Aba Prestadores
  const pHeaders = [
    { name: 'NOME DO PRESTADOR', width: 34 },
    { name: 'MATRÍCULA / CÓDIGO', width: 20 },
    { name: 'TOTAL GRUPOS DUPLICADOS', width: 24 },
    { name: 'CRÍTICAS (≤ 30 MIN)', width: 22 },
    { name: 'MÉDIAS (30 MIN - 1H)', width: 22 },
    { name: 'STATUS DE CONTATO', width: 24 },
    { name: 'OBSERVAÇÕES / E-MAIL DE CONTATO', width: 38 }
  ];

  const pHeaderRow = wsPrestadores.getRow(3);
  pHeaderRow.height = 26;
  pHeaders.forEach((h, i) => {
    const c = pHeaderRow.getCell(i + 1);
    c.value = h.name;
    c.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1E293B' } };
    c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    c.border = thinBorder;
    wsPrestadores.getColumn(i + 1).width = h.width;
  });

  // Agrupamento por Prestador
  const provMap: Record<string, { nome: string; matricula: string; total: number; high: number; medium: number }> = {};
  duplicateGroups.forEach(g => {
    const prov = g.params.prestador || 'Não informado';
    const mat = g.params.prestador_matricula || '';
    if (!provMap[prov]) {
      provMap[prov] = { nome: prov, matricula: mat, total: 0, high: 0, medium: 0 };
    }
    provMap[prov].total++;
    const intervalInfo = getGroupTimeInterval(g.rows);
    if (intervalInfo.level === 'high') provMap[prov].high++;
    else if (intervalInfo.level === 'medium') provMap[prov].medium++;
  });

  let pRowIdx = 4;
  Object.values(provMap).sort((a, b) => b.total - a.total).forEach((p, idx) => {
    const r = wsPrestadores.getRow(pRowIdx);
    r.height = 22;
    const isAlt = idx % 2 === 1;
    const bg = isAlt ? 'F8FAFC' : 'FFFFFF';

    const c1 = r.getCell(1);
    c1.value = p.nome.toUpperCase();
    c1.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF0F172A' } };
    c1.alignment = { vertical: 'middle', horizontal: 'left' };
    c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    c1.border = thinBorder;

    const c2 = r.getCell(2);
    c2.value = p.matricula;
    c2.font = { name: 'Calibri', size: 9, color: { argb: 'FF475569' } };
    c2.alignment = { vertical: 'middle', horizontal: 'center' };
    c2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    c2.border = thinBorder;

    const c3 = r.getCell(3);
    c3.value = p.total;
    c3.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FF1E293B' } };
    c3.alignment = { vertical: 'middle', horizontal: 'center' };
    c3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    c3.border = thinBorder;

    const c4 = r.getCell(4);
    c4.value = p.high;
    c4.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF991B1B' } };
    c4.alignment = { vertical: 'middle', horizontal: 'center' };
    c4.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: p.high > 0 ? 'FEE2E2' : bg } };
    c4.border = thinBorder;

    const c5 = r.getCell(5);
    c5.value = p.medium;
    c5.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF92400E' } };
    c5.alignment = { vertical: 'middle', horizontal: 'center' };
    c5.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: p.medium > 0 ? 'FEF3C7' : bg } };
    c5.border = thinBorder;

    const c6 = r.getCell(6);
    c6.value = 'Pendente de Notificação';
    c6.font = { name: 'Calibri', size: 9, italic: true, color: { argb: 'FF64748B' } };
    c6.alignment = { vertical: 'middle', horizontal: 'center' };
    c6.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F1F5F9' } };
    c6.border = thinBorder;

    const c7 = r.getCell(7);
    c7.value = '';
    c7.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF' } };
    c7.border = {
      top: { style: 'dashed', color: { argb: '94A3B8' } },
      bottom: { style: 'dashed', color: { argb: '94A3B8' } },
      left: { style: 'thin', color: { argb: '94A3B8' } },
      right: { style: 'thin', color: { argb: '94A3B8' } }
    };

    pRowIdx++;
  });

  if (itStats) {
    addITDiagnosticsWorksheet(workbook, itStats, transactionCount, 'Padrões e Comparativos');
  }

  // Geração do arquivo e download no browser
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  const outName = selectedPrestadorFilter
    ? `Notificacao_Duplicidades_${selectedPrestadorFilter.replace(/[^a-zA-Z0-9]/g, '_')}.xlsx`
    : 'Notificacao_Auditoria_Duplicidades.xlsx';
  a.download = outName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.URL.revokeObjectURL(url);
}

export interface DashboardReportOptions {
  transactionCount: number;
  duplicateGroups: any[];
  categoryStats: { high: number; medium: number; low: number; total: number };
  dashboardStats: {
    topDays: { day: string; count: number }[];
    topHours: { hour: string; count: number }[];
    topProviders: { provider: string; count: number }[];
    topProcedures: { procedure: string; count: number }[];
  };
  itStats?: ITDiagnosticsStats;
}

export async function exportFormattedDashboardReport(options: DashboardReportOptions) {
  const { transactionCount, duplicateGroups, categoryStats, dashboardStats, itStats } = options;
  if (duplicateGroups.length === 0) return;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'AuditCheck HealthTech';

  const ws = workbook.addWorksheet('Resumo Executivo & Picos', {
    views: [{ showGridLines: true }]
  });

  const thinBorder = {
    top: { style: 'thin' as const, color: { argb: 'CBD5E1' } },
    bottom: { style: 'thin' as const, color: { argb: 'CBD5E1' } },
    left: { style: 'thin' as const, color: { argb: 'CBD5E1' } },
    right: { style: 'thin' as const, color: { argb: 'CBD5E1' } }
  };

  // Banner Principal
  ws.mergeCells('A1:G1');
  const tit = ws.getCell('A1');
  tit.value = 'DASHBOARD EXECUTIVO DE AUDITORIA — PICOS E INCIDÊNCIAS DE DUPLICIDADE';
  tit.font = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  tit.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0F172A' } };
  tit.alignment = { vertical: 'middle', horizontal: 'center' };
  ws.getRow(1).height = 36;

  // Subtítulo
  ws.mergeCells('A2:G2');
  const sub = ws.getCell('A2');
  sub.value = 'Visão analítica gerencial para subsidiar negociações, notificações a prestadores e correções de fluxos de faturamento.';
  sub.font = { name: 'Calibri', size: 10, italic: true, color: { argb: 'FFE2E8F0' } };
  sub.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1E293B' } };
  sub.alignment = { vertical: 'middle', horizontal: 'center' };
  ws.getRow(2).height = 24;

  // Cards de Indicadores
  const cards = [
    { title: 'Total Registros Analisados', value: transactionCount, color: '0F172A', textColor: 'FFFFFF', col: 'A', endCol: 'B' },
    { title: 'Grupos em Duplicidade', value: categoryStats.total, color: '7F1D1D', textColor: 'FFFFFF', col: 'C', endCol: 'C' },
    { title: '≤ 30min (Crítico/Provável)', value: categoryStats.high, color: '991B1B', textColor: 'FFFFFF', col: 'D', endCol: 'E' },
    { title: '30m - 1h (Moderado)', value: categoryStats.medium, color: 'B45309', textColor: 'FFFFFF', col: 'F', endCol: 'F' },
    { title: '> 1h (Baixo)', value: categoryStats.low, color: '475569', textColor: 'FFFFFF', col: 'G', endCol: 'G' }
  ];

  ws.getRow(4).height = 22;
  ws.getRow(5).height = 28;

  cards.forEach(card => {
    const range = `${card.col}4:${card.endCol}4`;
    if (card.col !== card.endCol) ws.mergeCells(range);
    const cTitle = ws.getCell(`${card.col}4`);
    cTitle.value = card.title.toUpperCase();
    cTitle.font = { name: 'Calibri', size: 8, bold: true, color: { argb: 'FF475569' } };
    cTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F1F5F9' } };
    cTitle.alignment = { vertical: 'middle', horizontal: 'center' };
    cTitle.border = thinBorder;

    const valRange = `${card.col}5:${card.endCol}5`;
    if (card.col !== card.endCol) ws.mergeCells(valRange);
    const cVal = ws.getCell(`${card.col}5`);
    cVal.value = card.value;
    cVal.font = { name: 'Calibri', size: 14, bold: true, color: { argb: `FF${card.textColor}` } };
    cVal.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: card.color } };
    cVal.alignment = { vertical: 'middle', horizontal: 'center' };
    cVal.border = thinBorder;
  });

  // Tabela 1: Picos por Dia (Colunas A-C) e Tabela 2: Horários de Pico (Colunas E-G)
  ws.mergeCells('A7:C7');
  const t1 = ws.getCell('A7');
  t1.value = '📅 DIAS COM MAIOR INCIDÊNCIA DE DUPLICIDADES';
  t1.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  t1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '047857' } }; // Emerald
  t1.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.getRow(7).height = 26;

  ws.mergeCells('E7:G7');
  const t2 = ws.getCell('E7');
  t2.value = '⏰ FAIXAS HORÁRIAS DE PICO';
  t2.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  t2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'B45309' } }; // Amber
  t2.alignment = { vertical: 'middle', horizontal: 'left' };

  // Sub-headers
  ws.getCell('A8').value = 'DATA';
  ws.getCell('B8').value = 'QUANTIDADE';
  ws.getCell('C8').value = '% DO TOTAL';
  ['A8', 'B8', 'C8'].forEach(p => {
    const c = ws.getCell(p);
    c.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '065F46' } };
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    c.border = thinBorder;
  });

  ws.getCell('E8').value = 'FAIXA HORÁRIA';
  ws.getCell('F8').value = 'OCORRÊNCIAS';
  ws.getCell('G8').value = '% DO TOTAL';
  ['E8', 'F8', 'G8'].forEach(p => {
    const c = ws.getCell(p);
    c.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '92400E' } };
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    c.border = thinBorder;
  });
  ws.getRow(8).height = 20;

  const maxRows = Math.max(dashboardStats.topDays.length, dashboardStats.topHours.length);
  const totalOccurrences = dashboardStats.topDays.reduce((acc, d) => acc + d.count, 0) || 1;
  const totalHoursOccurrences = dashboardStats.topHours.reduce((acc, h) => acc + h.count, 0) || 1;

  for (let i = 0; i < maxRows; i++) {
    const rIdx = 9 + i;
    const r = ws.getRow(rIdx);
    r.height = 20;
    const isAlt = i % 2 === 1;
    const bg = isAlt ? 'F8FAFC' : 'FFFFFF';

    // Day
    if (i < dashboardStats.topDays.length) {
      const d = dashboardStats.topDays[i];
      const cA = r.getCell(1);
      cA.value = d.day || 'Não informada';
      cA.font = { name: 'Calibri', size: 9, color: { argb: 'FF0F172A' } };
      cA.alignment = { vertical: 'middle', horizontal: 'center' };
      cA.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
      cA.border = thinBorder;

      const cB = r.getCell(2);
      cB.value = d.count;
      cB.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF047857' } };
      cB.alignment = { vertical: 'middle', horizontal: 'center' };
      cB.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
      cB.border = thinBorder;

      const cC = r.getCell(3);
      cC.value = `${Math.round((d.count / totalOccurrences) * 100)}%`;
      cC.font = { name: 'Calibri', size: 9, color: { argb: 'FF64748B' } };
      cC.alignment = { vertical: 'middle', horizontal: 'center' };
      cC.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
      cC.border = thinBorder;
    }

    // Hour
    if (i < dashboardStats.topHours.length) {
      const h = dashboardStats.topHours[i];
      const cE = r.getCell(5);
      cE.value = h.hour;
      cE.font = { name: 'Calibri', size: 9, color: { argb: 'FF0F172A' } };
      cE.alignment = { vertical: 'middle', horizontal: 'center' };
      cE.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
      cE.border = thinBorder;

      const cF = r.getCell(6);
      cF.value = h.count;
      cF.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFB45309' } };
      cF.alignment = { vertical: 'middle', horizontal: 'center' };
      cF.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
      cF.border = thinBorder;

      const cG = r.getCell(7);
      cG.value = `${Math.round((h.count / totalHoursOccurrences) * 100)}%`;
      cG.font = { name: 'Calibri', size: 9, color: { argb: 'FF64748B' } };
      cG.alignment = { vertical: 'middle', horizontal: 'center' };
      cG.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
      cG.border = thinBorder;
    }
  }

  // Tabela 3: Top Prestadores
  const nextSectionRow = 9 + maxRows + 2;
  ws.mergeCells(`A${nextSectionRow}:C${nextSectionRow}`);
  const t3 = ws.getCell(`A${nextSectionRow}`);
  t3.value = '🏥 PRESTADORES COM MAIOR RECORRÊNCIA';
  t3.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  t3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1D4ED8' } };
  t3.alignment = { vertical: 'middle', horizontal: 'left' };

  ws.mergeCells(`E${nextSectionRow}:G${nextSectionRow}`);
  const t4 = ws.getCell(`E${nextSectionRow}`);
  t4.value = '📋 PROCEDIMENTOS MAIS AFETADOS';
  t4.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  t4.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '7E22CE' } };
  t4.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.getRow(nextSectionRow).height = 26;

  // Sub-headers
  ws.getCell(`A${nextSectionRow + 1}`).value = 'PRESTADOR';
  ws.getCell(`B${nextSectionRow + 1}`).value = 'CASOS';
  ws.getCell(`C${nextSectionRow + 1}`).value = 'AÇÃO SUGERIDA';
  ['A', 'B', 'C'].forEach(col => {
    const c = ws.getCell(`${col}${nextSectionRow + 1}`);
    c.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1E3A8A' } };
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    c.border = thinBorder;
  });

  ws.getCell(`E${nextSectionRow + 1}`).value = 'PROCEDIMENTO';
  ws.getCell(`F${nextSectionRow + 1}`).value = 'CASOS';
  ws.getCell(`G${nextSectionRow + 1}`).value = 'IMPACTO';
  ['E', 'F', 'G'].forEach(col => {
    const c = ws.getCell(`${col}${nextSectionRow + 1}`);
    c.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '6B21A8' } };
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    c.border = thinBorder;
  });
  ws.getRow(nextSectionRow + 1).height = 20;

  const maxItems2 = Math.max(dashboardStats.topProviders.length, dashboardStats.topProcedures.length);
  for (let i = 0; i < maxItems2; i++) {
    const rIdx = nextSectionRow + 2 + i;
    const r = ws.getRow(rIdx);
    r.height = 20;
    const isAlt = i % 2 === 1;
    const bg = isAlt ? 'F8FAFC' : 'FFFFFF';

    if (i < dashboardStats.topProviders.length) {
      const p = dashboardStats.topProviders[i];
      const cA = r.getCell(1);
      cA.value = p.provider;
      cA.font = { name: 'Calibri', size: 9, color: { argb: 'FF0F172A' } };
      cA.alignment = { vertical: 'middle', horizontal: 'left' };
      cA.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
      cA.border = thinBorder;

      const cB = r.getCell(2);
      cB.value = p.count;
      cB.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF1D4ED8' } };
      cB.alignment = { vertical: 'middle', horizontal: 'center' };
      cB.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
      cB.border = thinBorder;

      const cC = r.getCell(3);
      cC.value = p.count > 5 ? 'Notificação Prioritária' : 'Acompanhamento';
      cC.font = { name: 'Calibri', size: 8, italic: true, color: { argb: p.count > 5 ? 'FF991B1B' : 'FF475569' } };
      cC.alignment = { vertical: 'middle', horizontal: 'center' };
      cC.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
      cC.border = thinBorder;
    }

    if (i < dashboardStats.topProcedures.length) {
      const pr = dashboardStats.topProcedures[i];
      const cE = r.getCell(5);
      cE.value = pr.procedure;
      cE.font = { name: 'Calibri', size: 9, color: { argb: 'FF0F172A' } };
      cE.alignment = { vertical: 'middle', horizontal: 'left' };
      cE.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
      cE.border = thinBorder;

      const cF = r.getCell(6);
      cF.value = pr.count;
      cF.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF7E22CE' } };
      cF.alignment = { vertical: 'middle', horizontal: 'center' };
      cF.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
      cF.border = thinBorder;

      const cG = r.getCell(7);
      cG.value = 'Alto Risco';
      cG.font = { name: 'Calibri', size: 8, bold: true, color: { argb: 'FF991B1B' } };
      cG.alignment = { vertical: 'middle', horizontal: 'center' };
      cG.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FEE2E2' } };
      cG.border = thinBorder;
    }
  }

  // Auto-width
  ws.getColumn(1).width = 30;
  ws.getColumn(2).width = 16;
  ws.getColumn(3).width = 24;
  ws.getColumn(4).width = 6;
  ws.getColumn(5).width = 32;
  ws.getColumn(6).width = 16;
  ws.getColumn(7).width = 22;

  if (itStats) {
    addITDiagnosticsWorksheet(workbook, itStats, transactionCount, 'Padrões e Comparativos');
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'Dashboard_Auditoria_Duplicidades_Formatado.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.URL.revokeObjectURL(url);
}

export function addITDiagnosticsWorksheet(
  workbook: ExcelJS.Workbook, 
  itStats: ITDiagnosticsStats, 
  transactionCount: number,
  sheetName: string = 'Padrões e Comparativos'
) {
  const ws = workbook.addWorksheet(sheetName, {
    views: [{ showGridLines: true }]
  });

  const thinBorder = {
    top: { style: 'thin' as const, color: { argb: 'CBD5E1' } },
    bottom: { style: 'thin' as const, color: { argb: 'CBD5E1' } },
    left: { style: 'thin' as const, color: { argb: 'CBD5E1' } },
    right: { style: 'thin' as const, color: { argb: 'CBD5E1' } }
  };

  // Header Banner
  ws.mergeCells('A1:G1');
  const tit = ws.getCell('A1');
  tit.value = 'ANÁLISE ESTATÍSTICA DE PADRÕES E CRUZAMENTO DE DUPLICIDADES';
  tit.font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
  tit.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0F172A' } }; // Slate 900
  tit.alignment = { vertical: 'middle', horizontal: 'center' };
  ws.getRow(1).height = 34;

  ws.mergeCells('A2:G2');
  const sub = ws.getCell('A2');
  sub.value = 'Relatório analítico de cruzamento de dados: volumetria por dias da semana, evolução mensal, turnos do dia, beneficiários recorrentes e situação.';
  sub.font = { name: 'Calibri', size: 9, italic: true, color: { argb: 'FFE2E8F0' } };
  sub.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1E293B' } };
  sub.alignment = { vertical: 'middle', horizontal: 'center' };
  ws.getRow(2).height = 24;

  // Row 4-5: Cards Rápidos de Resumo Estatístico
  const cards = [
    { label: 'Grupos Analisados', val: `${itStats.totalGroups} (${transactionCount} reg.)`, bg: '1E293B', col: 'A', endCol: 'A' },
    { label: 'Simultâneos (0m)', val: `${itStats.deltaZeroCount} (${itStats.instantPercent}%)`, bg: '7F1D1D', col: 'B', endCol: 'C' },
    { label: 'Micro-intervalo (1-5m)', val: `${itStats.delta1to5Count} (${itStats.microPercent}%)`, bg: '92400E', col: 'D', endCol: 'E' },
    { label: 'Concentração Prestador Líder', val: `${itStats.topProviderConcentrationPercent}% (${(itStats.topProviderName || 'N/A').substring(0, 18)})`, bg: '065F46', col: 'F', endCol: 'G' }
  ];

  ws.getRow(4).height = 18;
  ws.getRow(5).height = 26;
  cards.forEach(card => {
    if (card.col !== card.endCol) {
      ws.mergeCells(`${card.col}4:${card.endCol}4`);
      ws.mergeCells(`${card.col}5:${card.endCol}5`);
    }
    const c1 = ws.getCell(`${card.col}4`);
    c1.value = card.label.toUpperCase();
    c1.font = { name: 'Calibri', size: 8, bold: true, color: { argb: 'FF64748B' } };
    c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F1F5F9' } };
    c1.alignment = { vertical: 'middle', horizontal: 'center' };
    c1.border = thinBorder;

    const c2 = ws.getCell(`${card.col}5`);
    c2.value = card.val;
    c2.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    c2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: card.bg } };
    c2.alignment = { vertical: 'middle', horizontal: 'center' };
    c2.border = thinBorder;
  });

  let curRow = 7;

  // Tabela 1: Comparativo por Dia da Semana
  ws.mergeCells(`A${curRow}:G${curRow}`);
  const t1 = ws.getCell(`A${curRow}`);
  t1.value = '1. COMPARATIVO DE INCIDÊNCIA POR DIA DA SEMANA';
  t1.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  t1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0369A1' } }; // Sky 700
  t1.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.getRow(curRow).height = 24;

  curRow++;
  ws.mergeCells(`A${curRow}:C${curRow}`);
  ws.getCell(`A${curRow}`).value = 'DIA DA SEMANA';
  ws.getCell(`D${curRow}`).value = 'QUANTIDADE';
  ws.getCell(`E${curRow}`).value = '% DO TOTAL';
  ws.mergeCells(`F${curRow}:G${curRow}`);
  ws.getCell(`F${curRow}`).value = 'DISTRIBUIÇÃO VISUAL';
  ['A', 'D', 'E', 'F'].forEach(cKey => {
    const c = ws.getCell(`${cKey}${curRow}`);
    c.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '075985' } };
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    c.border = thinBorder;
  });
  ws.getRow(curRow).height = 22;

  (itStats.dayOfWeekDistribution || []).forEach((dow, idx) => {
    curRow++;
    ws.mergeCells(`A${curRow}:C${curRow}`);
    ws.mergeCells(`F${curRow}:G${curRow}`);
    const r = ws.getRow(curRow);
    r.height = 20;
    const isAlt = idx % 2 === 1;
    const bgRow = isAlt ? 'F8FAFC' : 'FFFFFF';

    const c1 = r.getCell(1);
    c1.value = dow.dayName;
    c1.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF0F172A' } };
    c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    c1.alignment = { vertical: 'middle', horizontal: 'left' };
    c1.border = thinBorder;

    const cD = r.getCell(4);
    cD.value = dow.count;
    cD.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF0284C7' } };
    cD.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cD.alignment = { vertical: 'middle', horizontal: 'center' };
    cD.border = thinBorder;

    const cE = r.getCell(5);
    cE.value = `${dow.percent}%`;
    cE.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF334155' } };
    cE.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cE.alignment = { vertical: 'middle', horizontal: 'center' };
    cE.border = thinBorder;

    const barBlocks = Math.round((dow.percent / 100) * 20);
    const barStr = '█'.repeat(barBlocks) + '░'.repeat(Math.max(0, 20 - barBlocks));
    const cF = r.getCell(6);
    cF.value = `${barStr} (${dow.percent}%)`;
    cF.font = { name: 'Courier New', size: 8.5, bold: true, color: { argb: 'FF0284C7' } };
    cF.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cF.alignment = { vertical: 'middle', horizontal: 'left' };
    cF.border = thinBorder;
  });

  // Tabela 2: Distribuição por Turnos do Dia
  curRow += 2;
  ws.mergeCells(`A${curRow}:G${curRow}`);
  const t2 = ws.getCell(`A${curRow}`);
  t2.value = '2. DISTRIBUIÇÃO DE DUPLICIDADES POR TURNO DO DIA';
  t2.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  t2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'B45309' } }; // Amber 700
  t2.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.getRow(curRow).height = 24;

  curRow++;
  ws.mergeCells(`A${curRow}:C${curRow}`);
  ws.getCell(`A${curRow}`).value = 'TURNO DO DIA';
  ws.getCell(`D${curRow}`).value = 'OCORRÊNCIAS';
  ws.getCell(`E${curRow}`).value = '% DO TOTAL';
  ws.mergeCells(`F${curRow}:G${curRow}`);
  ws.getCell(`F${curRow}`).value = 'DISTRIBUIÇÃO VISUAL';
  ['A', 'D', 'E', 'F'].forEach(cKey => {
    const c = ws.getCell(`${cKey}${curRow}`);
    c.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '92400E' } };
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    c.border = thinBorder;
  });
  ws.getRow(curRow).height = 22;

  (itStats.shiftDistribution || []).forEach((sh, idx) => {
    curRow++;
    ws.mergeCells(`A${curRow}:C${curRow}`);
    ws.mergeCells(`F${curRow}:G${curRow}`);
    const r = ws.getRow(curRow);
    r.height = 20;
    const isAlt = idx % 2 === 1;
    const bgRow = isAlt ? 'F8FAFC' : 'FFFFFF';

    const c1 = r.getCell(1);
    c1.value = sh.shiftName;
    c1.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF0F172A' } };
    c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    c1.alignment = { vertical: 'middle', horizontal: 'left' };
    c1.border = thinBorder;

    const cD = r.getCell(4);
    cD.value = sh.count;
    cD.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFB45309' } };
    cD.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cD.alignment = { vertical: 'middle', horizontal: 'center' };
    cD.border = thinBorder;

    const cE = r.getCell(5);
    cE.value = `${sh.percent}%`;
    cE.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF334155' } };
    cE.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cE.alignment = { vertical: 'middle', horizontal: 'center' };
    cE.border = thinBorder;

    const barBlocks = Math.round((sh.percent / 100) * 20);
    const barStr = '█'.repeat(barBlocks) + '░'.repeat(Math.max(0, 20 - barBlocks));
    const cF = r.getCell(6);
    cF.value = `${barStr} (${sh.percent}%)`;
    cF.font = { name: 'Courier New', size: 8.5, bold: true, color: { argb: 'FFB45309' } };
    cF.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cF.alignment = { vertical: 'middle', horizontal: 'left' };
    cF.border = thinBorder;
  });

  // Tabela 3: Beneficiários / Pacientes com Maior Recorrência (Top 10)
  curRow += 2;
  ws.mergeCells(`A${curRow}:G${curRow}`);
  const t3 = ws.getCell(`A${curRow}`);
  t3.value = '3. PACIENTES / BENEFICIÁRIOS COM MAIOR RECORRÊNCIA DE DUPLICIDADES (TOP 10)';
  t3.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  t3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '047857' } }; // Emerald 700
  t3.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.getRow(curRow).height = 24;

  curRow++;
  ws.mergeCells(`A${curRow}:C${curRow}`);
  ws.getCell(`A${curRow}`).value = 'NOME / MATRÍCULA DO BENEFICIÁRIO';
  ws.getCell(`D${curRow}`).value = 'GRUPOS DUPLICADOS';
  ws.getCell(`E${curRow}`).value = '% SHARE';
  ws.mergeCells(`F${curRow}:G${curRow}`);
  ws.getCell(`F${curRow}`).value = 'NÍVEL DE RECORRÊNCIA';
  ['A', 'D', 'E', 'F'].forEach(cKey => {
    const c = ws.getCell(`${cKey}${curRow}`);
    c.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '065F46' } };
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    c.border = thinBorder;
  });
  ws.getRow(curRow).height = 22;

  const topBeneficiaries = (itStats.topBeneficiariesList || []).slice(0, 10);
  topBeneficiaries.forEach((b, idx) => {
    curRow++;
    ws.mergeCells(`A${curRow}:C${curRow}`);
    ws.mergeCells(`F${curRow}:G${curRow}`);
    const r = ws.getRow(curRow);
    r.height = 20;
    const isAlt = idx % 2 === 1;
    const bgRow = isAlt ? 'F8FAFC' : 'FFFFFF';

    const c1 = r.getCell(1);
    c1.value = `${idx + 1}. ${b.beneficiary}`;
    c1.font = { name: 'Calibri', size: 9, bold: idx === 0, color: { argb: 'FF0F172A' } };
    c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    c1.alignment = { vertical: 'middle', horizontal: 'left' };
    c1.border = thinBorder;

    const cD = r.getCell(4);
    cD.value = b.count;
    cD.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF047857' } };
    cD.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cD.alignment = { vertical: 'middle', horizontal: 'center' };
    cD.border = thinBorder;

    const cE = r.getCell(5);
    cE.value = `${b.percent}%`;
    cE.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF334155' } };
    cE.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cE.alignment = { vertical: 'middle', horizontal: 'center' };
    cE.border = thinBorder;

    const barBlocks = Math.round((b.percent / 100) * 20);
    const barStr = '█'.repeat(barBlocks) + '░'.repeat(Math.max(0, 20 - barBlocks));
    const cF = r.getCell(6);
    cF.value = `${barStr} (${b.percent}%)`;
    cF.font = { name: 'Courier New', size: 8.5, bold: true, color: { argb: 'FF047857' } };
    cF.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cF.alignment = { vertical: 'middle', horizontal: 'left' };
    cF.border = thinBorder;
  });

  // Tabela 4: Distribuição por Situação / Status da Transação
  curRow += 2;
  ws.mergeCells(`A${curRow}:G${curRow}`);
  const t4 = ws.getCell(`A${curRow}`);
  t4.value = '4. DISTRIBUIÇÃO POR SITUAÇÃO / STATUS DA TRANSAÇÃO';
  t4.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  t4.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '4338CA' } }; // Indigo 700
  t4.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.getRow(curRow).height = 24;

  curRow++;
  ws.mergeCells(`A${curRow}:C${curRow}`);
  ws.getCell(`A${curRow}`).value = 'SITUAÇÃO / STATUS';
  ws.getCell(`D${curRow}`).value = 'OCORRÊNCIAS';
  ws.getCell(`E${curRow}`).value = '% DO TOTAL';
  ws.mergeCells(`F${curRow}:G${curRow}`);
  ws.getCell(`F${curRow}`).value = 'DISTRIBUIÇÃO VISUAL';
  ['A', 'D', 'E', 'F'].forEach(cKey => {
    const c = ws.getCell(`${cKey}${curRow}`);
    c.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '3730A3' } };
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    c.border = thinBorder;
  });
  ws.getRow(curRow).height = 22;

  (itStats.situationsDistribution || []).forEach((st, idx) => {
    curRow++;
    ws.mergeCells(`A${curRow}:C${curRow}`);
    ws.mergeCells(`F${curRow}:G${curRow}`);
    const r = ws.getRow(curRow);
    r.height = 20;
    const isAlt = idx % 2 === 1;
    const bgRow = isAlt ? 'F8FAFC' : 'FFFFFF';

    const c1 = r.getCell(1);
    c1.value = st.situation;
    c1.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF0F172A' } };
    c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    c1.alignment = { vertical: 'middle', horizontal: 'left' };
    c1.border = thinBorder;

    const cD = r.getCell(4);
    cD.value = st.count;
    cD.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF4338CA' } };
    cD.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cD.alignment = { vertical: 'middle', horizontal: 'center' };
    cD.border = thinBorder;

    const cE = r.getCell(5);
    cE.value = `${st.percent}%`;
    cE.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF334155' } };
    cE.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cE.alignment = { vertical: 'middle', horizontal: 'center' };
    cE.border = thinBorder;

    const barBlocks = Math.round((st.percent / 100) * 20);
    const barStr = '█'.repeat(barBlocks) + '░'.repeat(Math.max(0, 20 - barBlocks));
    const cF = r.getCell(6);
    cF.value = `${barStr} (${st.percent}%)`;
    cF.font = { name: 'Courier New', size: 8.5, bold: true, color: { argb: 'FF4338CA' } };
    cF.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cF.alignment = { vertical: 'middle', horizontal: 'left' };
    cF.border = thinBorder;
  });

  // Tabela 5: Concentração por Prestador
  curRow += 2;
  ws.mergeCells(`A${curRow}:G${curRow}`);
  const t5 = ws.getCell(`A${curRow}`);
  t5.value = '5. CONCENTRAÇÃO DE VOLUME POR PRESTADOR (TOP 10 MAIORES VOLUMES)';
  t5.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  t5.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0F766E' } }; // Teal 700
  t5.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.getRow(curRow).height = 24;

  curRow++;
  ws.mergeCells(`A${curRow}:C${curRow}`);
  ws.getCell(`A${curRow}`).value = 'NOME DO PRESTADOR';
  ws.getCell(`D${curRow}`).value = 'VOLUME DUPLICADO';
  ws.getCell(`E${curRow}`).value = '% SHARE';
  ws.mergeCells(`F${curRow}:G${curRow}`);
  ws.getCell(`F${curRow}`).value = 'INDICADOR VISUAL';
  ['A', 'D', 'E', 'F'].forEach(cKey => {
    const c = ws.getCell(`${cKey}${curRow}`);
    c.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0D9488' } };
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    c.border = thinBorder;
  });
  ws.getRow(curRow).height = 22;

  const topProviders = (itStats.topProvidersList || []).slice(0, 10);
  topProviders.forEach((p, idx) => {
    curRow++;
    ws.mergeCells(`A${curRow}:C${curRow}`);
    ws.mergeCells(`F${curRow}:G${curRow}`);
    const r = ws.getRow(curRow);
    r.height = 20;
    const isAlt = idx % 2 === 1;
    const bgRow = isAlt ? 'F8FAFC' : 'FFFFFF';

    const c1 = r.getCell(1);
    c1.value = `${idx + 1}. ${p.provider}`;
    c1.font = { name: 'Calibri', size: 9, bold: idx === 0, color: { argb: 'FF0F172A' } };
    c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    c1.alignment = { vertical: 'middle', horizontal: 'left' };
    c1.border = thinBorder;

    const cD = r.getCell(4);
    cD.value = p.count;
    cD.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF0F766E' } };
    cD.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cD.alignment = { vertical: 'middle', horizontal: 'center' };
    cD.border = thinBorder;

    const cE = r.getCell(5);
    cE.value = `${p.percent}%`;
    cE.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF334155' } };
    cE.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cE.alignment = { vertical: 'middle', horizontal: 'center' };
    cE.border = thinBorder;

    const barBlocks = Math.round((p.percent / 100) * 20);
    const barStr = '█'.repeat(barBlocks) + '░'.repeat(Math.max(0, 20 - barBlocks));
    const cF = r.getCell(6);
    cF.value = `${barStr} (${p.percent}%)`;
    cF.font = { name: 'Courier New', size: 8.5, bold: true, color: { argb: 'FF0F766E' } };
    cF.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cF.alignment = { vertical: 'middle', horizontal: 'left' };
    cF.border = thinBorder;
  });

  // Tabela 6: Ocorrência por Procedimento
  curRow += 2;
  ws.mergeCells(`A${curRow}:G${curRow}`);
  const t6 = ws.getCell(`A${curRow}`);
  t6.value = '6. OCORRÊNCIA FREQUENCIAL POR PROCEDIMENTO (TOP 10 MAIORES VOLUMES)';
  t6.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  t6.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '4338CA' } };
  t6.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.getRow(curRow).height = 24;

  curRow++;
  ws.mergeCells(`A${curRow}:C${curRow}`);
  ws.getCell(`A${curRow}`).value = 'NOME DO PROCEDIMENTO';
  ws.getCell(`D${curRow}`).value = 'VOLUME DUPLICADO';
  ws.getCell(`E${curRow}`).value = '% SHARE';
  ws.mergeCells(`F${curRow}:G${curRow}`);
  ws.getCell(`F${curRow}`).value = 'PROPORÇÃO FREQUENCIAL';
  ['A', 'D', 'E', 'F'].forEach(cKey => {
    const c = ws.getCell(`${cKey}${curRow}`);
    c.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '3730A3' } };
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    c.border = thinBorder;
  });
  ws.getRow(curRow).height = 22;

  const topProcedures = (itStats.topProceduresList || []).slice(0, 10);
  topProcedures.forEach((pr, idx) => {
    curRow++;
    ws.mergeCells(`A${curRow}:C${curRow}`);
    ws.mergeCells(`F${curRow}:G${curRow}`);
    const r = ws.getRow(curRow);
    r.height = 20;
    const isAlt = idx % 2 === 1;
    const bgRow = isAlt ? 'F8FAFC' : 'FFFFFF';

    const c1 = r.getCell(1);
    c1.value = `${idx + 1}. ${pr.procedure}`;
    c1.font = { name: 'Calibri', size: 9, bold: idx === 0, color: { argb: 'FF0F172A' } };
    c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    c1.alignment = { vertical: 'middle', horizontal: 'left' };
    c1.border = thinBorder;

    const cD = r.getCell(4);
    cD.value = pr.count;
    cD.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF4338CA' } };
    cD.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cD.alignment = { vertical: 'middle', horizontal: 'center' };
    cD.border = thinBorder;

    const cE = r.getCell(5);
    cE.value = `${pr.percent}%`;
    cE.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF334155' } };
    cE.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cE.alignment = { vertical: 'middle', horizontal: 'center' };
    cE.border = thinBorder;

    const barBlocks = Math.round((pr.percent / 100) * 20);
    const barStr = '█'.repeat(barBlocks) + '░'.repeat(Math.max(0, 20 - barBlocks));
    const cF = r.getCell(6);
    cF.value = `${barStr} (${pr.percent}%)`;
    cF.font = { name: 'Courier New', size: 8.5, bold: true, color: { argb: 'FF4338CA' } };
    cF.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cF.alignment = { vertical: 'middle', horizontal: 'left' };
    cF.border = thinBorder;
  });

  // Tabela 7: Comparativo por Dia do Mês (Dias 01 a 31)
  curRow += 2;
  ws.mergeCells(`A${curRow}:G${curRow}`);
  const t7 = ws.getCell(`A${curRow}`);
  t7.value = '7. COMPARATIVO DE INCIDÊNCIA POR DIA DO MÊS (PADRÃO MENSAL DIAS 01 A 31)';
  t7.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  t7.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '047857' } }; // Emerald 700
  t7.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.getRow(curRow).height = 24;

  curRow++;
  ws.mergeCells(`A${curRow}:C${curRow}`);
  ws.getCell(`A${curRow}`).value = 'DIA DO MÊS';
  ws.getCell(`D${curRow}`).value = 'QUANTIDADE';
  ws.getCell(`E${curRow}`).value = '% DO TOTAL';
  ws.mergeCells(`F${curRow}:G${curRow}`);
  ws.getCell(`F${curRow}`).value = 'DISTRIBUIÇÃO VISUAL';
  ['A', 'D', 'E', 'F'].forEach(cKey => {
    const c = ws.getCell(`${cKey}${curRow}`);
    c.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '065F46' } };
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    c.border = thinBorder;
  });
  ws.getRow(curRow).height = 22;

  (itStats.dayOfMonthDistribution || []).forEach((dom, idx) => {
    curRow++;
    ws.mergeCells(`A${curRow}:C${curRow}`);
    ws.mergeCells(`F${curRow}:G${curRow}`);
    const r = ws.getRow(curRow);
    r.height = 20;
    const isAlt = idx % 2 === 1;
    const bgRow = isAlt ? 'F8FAFC' : 'FFFFFF';

    const c1 = r.getCell(1);
    c1.value = dom.dayLabel;
    c1.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF0F172A' } };
    c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    c1.alignment = { vertical: 'middle', horizontal: 'left' };
    c1.border = thinBorder;

    const cD = r.getCell(4);
    cD.value = dom.count;
    cD.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF047857' } };
    cD.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cD.alignment = { vertical: 'middle', horizontal: 'center' };
    cD.border = thinBorder;

    const cE = r.getCell(5);
    cE.value = `${dom.percent}%`;
    cE.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF334155' } };
    cE.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cE.alignment = { vertical: 'middle', horizontal: 'center' };
    cE.border = thinBorder;

    const barBlocks = Math.round((dom.percent / 100) * 20);
    const barStr = '█'.repeat(barBlocks) + '░'.repeat(Math.max(0, 20 - barBlocks));
    const cF = r.getCell(6);
    cF.value = `${barStr} (${dom.percent}%)`;
    cF.font = { name: 'Courier New', size: 8.5, bold: true, color: { argb: 'FF047857' } };
    cF.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cF.alignment = { vertical: 'middle', horizontal: 'left' };
    cF.border = thinBorder;
  });

  // Tabela 8: Análise de Ciclo Mensal
  curRow += 2;
  ws.mergeCells(`A${curRow}:G${curRow}`);
  const t8 = ws.getCell(`A${curRow}`);
  t8.value = '8. DISTRIBUIÇÃO POR CICLO DO MÊS (INÍCIO, MEIO E FECHAMENTO)';
  t8.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  t8.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'B45309' } };
  t8.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.getRow(curRow).height = 24;

  curRow++;
  ws.mergeCells(`A${curRow}:C${curRow}`);
  ws.getCell(`A${curRow}`).value = 'ETAPA DO MÊS';
  ws.getCell(`D${curRow}`).value = 'QUANTIDADE';
  ws.getCell(`E${curRow}`).value = '% DO TOTAL';
  ws.mergeCells(`F${curRow}:G${curRow}`);
  ws.getCell(`F${curRow}`).value = 'DISTRIBUIÇÃO VISUAL';
  ['A', 'D', 'E', 'F'].forEach(cKey => {
    const c = ws.getCell(`${cKey}${curRow}`);
    c.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '92400E' } };
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    c.border = thinBorder;
  });
  ws.getRow(curRow).height = 22;

  (itStats.monthCycleDistribution || []).forEach((cy, idx) => {
    curRow++;
    ws.mergeCells(`A${curRow}:C${curRow}`);
    ws.mergeCells(`F${curRow}:G${curRow}`);
    const r = ws.getRow(curRow);
    r.height = 20;
    const isAlt = idx % 2 === 1;
    const bgRow = isAlt ? 'F8FAFC' : 'FFFFFF';

    const c1 = r.getCell(1);
    c1.value = cy.cycleName;
    c1.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF0F172A' } };
    c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    c1.alignment = { vertical: 'middle', horizontal: 'left' };
    c1.border = thinBorder;

    const cD = r.getCell(4);
    cD.value = cy.count;
    cD.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFB45309' } };
    cD.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cD.alignment = { vertical: 'middle', horizontal: 'center' };
    cD.border = thinBorder;

    const cE = r.getCell(5);
    cE.value = `${cy.percent}%`;
    cE.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF334155' } };
    cE.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cE.alignment = { vertical: 'middle', horizontal: 'center' };
    cE.border = thinBorder;

    const barBlocks = Math.round((cy.percent / 100) * 20);
    const barStr = '█'.repeat(barBlocks) + '░'.repeat(Math.max(0, 20 - barBlocks));
    const cF = r.getCell(6);
    cF.value = `${barStr} (${cy.percent}%)`;
    cF.font = { name: 'Courier New', size: 8.5, bold: true, color: { argb: 'FFB45309' } };
    cF.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cF.alignment = { vertical: 'middle', horizontal: 'left' };
    cF.border = thinBorder;
  });

  // Tabela 9: Evolução Temporal entre Meses
  curRow += 2;
  ws.mergeCells(`A${curRow}:G${curRow}`);
  const t9 = ws.getCell(`A${curRow}`);
  t9.value = '9. EVOLUÇÃO TEMPORAL ENTRE MESES (COMPARAÇÃO MENSAL)';
  t9.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  t9.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1D4ED8' } };
  t9.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.getRow(curRow).height = 24;

  curRow++;
  ws.mergeCells(`A${curRow}:C${curRow}`);
  ws.getCell(`A${curRow}`).value = 'MÊS / ANO';
  ws.getCell(`D${curRow}`).value = 'OCORRÊNCIAS';
  ws.getCell(`E${curRow}`).value = '% DO TOTAL';
  ws.mergeCells(`F${curRow}:G${curRow}`);
  ws.getCell(`F${curRow}`).value = 'DISTRIBUIÇÃO VISUAL';
  ['A', 'D', 'E', 'F'].forEach(cKey => {
    const c = ws.getCell(`${cKey}${curRow}`);
    c.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1E3A8A' } };
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    c.border = thinBorder;
  });
  ws.getRow(curRow).height = 22;

  (itStats.monthlyDistribution || []).forEach((m, idx) => {
    curRow++;
    ws.mergeCells(`A${curRow}:C${curRow}`);
    ws.mergeCells(`F${curRow}:G${curRow}`);
    const r = ws.getRow(curRow);
    r.height = 20;
    const isAlt = idx % 2 === 1;
    const bgRow = isAlt ? 'F8FAFC' : 'FFFFFF';

    const c1 = r.getCell(1);
    c1.value = m.monthLabel;
    c1.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF0F172A' } };
    c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    c1.alignment = { vertical: 'middle', horizontal: 'left' };
    c1.border = thinBorder;

    const cD = r.getCell(4);
    cD.value = m.count;
    cD.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF1D4ED8' } };
    cD.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cD.alignment = { vertical: 'middle', horizontal: 'center' };
    cD.border = thinBorder;

    const cE = r.getCell(5);
    cE.value = `${m.percent}%`;
    cE.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF334155' } };
    cE.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cE.alignment = { vertical: 'middle', horizontal: 'center' };
    cE.border = thinBorder;

    const barBlocks = Math.round((m.percent / 100) * 20);
    const barStr = '█'.repeat(barBlocks) + '░'.repeat(Math.max(0, 20 - barBlocks));
    const cF = r.getCell(6);
    cF.value = `${barStr} (${m.percent}%)`;
    cF.font = { name: 'Courier New', size: 8.5, bold: true, color: { argb: 'FF1D4ED8' } };
    cF.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cF.alignment = { vertical: 'middle', horizontal: 'left' };
    cF.border = thinBorder;
  });

  // Tabela 10: Intervalo Temporal Delta T
  curRow += 2;
  ws.mergeCells(`A${curRow}:G${curRow}`);
  const t10 = ws.getCell(`A${curRow}`);
  t10.value = '10. JANELA DE INTERVALO TEMPORAL ENTRE OCORRÊNCIAS (DELTA T)';
  t10.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  t10.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0369A1' } };
  t10.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.getRow(curRow).height = 24;

  curRow++;
  ws.mergeCells(`A${curRow}:C${curRow}`);
  ws.getCell(`A${curRow}`).value = 'FAIXA DE TEMPO';
  ws.getCell(`D${curRow}`).value = 'CASOS';
  ws.getCell(`E${curRow}`).value = '% DO TOTAL';
  ws.mergeCells(`F${curRow}:G${curRow}`);
  ws.getCell(`F${curRow}`).value = 'CATEGORIA';
  ['A', 'D', 'E', 'F'].forEach(cKey => {
    const c = ws.getCell(`${cKey}${curRow}`);
    c.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '075985' } };
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    c.border = thinBorder;
  });
  ws.getRow(curRow).height = 22;

  const dtRows = [
    { label: 'Δt = 0 min (Simultâneo / Mesmo Minuto)', count: itStats.deltaZeroCount, pct: `${itStats.instantPercent}%`, cat: 'Crítico (Bem Provável)' },
    { label: 'Δt = 1 a 5 min (Micro-intervalo)', count: itStats.delta1to5Count, pct: `${itStats.microPercent}%`, cat: 'Alto Risco' },
    { label: 'Δt = 6 a 30 min (Mesmo Turno)', count: itStats.delta6to30Count, pct: `${itStats.shortPercent}%`, cat: 'Risco Moderado' },
    { label: 'Δt > 30 min (Reemissão Tardia)', count: itStats.deltaOver30Count, pct: `${itStats.extendedPercent}%`, cat: 'Acompanhamento' }
  ];

  dtRows.forEach((dtr, idx) => {
    curRow++;
    ws.mergeCells(`A${curRow}:C${curRow}`);
    ws.mergeCells(`F${curRow}:G${curRow}`);
    const r = ws.getRow(curRow);
    r.height = 20;
    const isAlt = idx % 2 === 1;
    const bgRow = isAlt ? 'F8FAFC' : 'FFFFFF';

    const c1 = r.getCell(1);
    c1.value = dtr.label;
    c1.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF0F172A' } };
    c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    c1.alignment = { vertical: 'middle', horizontal: 'left' };
    c1.border = thinBorder;

    const cD = r.getCell(4);
    cD.value = dtr.count;
    cD.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF0369A1' } };
    cD.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cD.alignment = { vertical: 'middle', horizontal: 'center' };
    cD.border = thinBorder;

    const cE = r.getCell(5);
    cE.value = dtr.pct;
    cE.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF334155' } };
    cE.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cE.alignment = { vertical: 'middle', horizontal: 'center' };
    cE.border = thinBorder;

    const cF = r.getCell(6);
    cF.value = dtr.cat;
    cF.font = { name: 'Calibri', size: 8.5, bold: true, color: { argb: idx === 0 ? 'FF991B1B' : 'FF0369A1' } };
    cF.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: idx === 0 ? 'FEE2E2' : bgRow } };
    cF.alignment = { vertical: 'middle', horizontal: 'center' };
    cF.border = thinBorder;
  });

  // Tabela 11: Comportamento de Identificadores
  curRow += 2;
  ws.mergeCells(`A${curRow}:G${curRow}`);
  const t11 = ws.getCell(`A${curRow}`);
  t11.value = '11. COMPORTAMENTO DE IDENTIFICADORES (PROTOCOLOS E REMESSAS)';
  t11.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  t11.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0F766E' } };
  t11.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.getRow(curRow).height = 24;

  curRow++;
  ws.mergeCells(`A${curRow}:C${curRow}`);
  ws.getCell(`A${curRow}`).value = 'CENÁRIO DE IDENTIFICADOR';
  ws.getCell(`D${curRow}`).value = 'CASOS';
  ws.getCell(`E${curRow}`).value = '% DO TOTAL';
  ws.mergeCells(`F${curRow}:G${curRow}`);
  ws.getCell(`F${curRow}`).value = 'PROPORÇÃO VISUAL';
  ['A', 'D', 'E', 'F'].forEach(cKey => {
    const c = ws.getCell(`${cKey}${curRow}`);
    c.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0D9488' } };
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    c.border = thinBorder;
  });
  ws.getRow(curRow).height = 22;

  const idRows = [
    { label: 'Mesmo Número de Protocolo', count: itStats.sameProtocolCount, pct: `${itStats.sameProtocolPercent}%` },
    { label: 'Protocolos Distintos Gerados', count: itStats.differentProtocolCount, pct: `${itStats.differentProtocolPercent}%` },
    { label: 'Mesma Remessa de Faturamento', count: itStats.sameRemessaCount, pct: `${itStats.sameRemessaPercent}%` },
    { label: 'Remessas Diferentes / Lotes Distintos', count: itStats.differentRemessaCount, pct: `${itStats.differentRemessaPercent}%` }
  ];

  idRows.forEach((idr, idx) => {
    curRow++;
    ws.mergeCells(`A${curRow}:C${curRow}`);
    ws.mergeCells(`F${curRow}:G${curRow}`);
    const r = ws.getRow(curRow);
    r.height = 20;
    const isAlt = idx % 2 === 1;
    const bgRow = isAlt ? 'F8FAFC' : 'FFFFFF';

    const c1 = r.getCell(1);
    c1.value = idr.label;
    c1.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF0F172A' } };
    c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    c1.alignment = { vertical: 'middle', horizontal: 'left' };
    c1.border = thinBorder;

    const cD = r.getCell(4);
    cD.value = idr.count;
    cD.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF0F766E' } };
    cD.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cD.alignment = { vertical: 'middle', horizontal: 'center' };
    cD.border = thinBorder;

    const cE = r.getCell(5);
    cE.value = idr.pct;
    cE.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF334155' } };
    cE.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cE.alignment = { vertical: 'middle', horizontal: 'center' };
    cE.border = thinBorder;

    const numPct = parseInt(idr.pct, 10) || 0;
    const barBlocks = Math.round((numPct / 100) * 20);
    const barStr = '█'.repeat(barBlocks) + '░'.repeat(Math.max(0, 20 - barBlocks));
    const cF = r.getCell(6);
    cF.value = `${barStr} (${idr.pct})`;
    cF.font = { name: 'Courier New', size: 8.5, bold: true, color: { argb: 'FF0F766E' } };
    cF.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgRow } };
    cF.alignment = { vertical: 'middle', horizontal: 'left' };
    cF.border = thinBorder;
  });

  // Ajuste de larguras de colunas
  ws.getColumn(1).width = 28;
  ws.getColumn(2).width = 18;
  ws.getColumn(3).width = 16;
  ws.getColumn(4).width = 20;
  ws.getColumn(5).width = 16;
  ws.getColumn(6).width = 28;
  ws.getColumn(7).width = 24;
}

export async function exportFormattedITReport(options: {
  transactionCount: number;
  itStats: ITDiagnosticsStats;
  fileName?: string;
}) {
  const { transactionCount, itStats, fileName } = options;
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'AuditCheck HealthTech';

  addITDiagnosticsWorksheet(workbook, itStats, transactionCount, 'Padrões e Comparativos');

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName ? `Padroes_Estatisticos_${fileName.replace(/[^a-zA-Z0-9]/g, '_')}.xlsx` : 'Padroes_Estatisticos_Duplicidades.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.URL.revokeObjectURL(url);
}

