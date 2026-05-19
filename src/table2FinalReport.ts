import type { ProjectReportState, ReportPeriod, StationId } from './state/projectReportState';

export type FinalReportRow = {
  section: 'purchase' | 'sale' | 'saldo';
  operation: string;
  volumeMwh: number | null;
  averagePriceUahMwh: number | null;
  costWithoutVatUah: number | null;
  costWithVatUah: number | null;
  note?: string;
};

export type FinalReportStation = {
  stationId: StationId;
  stationName: string;
  readiness: {
    rdnVdr: boolean;
    datahub: boolean;
    imbalances: boolean;
    mms: boolean;
    balancingEnergy: boolean;
  };
  warnings: string[];
  rows: FinalReportRow[];
};

export type FinalReportData = {
  period: ReportPeriod;
  monthSheetName: string;
  stations: FinalReportStation[];
  warnings: string[];
};

export type FinalReportExportResult = {
  fileName: string;
  sheetName: string;
  templateSource: string;
  warnings: string[];
  foundStations: string[];
  filledRows: string[];
  missingRows: string[];
};

const stationOrder: StationId[] = ['znamyanka', 'oleksandriya'];

const stationLabels: Record<StationId, string> = {
  oleksandriya: 'ОЛЕКСАНДРІЙСЬКА БЕСС',
  znamyanka: 'ЗНАМЕНСЬКА БЕСС',
};

const monthNames: Record<string, string> = {
  '01': 'СІЧЕНЬ',
  '02': 'ЛЮТИЙ',
  '03': 'БЕРЕЗЕНЬ',
  '04': 'КВІТЕНЬ',
  '05': 'ТРАВЕНЬ',
  '06': 'ЧЕРВЕНЬ',
  '07': 'ЛИПЕНЬ',
  '08': 'СЕРПЕНЬ',
  '09': 'ВЕРЕСЕНЬ',
  '10': 'ЖОВТЕНЬ',
  '11': 'ЛИСТОПАД',
  '12': 'ГРУДЕНЬ',
};

function roundEnergy(value: number) {
  return Math.round((value + Number.EPSILON) * 1000) / 1000;
}

function roundMoney(value: number) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function withVat(value: number | null) {
  return value === null ? null : roundMoney(value * 1.2);
}

function averagePrice(cost: number, volume: number) {
  return volume > 0 ? roundMoney(cost / volume) : 0;
}

export function getTable2MonthSheetName(period: ReportPeriod) {
  const [year, month] = period.split('-');
  const monthName = monthNames[month];
  if (!monthName) {
    throw new Error(`Неизвестный месяц в периоде ${period}.`);
  }

  return `${monthName}_${year}`;
}

function sumRows(rows: FinalReportRow[]) {
  const volume = rows.reduce((sum, row) => sum + (row.volumeMwh ?? 0), 0);
  const cost = rows.reduce((sum, row) => sum + (row.costWithoutVatUah ?? 0), 0);
  return {
    volumeMwh: roundEnergy(volume),
    averagePriceUahMwh: averagePrice(cost, volume),
    costWithoutVatUah: roundMoney(cost),
    costWithVatUah: withVat(cost),
  };
}

function makeMarketRow(section: 'purchase' | 'sale', operation: string, volume: number, amount: number): FinalReportRow {
  return {
    section,
    operation,
    volumeMwh: roundEnergy(volume),
    averagePriceUahMwh: averagePrice(amount, volume),
    costWithoutVatUah: roundMoney(amount),
    costWithVatUah: withVat(amount),
  };
}

function makeEmptyCostRow(section: 'purchase' | 'sale', operation: string, volume: number, note: string): FinalReportRow {
  return {
    section,
    operation,
    volumeMwh: roundEnergy(volume),
    averagePriceUahMwh: null,
    costWithoutVatUah: null,
    costWithVatUah: null,
    note,
  };
}

function hasFormula(cell: import('exceljs').Cell) {
  const value = cell.value;
  return Boolean(value && typeof value === 'object' && ('formula' in value || 'sharedFormula' in value));
}

function buildStationRows(state: ProjectReportState, period: ReportPeriod, stationId: StationId): FinalReportStation {
  const periodState = state.periods[period];
  const stationState = periodState?.stations?.[stationId];
  const rdnVdr = stationState?.rdnVdr?.result;
  const datahub = stationState?.datahub?.result;
  const imbalances = stationState?.imbalances?.result;
  const mms = stationState?.mms?.result;
  const balancingEnergy = stationState?.balancingEnergy?.result;
  const warnings: string[] = [];

  if (!rdnVdr) warnings.push(`${stationLabels[stationId]}: отсутствуют RDN/VDR.`);
  if (!datahub) warnings.push(`${stationLabels[stationId]}: отсутствует DataHub.`);
  if (!mms) warnings.push(`${stationLabels[stationId]}: отсутствует MMS/KNESS.`);
  if (!balancingEnergy) {
    warnings.push(`${stationLabels[stationId]}: нет файла балансирующей энергии Укренерго для выбранной станции/месяца.`);
  }
  warnings.push(`${stationLabels[stationId]}: нет данных стоимости КНЕСС. Нужны акты/расчет КНЕСС.`);

  const purchaseDetails: FinalReportRow[] = [
    makeEmptyCostRow('purchase', 'постачання активної електроенергії (КНЕСС)', 0, 'Стоимость будет добавлена после актов КНЕСС.'),
    makeMarketRow('purchase', 'на ринку РДН', rdnVdr?.markets.rdn.purchaseVolumeMwh ?? 0, rdnVdr?.markets.rdn.purchaseAmountUah ?? 0),
    makeMarketRow('purchase', 'на ринку ВДР', rdnVdr?.markets.vdr.purchaseVolumeMwh ?? 0, rdnVdr?.markets.vdr.purchaseAmountUah ?? 0),
    makeMarketRow(
      'purchase',
      'купівля балансуючої електроенергії (НЕК "УКРЕНЕРГО")',
      balancingEnergy?.purchase.volumeMwh ?? 0,
      balancingEnergy?.purchase.amountWithoutVatUah ?? 0,
    ),
    makeEmptyCostRow('purchase', 'електроенергія для врегулювання небалансів (КНЕСС)', mms?.stationToKnessMwh ?? 0, 'Стоимость будет добавлена после актов КНЕСС.'),
  ];
  const saleDetails: FinalReportRow[] = [
    makeMarketRow('sale', 'на ринку РДН', rdnVdr?.markets.rdn.saleVolumeMwh ?? 0, rdnVdr?.markets.rdn.saleAmountUah ?? 0),
    makeMarketRow('sale', 'на ринку ВДР', rdnVdr?.markets.vdr.saleVolumeMwh ?? 0, rdnVdr?.markets.vdr.saleAmountUah ?? 0),
    makeMarketRow(
      'sale',
      'продаж балансуючої електроенергії (НЕК "УКРЕНЕРГО")',
      balancingEnergy?.sale.volumeMwh ?? 0,
      balancingEnergy?.sale.amountWithoutVatUah ?? 0,
    ),
    makeEmptyCostRow('sale', 'електроенергія для врегулювання небалансів (КНЕСС)', mms?.knessToStationMwh ?? 0, 'Стоимость будет добавлена после актов КНЕСС.'),
  ];
  const purchaseTotal = sumRows(purchaseDetails);
  const saleTotal = sumRows(saleDetails);
  const saldoCost = saleTotal.costWithoutVatUah - purchaseTotal.costWithoutVatUah;
  const saldoVolume = saleTotal.volumeMwh - purchaseTotal.volumeMwh;
  const rows: FinalReportRow[] = [
    {
      section: 'purchase',
      operation: 'Купівля електроенергії:',
      ...purchaseTotal,
    },
    { section: 'purchase', operation: 'у т.ч.', volumeMwh: null, averagePriceUahMwh: null, costWithoutVatUah: null, costWithVatUah: null },
    ...purchaseDetails,
    {
      section: 'sale',
      operation: 'Продаж електроенергії',
      ...saleTotal,
    },
    { section: 'sale', operation: 'у т.ч.', volumeMwh: null, averagePriceUahMwh: null, costWithoutVatUah: null, costWithVatUah: null },
    ...saleDetails,
    {
      section: 'saldo',
      operation: 'САЛЬДО (продаж- купівля):',
      volumeMwh: roundEnergy(saldoVolume),
      averagePriceUahMwh: averagePrice(saldoCost, Math.abs(saldoVolume)),
      costWithoutVatUah: roundMoney(saldoCost),
      costWithVatUah: withVat(saldoCost),
    },
  ];

  return {
    stationId,
    stationName: stationLabels[stationId],
    readiness: {
      rdnVdr: Boolean(rdnVdr),
      datahub: Boolean(datahub),
      imbalances: Boolean(imbalances),
      mms: Boolean(mms),
      balancingEnergy: Boolean(balancingEnergy),
    },
    warnings,
    rows,
  };
}

export function buildTable2FinalReportData(state: ProjectReportState | null, period: ReportPeriod): FinalReportData {
  const monthSheetName = getTable2MonthSheetName(period);
  if (!state?.periods?.[period]) {
    return {
      period,
      monthSheetName,
      stations: stationOrder.map((stationId) => buildStationRows({ version: 1, activePeriod: period, periods: { [period]: { period, stations: {} as never, marketPrices: {} as never, summary: null, lastUpdatedAt: null } }, lastUpdatedAt: null }, period, stationId)),
      warnings: [`В ProjectReportState нет периода ${period}.`],
    };
  }

  const stations = stationOrder.map((stationId) => buildStationRows(state, period, stationId));
  return {
    period,
    monthSheetName,
    stations,
    warnings: stations.flatMap((station) => station.warnings),
  };
}

function setCellValueKeepingFormula(cell: import('exceljs').Cell, value: number | string | null) {
  if (hasFormula(cell)) {
    return;
  }

  if (value === null) {
    cell.value = null;
    return;
  }

  cell.value = value;
}

function writeReportRow(worksheet: import('exceljs').Worksheet, rowNumber: number, row: FinalReportRow) {
  setCellValueKeepingFormula(worksheet.getCell(rowNumber, 2), row.volumeMwh);
  setCellValueKeepingFormula(worksheet.getCell(rowNumber, 3), row.averagePriceUahMwh);
  setCellValueKeepingFormula(worksheet.getCell(rowNumber, 4), row.costWithoutVatUah);
  setCellValueKeepingFormula(worksheet.getCell(rowNumber, 5), row.costWithVatUah);
}

function getCellText(cell: import('exceljs').Cell) {
  try {
    const value = cell.value;
    if (value === null || value === undefined) {
      return cell.text ?? '';
    }
    if (typeof value === 'object') {
      if ('richText' in value && Array.isArray(value.richText)) {
        return value.richText.map((item) => item.text ?? '').join('');
      }
      if ('formula' in value && typeof value.formula === 'string') {
        return value.result === undefined || value.result === null ? value.formula : String(value.result);
      }
      if ('text' in value && typeof value.text === 'string') {
        return value.text;
      }
    }
    return String(cell.text || value || '');
  } catch {
    return '';
  }
}

function normalizeSearchText(value: unknown) {
  return String(value ?? '')
    .toUpperCase()
    .replace(/[’'`ʼ]/g, '')
    .replace(/І/g, 'И')
    .replace(/Ї/g, 'И')
    .replace(/Є/g, 'Е')
    .replace(/[^A-ZА-Я0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function rowText(worksheet: import('exceljs').Worksheet, rowNumber: number) {
  const row = worksheet.getRow(rowNumber);
  const parts: string[] = [];
  for (let columnNumber = 1; columnNumber <= Math.max(worksheet.columnCount, 8); columnNumber += 1) {
    const text = getCellText(row.getCell(columnNumber));
    if (text) {
      parts.push(text);
    }
  }
  return parts.join(' ');
}

function isStationTitle(text: string, stationId: StationId) {
  const normalized = normalizeSearchText(text);
  if (stationId === 'znamyanka') {
    return /ЗНАМЕНСЬКА|ЗНАМЯНСЬКА|ЗНАМЯНСКАЯ|ЗНАМЕНКА|ЗНАМЯНКА|ZNAM/.test(normalized);
  }

  return /ОЛЕКСАНДРИЙСЬКА|ОЛЕКСАНДРИЯ|ОЛЕКСАНДРІЯ|ОЛЕКСАНДРИЙСКАЯ|OLEKSAND/.test(normalized);
}

function findStationTitleRows(worksheet: import('exceljs').Worksheet) {
  const result = new Map<StationId, number>();
  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const title = rowText(worksheet, rowNumber);
    if (!result.has('znamyanka') && isStationTitle(title, 'znamyanka')) {
      result.set('znamyanka', rowNumber);
    }
    if (!result.has('oleksandriya') && isStationTitle(title, 'oleksandriya')) {
      result.set('oleksandriya', rowNumber);
    }
  }
  return result;
}

function normalizeOperation(value: unknown) {
  return normalizeSearchText(value);
}

function findRowInStationBlock(
  worksheet: import('exceljs').Worksheet,
  startRow: number,
  endRow: number,
  operation: string,
) {
  const target = normalizeOperation(operation);
  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
    const label = normalizeOperation(getCellText(worksheet.getCell(rowNumber, 1)));
    if (label === target) {
      return rowNumber;
    }
  }
  return null;
}

function findWorksheetByName(workbook: import('exceljs').Workbook, sheetName: string) {
  const direct = workbook.getWorksheet(sheetName);
  if (direct) {
    return direct;
  }

  const normalizedTarget = normalizeSearchText(sheetName);
  return workbook.worksheets.find((worksheet) => normalizeSearchText(worksheet.name) === normalizedTarget) ?? null;
}

async function loadTemplateWorkbook() {
  const ExcelJS = await import('exceljs');
  const workbook = new ExcelJS.Workbook();
  const templateCandidates = [
    '/templates/table_2_final_report_template.xlsx',
    '/templates/table_2_final_report_template.xlsx.xlsx',
  ];
  let lastError: unknown = null;

  for (const templateSource of templateCandidates) {
    try {
      const response = await fetch(`${templateSource}?v=${Date.now()}`, { cache: 'no-store' });
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }
      await workbook.xlsx.load(await response.arrayBuffer());
      return { workbook, templateSource };
    } catch (error) {
      lastError = error;
    }
  }

  throw new Error(`Не удалось открыть шаблон Table_2. Проверьте файл public/templates/table_2_final_report_template.xlsx. Последняя ошибка: ${lastError instanceof Error ? lastError.message : String(lastError)}`);
}

export async function exportTable2FinalReportInBrowser(report: FinalReportData): Promise<FinalReportExportResult> {
  const { workbook, templateSource } = await loadTemplateWorkbook();
  workbook.calcProperties.fullCalcOnLoad = true;
  const worksheet = findWorksheetByName(workbook, report.monthSheetName);
  if (!worksheet) {
    throw new Error(`В шаблоне не найден лист "${report.monthSheetName}" для периода ${report.period}.`);
  }

  const titleRows = findStationTitleRows(worksheet);
  const warnings = [...report.warnings];
  const foundStations = [...titleRows.entries()].map(([stationId, rowNumber]) => `${stationLabels[stationId]}: row ${rowNumber}`);
  const filledRows: string[] = [];
  const missingRows: string[] = [];
  const missingStationBlocks = stationOrder.filter((stationId) => !titleRows.has(stationId));
  if (missingStationBlocks.length > 0) {
    throw new Error(`В листе "${worksheet.name}" не найден блок станции: ${missingStationBlocks.map((stationId) => stationLabels[stationId]).join(', ')}.`);
  }

  for (const station of report.stations) {
    const titleRow = titleRows.get(station.stationId);
    if (!titleRow) {
      warnings.push(`В листе "${report.monthSheetName}" не найден блок станции ${station.stationName}.`);
      continue;
    }

    const nextTitleRows = [...titleRows.values()].filter((rowNumber) => rowNumber > titleRow).sort((a, b) => a - b);
    const endRow = (nextTitleRows[0] ?? worksheet.rowCount + 1) - 1;
    let searchStartRow = titleRow;

    for (const row of station.rows) {
      const rowNumber = findRowInStationBlock(worksheet, searchStartRow, endRow, row.operation);
      if (!rowNumber) {
        const message = `${station.stationName}: не найдена строка "${row.operation}" в листе "${worksheet.name}".`;
        warnings.push(message);
        missingRows.push(message);
        continue;
      }
      writeReportRow(worksheet, rowNumber, row);
      filledRows.push(`${station.stationName}: row ${rowNumber} "${row.operation}"`);
      searchStartRow = rowNumber + 1;
    }
  }

  const [year, month] = report.period.split('-');
  const fileName = `Table_2_Final_Report_${year}_${month}.xlsx`;
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);

  return {
    fileName,
    sheetName: worksheet.name,
    templateSource,
    warnings,
    foundStations,
    filledRows,
    missingRows,
  };
}
