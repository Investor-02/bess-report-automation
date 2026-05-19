import type { ReportPeriod, RdnVdrHourlyRow, RdnVdrMarketResult, StationId } from './state/projectReportState';

export type RdnVdrMarket = 'РДН' | 'ВДР';

export type RdnVdrFileInput = {
  file: File;
  stationId: StationId;
  stationName: string;
  market: RdnVdrMarket;
};

export type RdnVdrParsedFile = RdnVdrMarketResult & {
  stationId: StationId;
  stationName: string;
  period: ReportPeriod;
  fileName: string;
  hourlyRows: RdnVdrHourlyRow[];
  warnings: string[];
};

export type RdnVdrStationDraft = {
  stationId: StationId;
  stationName: string;
  rdn: RdnVdrParsedFile;
  vdr: RdnVdrParsedFile;
  totalTradingResultUah: number;
};

export type RdnVdrDraftResult = {
  period: ReportPeriod;
  stations: Record<StationId, RdnVdrStationDraft>;
  summary: {
    rdnPurchaseVolumeMwh: number;
    rdnPurchaseAmountUah: number;
    rdnSaleVolumeMwh: number;
    rdnSaleAmountUah: number;
    vdrPurchaseVolumeMwh: number;
    vdrPurchaseAmountUah: number;
    vdrSaleVolumeMwh: number;
    vdrSaleAmountUah: number;
    totalTradingResultUah: number;
  };
  warnings: string[];
  calculatedAt: string;
};

const stationFileMarkers: Record<StationId, string[]> = {
  oleksandriya: ['ОЛЕКСАНДР', 'OLEKSANDR'],
  znamyanka: ['ЗНАМЕН', 'ZNAMEN', 'ZNAMYANKA'],
};

function parseNumber(value: unknown) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }

  if (value === null || value === undefined || value === '') {
    return 0;
  }

  return Number(String(value).replace(/\s/g, '').replace(',', '.')) || 0;
}

function round(value: number, digits = 2) {
  const factor = 10 ** digits;
  return Math.round((value + Number.EPSILON) * factor) / factor;
}

function normalizeHeader(value: unknown) {
  return String(value ?? '').replace(/\s+/g, ' ').trim().toLowerCase();
}

function normalizeDate(value: unknown) {
  const text = String(value ?? '').trim();
  const slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (slashMatch) {
    const [, month, day, year] = slashMatch;
    const fullYear = year.length === 2 ? `20${year}` : year;
    return `${fullYear}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
  }

  const dotMatch = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  if (dotMatch) {
    const [, day, month, year] = dotMatch;
    return `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
  }

  const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
  }

  return '';
}

function normalizeHour(value: unknown) {
  const text = String(value ?? '').trim();
  const numericHour = Number(text);
  if (Number.isFinite(numericHour)) {
    const startHour = Math.max(0, Math.min(23, Math.trunc(numericHour)));
    return String(startHour).padStart(2, '0');
  }

  const rangeMatch = text.match(/(\d{1,2})\s*:\s*\d{2}/);
  if (rangeMatch) {
    return rangeMatch[1].padStart(2, '0');
  }

  return '';
}

function findColumn(headers: unknown[], predicates: string[]) {
  const index = headers.findIndex((header) => {
    const normalizedHeader = normalizeHeader(header);
    return predicates.every((predicate) => normalizedHeader.includes(predicate));
  });

  if (index < 0) {
    throw new Error(`Не найдена колонка: ${predicates.join(' + ')}`);
  }

  return index;
}

function getPeriodFromFileName(fileName: string) {
  const match = fileName.match(/(\d{4})_(\d{2})_\d{2}_\d{4}_\d{2}_\d{2}/);
  if (!match) {
    throw new Error(`Не получилось определить период из имени файла "${fileName}".`);
  }

  return `${match[1]}-${match[2]}` as ReportPeriod;
}

function getFileWarnings(fileName: string, stationId: StationId, market: RdnVdrMarket, period: ReportPeriod) {
  const warnings: string[] = [];
  const normalizedFileName = fileName.toUpperCase();
  const stationMarkers = stationFileMarkers[stationId];
  const otherStationId: StationId = stationId === 'oleksandriya' ? 'znamyanka' : 'oleksandriya';
  const hasOwnMarker = stationMarkers.some((marker) => normalizedFileName.includes(marker));
  const hasOtherMarker = stationFileMarkers[otherStationId].some((marker) => normalizedFileName.includes(marker));

  if (hasOtherMarker && !hasOwnMarker) {
    warnings.push('Возможно, файл не соответствует выбранной станции.');
  }
  if (!normalizedFileName.includes(market)) {
    warnings.push(`Возможно, файл не соответствует рынку ${market}.`);
  }
  if (!normalizedFileName.includes(period.replace('-', '_'))) {
    warnings.push('Период в имени файла отличается от периода расчета.');
  }

  return warnings;
}

function calculateAverages(result: Omit<RdnVdrMarketResult, 'averagePurchasePriceUah' | 'averageSalePriceUah' | 'tradingResultUah'>): RdnVdrMarketResult {
  return {
    ...result,
    purchaseVolumeMwh: round(result.purchaseVolumeMwh, 6),
    purchaseAmountUah: round(result.purchaseAmountUah),
    saleVolumeMwh: round(result.saleVolumeMwh, 6),
    saleAmountUah: round(result.saleAmountUah),
    averagePurchasePriceUah: result.purchaseVolumeMwh > 0 ? round(result.purchaseAmountUah / result.purchaseVolumeMwh) : 0,
    averageSalePriceUah: result.saleVolumeMwh > 0 ? round(result.saleAmountUah / result.saleVolumeMwh) : 0,
    tradingResultUah: round(result.saleAmountUah - result.purchaseAmountUah),
  };
}

function calculateRowPrices(row: Omit<RdnVdrHourlyRow, 'purchasePriceUahMwh' | 'salePriceUahMwh'>): RdnVdrHourlyRow {
  return {
    ...row,
    purchasePriceUahMwh: row.purchaseVolumeMwh > 0 ? round(row.purchaseAmountUah / row.purchaseVolumeMwh) : 0,
    salePriceUahMwh: row.saleVolumeMwh > 0 ? round(row.saleAmountUah / row.saleVolumeMwh) : 0,
  };
}

function parseRdnRows(rows: unknown[][]): RdnVdrMarketResult & { hourlyRows: RdnVdrHourlyRow[] } {
  const headers = rows[0] ?? [];
  const dateColumn = findColumn(headers, ['доба']);
  const hourColumn = findColumn(headers, ['розрахунковий', 'період']);
  const purchaseVolumeColumn = findColumn(headers, ['обсяг', 'куп']);
  const purchaseAmountColumn = findColumn(headers, ['варт', 'куп']);
  const saleVolumeColumn = findColumn(headers, ['обсяг', 'прод']);
  const saleAmountColumn = findColumn(headers, ['варт', 'прод']);
  const hourlyRows: RdnVdrHourlyRow[] = [];
  const result = {
    market: 'РДН' as const,
    purchaseVolumeMwh: 0,
    purchaseAmountUah: 0,
    saleVolumeMwh: 0,
    saleAmountUah: 0,
    rowsRead: 0,
  };

  for (const row of rows.slice(1)) {
    if (!row.some((cell) => cell !== null && cell !== undefined && cell !== '')) {
      continue;
    }

    result.rowsRead += 1;
    const hourlyRow = calculateRowPrices({
      date: normalizeDate(row[dateColumn]),
      hour: normalizeHour(row[hourColumn]),
      market: 'РДН',
      purchaseVolumeMwh: parseNumber(row[purchaseVolumeColumn]),
      purchaseAmountUah: parseNumber(row[purchaseAmountColumn]),
      saleVolumeMwh: parseNumber(row[saleVolumeColumn]),
      saleAmountUah: parseNumber(row[saleAmountColumn]),
    });
    result.purchaseVolumeMwh += hourlyRow.purchaseVolumeMwh;
    result.purchaseAmountUah += hourlyRow.purchaseAmountUah;
    result.saleVolumeMwh += hourlyRow.saleVolumeMwh;
    result.saleAmountUah += hourlyRow.saleAmountUah;
    if (hourlyRow.date && hourlyRow.hour) {
      hourlyRows.push(hourlyRow);
    }
  }

  return {
    ...calculateAverages(result),
    hourlyRows,
  };
}

function parseVdrRows(rows: unknown[][]): RdnVdrMarketResult & { hourlyRows: RdnVdrHourlyRow[] } {
  const headers = rows[0] ?? [];
  const operationColumn = findColumn(headers, ['вид', 'опера']);
  const dateColumn = findColumn(headers, ['доба']);
  const hourColumn = findColumn(headers, ['розрахунковий', 'період']);
  const volumeColumn = findColumn(headers, ['обсяг']);
  const amountColumn = findColumn(headers, ['варт']);
  const priceColumn = findColumn(headers, ['ціна']);
  const hourlyRows: RdnVdrHourlyRow[] = [];
  const result = {
    market: 'ВДР' as const,
    purchaseVolumeMwh: 0,
    purchaseAmountUah: 0,
    saleVolumeMwh: 0,
    saleAmountUah: 0,
    rowsRead: 0,
  };

  for (const row of rows.slice(1)) {
    if (!row.some((cell) => cell !== null && cell !== undefined && cell !== '')) {
      continue;
    }

    const operation = normalizeHeader(row[operationColumn]);
    const volumeMwh = parseNumber(row[volumeColumn]);
    const amountUah = parseNumber(row[amountColumn]);
    const priceUahMwh = parseNumber(row[priceColumn]);
    const hourlyRow: RdnVdrHourlyRow = {
      date: normalizeDate(row[dateColumn]),
      hour: normalizeHour(row[hourColumn]),
      market: 'ВДР',
      purchaseVolumeMwh: 0,
      purchaseAmountUah: 0,
      saleVolumeMwh: 0,
      saleAmountUah: 0,
      purchasePriceUahMwh: 0,
      salePriceUahMwh: 0,
    };
    result.rowsRead += 1;
    if (operation.includes('куп')) {
      result.purchaseVolumeMwh += volumeMwh;
      result.purchaseAmountUah += amountUah;
      hourlyRow.purchaseVolumeMwh = volumeMwh;
      hourlyRow.purchaseAmountUah = amountUah;
      hourlyRow.purchasePriceUahMwh = priceUahMwh || (volumeMwh > 0 ? round(amountUah / volumeMwh) : 0);
    }
    if (operation.includes('прод')) {
      result.saleVolumeMwh += volumeMwh;
      result.saleAmountUah += amountUah;
      hourlyRow.saleVolumeMwh = volumeMwh;
      hourlyRow.saleAmountUah = amountUah;
      hourlyRow.salePriceUahMwh = priceUahMwh || (volumeMwh > 0 ? round(amountUah / volumeMwh) : 0);
    }
    if (hourlyRow.date && hourlyRow.hour && (hourlyRow.purchaseVolumeMwh > 0 || hourlyRow.saleVolumeMwh > 0)) {
      hourlyRows.push(hourlyRow);
    }
  }

  return {
    ...calculateAverages(result),
    hourlyRows,
  };
}

export async function parseRdnVdrFile(input: RdnVdrFileInput): Promise<RdnVdrParsedFile> {
  const XLSX = await import('xlsx');
  const workbook = XLSX.read(await input.file.arrayBuffer(), { cellDates: true });
  const sheetName = input.market === 'РДН' ? 'Купівля продаж' : 'Щоденна деталізація ВДР';
  const worksheet = workbook.Sheets[sheetName] ?? workbook.Sheets[workbook.SheetNames[0]];
  if (!worksheet) {
    throw new Error(`В файле "${input.file.name}" не найден лист для ${input.market}.`);
  }

  const rows = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1, defval: null, raw: false });
  const period = getPeriodFromFileName(input.file.name);
  const result = input.market === 'РДН' ? parseRdnRows(rows) : parseVdrRows(rows);

  return {
    ...result,
    stationId: input.stationId,
    stationName: input.stationName,
    period,
    fileName: input.file.name,
    hourlyRows: result.hourlyRows,
    warnings: getFileWarnings(input.file.name, input.stationId, input.market, period),
  };
}

export function buildRdnVdrDraft(files: RdnVdrParsedFile[]): RdnVdrDraftResult {
  const warnings = files.flatMap((file) => file.warnings.map((warning) => `${file.fileName}: ${warning}`));
  const periods = [...new Set(files.map((file) => file.period))].sort();
  const period = periods.at(-1);
  if (!period) {
    throw new Error('Нет данных РДН/ВДР для расчета.');
  }
  if (periods.length > 1) {
    warnings.push(`Файлы содержат разные периоды: ${periods.join(', ')}.`);
  }

  const getFile = (stationId: StationId, market: RdnVdrMarket) => {
    const file = files.find((item) => item.stationId === stationId && item.market === market);
    if (!file) {
      throw new Error(`Не загружен файл ${market} для станции ${stationId}.`);
    }
    return file;
  };
  const buildStation = (stationId: StationId, stationName: string): RdnVdrStationDraft => {
    const rdn = getFile(stationId, 'РДН');
    const vdr = getFile(stationId, 'ВДР');
    return {
      stationId,
      stationName,
      rdn,
      vdr,
      totalTradingResultUah: round(rdn.tradingResultUah + vdr.tradingResultUah),
    };
  };
  const stations = {
    oleksandriya: buildStation('oleksandriya', 'Олександрійська БЕСС'),
    znamyanka: buildStation('znamyanka', 'Знаменська БЕСС'),
  };

  return {
    period,
    stations,
    summary: {
      rdnPurchaseVolumeMwh: round(stations.oleksandriya.rdn.purchaseVolumeMwh + stations.znamyanka.rdn.purchaseVolumeMwh, 6),
      rdnPurchaseAmountUah: round(stations.oleksandriya.rdn.purchaseAmountUah + stations.znamyanka.rdn.purchaseAmountUah),
      rdnSaleVolumeMwh: round(stations.oleksandriya.rdn.saleVolumeMwh + stations.znamyanka.rdn.saleVolumeMwh, 6),
      rdnSaleAmountUah: round(stations.oleksandriya.rdn.saleAmountUah + stations.znamyanka.rdn.saleAmountUah),
      vdrPurchaseVolumeMwh: round(stations.oleksandriya.vdr.purchaseVolumeMwh + stations.znamyanka.vdr.purchaseVolumeMwh, 6),
      vdrPurchaseAmountUah: round(stations.oleksandriya.vdr.purchaseAmountUah + stations.znamyanka.vdr.purchaseAmountUah),
      vdrSaleVolumeMwh: round(stations.oleksandriya.vdr.saleVolumeMwh + stations.znamyanka.vdr.saleVolumeMwh, 6),
      vdrSaleAmountUah: round(stations.oleksandriya.vdr.saleAmountUah + stations.znamyanka.vdr.saleAmountUah),
      totalTradingResultUah: round(stations.oleksandriya.totalTradingResultUah + stations.znamyanka.totalTradingResultUah),
    },
    warnings,
    calculatedAt: new Date().toISOString(),
  };
}
