import type {
  DataHubHourlyRow,
  ImbalanceStationResult,
  MarketPriceHourlyRow,
  ProjectReportState,
  RdnVdrHourlyRow,
  ReportPeriod,
  StationId,
} from './state/projectReportState';

const stationIds: StationId[] = ['oleksandriya', 'znamyanka'];

const stationNames: Record<StationId, string> = {
  oleksandriya: 'Олександрійська БЕСС',
  znamyanka: 'Знаменська БЕСС',
};

export type ImbalancesDraftResult = {
  period: ReportPeriod;
  calculationMode: 'hourly' | 'monthlyApprox';
  hourlyRowsUsed: number;
  missingHours: number;
  stations: Partial<Record<StationId, ImbalanceStationResult>>;
  summary: {
    negativeImbalanceVolumeMwh: number;
    negativeImbalanceCostUah: number;
    positiveImbalanceVolumeMwh: number;
    positiveImbalanceCostUah: number;
    netImbalanceResultUah: number;
    hourlyRowsUsed: number;
    missingHours: number;
  };
  warnings: string[];
  calculatedAt: string;
};

export type HourlyImbalanceDetailRow = {
  stationId: StationId;
  stationName: string;
  date: string;
  hour: string;
  dataHubInMwh: number;
  dataHubOutMwh: number;
  purchaseVolumeMwh: number;
  saleVolumeMwh: number;
  negativeImbalanceVolumeMwh: number;
  positiveImbalanceVolumeMwh: number;
  rdnPriceUah: number;
  negativeImbalancePriceUah: number;
  positiveImbalancePriceUah: number;
  negativePriceUsedUah: number;
  positivePriceUsedUah: number;
  negativeCostUah: number;
  positiveCostUah: number;
  netResultUah: number;
};

export type HourlyImbalanceStationFilter = StationId | 'all';

type TradingHour = {
  purchaseVolumeMwh: number;
  saleVolumeMwh: number;
};

type HourlyCalculationAccumulator = {
  negativeVolume: number;
  negativeCost: number;
  negativeWeightedPrice: number;
  positiveVolume: number;
  positiveCost: number;
  positiveWeightedPrice: number;
  rowsUsed: number;
  missingHours: number;
};

function roundEnergy(value: number) {
  return Math.round((value + Number.EPSILON) * 1000) / 1000;
}

function roundMoney(value: number) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function average(values: number[]) {
  if (values.length === 0) {
    return 0;
  }

  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function emptySummary(): ImbalancesDraftResult['summary'] {
  return {
    negativeImbalanceVolumeMwh: 0,
    negativeImbalanceCostUah: 0,
    positiveImbalanceVolumeMwh: 0,
    positiveImbalanceCostUah: 0,
    netImbalanceResultUah: 0,
    hourlyRowsUsed: 0,
    missingHours: 0,
  };
}

function normalizeHour(value: string) {
  const match = String(value ?? '').match(/(\d{1,2})/);
  return match ? match[1].padStart(2, '0') : '';
}

function normalizeDate(value: string) {
  const text = String(value ?? '').trim();
  const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
  }

  const dotMatch = text.match(/^(\d{2})\.(\d{2})\.(\d{4})/);
  if (dotMatch) {
    return `${dotMatch[3]}-${dotMatch[2]}-${dotMatch[1]}`;
  }

  return text.slice(0, 10);
}

function getHourKey(date: string, hour: string) {
  return `${normalizeDate(date)}|${normalizeHour(hour)}`;
}

function buildPriceMap(priceRows: MarketPriceHourlyRow[]) {
  const priceMap = new Map<string, MarketPriceHourlyRow>();
  for (const row of priceRows) {
    const key = getHourKey(row.date, row.hour);
    if (key.length > 4) {
      priceMap.set(key, row);
    }
  }
  return priceMap;
}

function buildTradingMap(rows: RdnVdrHourlyRow[]) {
  const tradingMap = new Map<string, TradingHour>();
  for (const row of rows) {
    const key = getHourKey(row.date, row.hour);
    if (key.length <= 4) {
      continue;
    }

    const current = tradingMap.get(key) ?? { purchaseVolumeMwh: 0, saleVolumeMwh: 0 };
    current.purchaseVolumeMwh += row.purchaseVolumeMwh;
    current.saleVolumeMwh += row.saleVolumeMwh;
    tradingMap.set(key, current);
  }
  return tradingMap;
}

function calculateHourly(
  dataHubRows: DataHubHourlyRow[],
  rdnVdrRows: RdnVdrHourlyRow[],
  priceRows: MarketPriceHourlyRow[],
) {
  const priceMap = buildPriceMap(priceRows);
  const tradingMap = buildTradingMap(rdnVdrRows);
  const accumulator: HourlyCalculationAccumulator = {
    negativeVolume: 0,
    negativeCost: 0,
    negativeWeightedPrice: 0,
    positiveVolume: 0,
    positiveCost: 0,
    positiveWeightedPrice: 0,
    rowsUsed: 0,
    missingHours: 0,
  };

  for (const dataHubRow of dataHubRows) {
    const key = getHourKey(dataHubRow.date, dataHubRow.hour);
    const prices = priceMap.get(key);
    if (!prices) {
      accumulator.missingHours += 1;
      continue;
    }

    const trading = tradingMap.get(key) ?? { purchaseVolumeMwh: 0, saleVolumeMwh: 0 };
    const negativeVolume = Math.max(dataHubRow.outMwh - trading.purchaseVolumeMwh, 0);
    const positiveVolume = Math.max(dataHubRow.inMwh - trading.saleVolumeMwh, 0);
    const negativePriceUsed = Math.max(prices.negativeImbalancePriceUah, prices.rdnPriceUah * 1.05);
    const positivePriceUsed = Math.min(prices.positiveImbalancePriceUah, prices.rdnPriceUah * 0.95);
    const negativeCost = negativeVolume * negativePriceUsed;
    const positiveCost = positiveVolume * positivePriceUsed;

    accumulator.negativeVolume += negativeVolume;
    accumulator.negativeCost += negativeCost;
    accumulator.negativeWeightedPrice += negativeVolume * negativePriceUsed;
    accumulator.positiveVolume += positiveVolume;
    accumulator.positiveCost += positiveCost;
    accumulator.positiveWeightedPrice += positiveVolume * positivePriceUsed;
    accumulator.rowsUsed += 1;
  }

  return accumulator;
}

function buildStationHourlyDetailRows(input: {
  stationId: StationId;
  stationName: string;
  dataHubRows: DataHubHourlyRow[];
  rdnVdrRows: RdnVdrHourlyRow[];
  priceRows: MarketPriceHourlyRow[];
}) {
  const priceMap = buildPriceMap(input.priceRows);
  const tradingMap = buildTradingMap(input.rdnVdrRows);
  const detailRows: HourlyImbalanceDetailRow[] = [];

  for (const dataHubRow of input.dataHubRows) {
    const key = getHourKey(dataHubRow.date, dataHubRow.hour);
    const prices = priceMap.get(key);
    if (!prices) {
      continue;
    }

    const trading = tradingMap.get(key) ?? { purchaseVolumeMwh: 0, saleVolumeMwh: 0 };
    const negativeImbalanceVolumeMwh = Math.max(dataHubRow.outMwh - trading.purchaseVolumeMwh, 0);
    const positiveImbalanceVolumeMwh = Math.max(dataHubRow.inMwh - trading.saleVolumeMwh, 0);
    const negativePriceUsedUah = Math.max(prices.negativeImbalancePriceUah, prices.rdnPriceUah * 1.05);
    const positivePriceUsedUah = Math.min(prices.positiveImbalancePriceUah, prices.rdnPriceUah * 0.95);
    const negativeCostUah = negativeImbalanceVolumeMwh * negativePriceUsedUah;
    const positiveCostUah = positiveImbalanceVolumeMwh * positivePriceUsedUah;

    detailRows.push({
      stationId: input.stationId,
      stationName: input.stationName,
      date: dataHubRow.date,
      hour: dataHubRow.hour,
      dataHubInMwh: dataHubRow.inMwh,
      dataHubOutMwh: dataHubRow.outMwh,
      purchaseVolumeMwh: roundEnergy(trading.purchaseVolumeMwh),
      saleVolumeMwh: roundEnergy(trading.saleVolumeMwh),
      negativeImbalanceVolumeMwh: roundEnergy(negativeImbalanceVolumeMwh),
      positiveImbalanceVolumeMwh: roundEnergy(positiveImbalanceVolumeMwh),
      rdnPriceUah: roundMoney(prices.rdnPriceUah),
      negativeImbalancePriceUah: roundMoney(prices.negativeImbalancePriceUah),
      positiveImbalancePriceUah: roundMoney(prices.positiveImbalancePriceUah),
      negativePriceUsedUah: roundMoney(negativePriceUsedUah),
      positivePriceUsedUah: roundMoney(positivePriceUsedUah),
      negativeCostUah: roundMoney(negativeCostUah),
      positiveCostUah: roundMoney(positiveCostUah),
      netResultUah: roundMoney(positiveCostUah - negativeCostUah),
    });
  }

  return detailRows;
}

export function buildHourlyImbalanceDetailRows(
  state: ProjectReportState,
  period: ReportPeriod,
  stationFilter: HourlyImbalanceStationFilter,
) {
  const periodState = state.periods[period];
  if (!periodState) {
    throw new Error(`В ProjectReportState нет периода ${period}.`);
  }

  const priceRows = periodState.marketPrices?.result?.rows ?? [];
  if (priceRows.length === 0) {
    throw new Error(`Для периода ${period} нет сохраненных почасовых цен небалансов.`);
  }

  const selectedStationIds = stationFilter === 'all' ? stationIds : [stationFilter];
  const detailRows = selectedStationIds.flatMap((stationId) => {
    const stationState = periodState.stations?.[stationId];
    const rdnVdrRows = stationState?.rdnVdr?.hourlyRows ?? stationState?.rdnVdr?.parsedData?.hourlyRows ?? [];
    const dataHubRows = stationState?.datahub?.hourlyRows ?? stationState?.datahub?.parsedData?.hourlyRows ?? [];

    if (rdnVdrRows.length === 0 || dataHubRows.length === 0) {
      return [];
    }

    return buildStationHourlyDetailRows({
      stationId,
      stationName: stationState?.datahub?.result?.stationName ?? stationNames[stationId],
      dataHubRows,
      rdnVdrRows,
      priceRows,
    });
  });

  if (detailRows.length === 0) {
    throw new Error('Нет почасовой детализации. Заново сохраните RDN/VDR и DataHub в месячный отчет, чтобы появились hourlyRows.');
  }

  return detailRows.sort((first, second) => `${first.stationId}|${first.date}|${first.hour}`.localeCompare(`${second.stationId}|${second.date}|${second.hour}`));
}

function getHourlyExportFileName(period: ReportPeriod, stationFilter: HourlyImbalanceStationFilter) {
  const stationPart = stationFilter === 'all'
    ? 'all_stations'
    : stationFilter === 'oleksandriya'
      ? 'oleksandriya'
      : 'znamyanka';
  const [year, month] = period.split('-');
  return `Hourly_imbalances_${stationPart}_${month}_${year}.xlsx`;
}

export async function exportHourlyImbalancesToExcel(input: {
  rows: HourlyImbalanceDetailRow[];
  period: ReportPeriod;
  stationFilter: HourlyImbalanceStationFilter;
}) {
  if (input.rows.length === 0) {
    throw new Error('Нет строк для экспорта почасовой детализации.');
  }

  const ExcelJS = await import('exceljs');
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Hourly imbalances');
  const fileName = getHourlyExportFileName(input.period, input.stationFilter);
  worksheet.columns = [
    { header: 'station', key: 'stationName', width: 28 },
    { header: 'date', key: 'date', width: 12 },
    { header: 'hour', key: 'hour', width: 8 },
    { header: 'DataHub IN, MWh', key: 'dataHubInMwh', width: 16 },
    { header: 'DataHub OUT, MWh', key: 'dataHubOutMwh', width: 17 },
    { header: 'RDN/VDR purchase volume, MWh', key: 'purchaseVolumeMwh', width: 24 },
    { header: 'RDN/VDR sale volume, MWh', key: 'saleVolumeMwh', width: 22 },
    { header: 'negative imbalance volume, MWh', key: 'negativeImbalanceVolumeMwh', width: 26 },
    { header: 'positive imbalance volume, MWh', key: 'positiveImbalanceVolumeMwh', width: 26 },
    { header: 'RDN price', key: 'rdnPriceUah', width: 14 },
    { header: 'negative imbalance price', key: 'negativeImbalancePriceUah', width: 24 },
    { header: 'positive imbalance price', key: 'positiveImbalancePriceUah', width: 24 },
    { header: 'negative price used', key: 'negativePriceUsedUah', width: 20 },
    { header: 'positive price used', key: 'positivePriceUsedUah', width: 20 },
    { header: 'negative cost', key: 'negativeCostUah', width: 16 },
    { header: 'positive cost', key: 'positiveCostUah', width: 16 },
    { header: 'net result', key: 'netResultUah', width: 16 },
  ];
  worksheet.getRow(1).font = { bold: true };
  worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

  for (const row of input.rows) {
    worksheet.addRow(row);
  }

  const totals = input.rows.reduce(
    (accumulator, row) => ({
      dataHubInMwh: accumulator.dataHubInMwh + row.dataHubInMwh,
      dataHubOutMwh: accumulator.dataHubOutMwh + row.dataHubOutMwh,
      purchaseVolumeMwh: accumulator.purchaseVolumeMwh + row.purchaseVolumeMwh,
      saleVolumeMwh: accumulator.saleVolumeMwh + row.saleVolumeMwh,
      negativeImbalanceVolumeMwh: accumulator.negativeImbalanceVolumeMwh + row.negativeImbalanceVolumeMwh,
      positiveImbalanceVolumeMwh: accumulator.positiveImbalanceVolumeMwh + row.positiveImbalanceVolumeMwh,
      negativeCostUah: accumulator.negativeCostUah + row.negativeCostUah,
      positiveCostUah: accumulator.positiveCostUah + row.positiveCostUah,
      netResultUah: accumulator.netResultUah + row.netResultUah,
    }),
    {
      dataHubInMwh: 0,
      dataHubOutMwh: 0,
      purchaseVolumeMwh: 0,
      saleVolumeMwh: 0,
      negativeImbalanceVolumeMwh: 0,
      positiveImbalanceVolumeMwh: 0,
      negativeCostUah: 0,
      positiveCostUah: 0,
      netResultUah: 0,
    },
  );
  const totalsRow = worksheet.addRow({
    stationName: 'TOTAL',
    date: '-',
    hour: '-',
    dataHubInMwh: roundEnergy(totals.dataHubInMwh),
    dataHubOutMwh: roundEnergy(totals.dataHubOutMwh),
    purchaseVolumeMwh: roundEnergy(totals.purchaseVolumeMwh),
    saleVolumeMwh: roundEnergy(totals.saleVolumeMwh),
    negativeImbalanceVolumeMwh: roundEnergy(totals.negativeImbalanceVolumeMwh),
    positiveImbalanceVolumeMwh: roundEnergy(totals.positiveImbalanceVolumeMwh),
    rdnPriceUah: '-',
    negativeImbalancePriceUah: '-',
    positiveImbalancePriceUah: '-',
    negativePriceUsedUah: '-',
    positivePriceUsedUah: '-',
    negativeCostUah: roundMoney(totals.negativeCostUah),
    positiveCostUah: roundMoney(totals.positiveCostUah),
    netResultUah: roundMoney(totals.netResultUah),
  });
  totalsRow.getCell(2).value = '-';
  totalsRow.getCell(3).value = '-';
  totalsRow.font = { bold: true };
  totalsRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFEFF6FF' },
  };

  worksheet.views = [{ state: 'frozen', ySplit: 1 }];
  worksheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: worksheet.columns.length },
  };

  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    for (const columnNumber of [4, 5, 6, 7, 8, 9]) {
      row.getCell(columnNumber).numFmt = '0.000';
    }
    for (const columnNumber of [10, 11, 12, 13, 14, 15, 16, 17]) {
      row.getCell(columnNumber).numFmt = '0.00';
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);

  return { fileName, rowsCount: input.rows.length };
}

function buildStationResult(input: {
  period: ReportPeriod;
  stationId: StationId;
  stationName: string;
  calculationMode: 'hourly' | 'monthlyApprox';
  hourlyRowsUsed: number;
  missingHours: number;
  negativeVolume: number;
  negativeCost: number;
  averageNegativePrice: number;
  positiveVolume: number;
  positiveCost: number;
  averagePositivePrice: number;
  mmsKnessToStationMwh: number;
  mmsStationToKnessMwh: number;
  mmsBalanceMwh: number;
  source: ImbalanceStationResult['source'];
}): ImbalanceStationResult {
  const negativeVolume = roundEnergy(input.negativeVolume);
  const negativeCost = roundMoney(input.negativeCost);
  const positiveVolume = roundEnergy(input.positiveVolume);
  const positiveCost = roundMoney(input.positiveCost);
  const averageNegativePrice = roundMoney(input.averageNegativePrice);
  const averagePositivePrice = roundMoney(input.averagePositivePrice);
  const netImbalanceResultUah = roundMoney(positiveCost - negativeCost);

  return {
    period: input.period,
    stationId: input.stationId,
    stationName: input.stationName,
    calculationMode: input.calculationMode,
    hourlyRowsUsed: input.hourlyRowsUsed,
    missingHours: input.missingHours,
    negativeImbalanceVolumeMwh: negativeVolume,
    totalNegativeVolumeMwh: negativeVolume,
    negativeImbalanceCostUah: negativeCost,
    totalNegativeCostUah: negativeCost,
    averageNegativeImbalancePriceUsedUah: averageNegativePrice,
    averageNegativePriceUsed: averageNegativePrice,
    positiveImbalanceVolumeMwh: positiveVolume,
    totalPositiveVolumeMwh: positiveVolume,
    positiveImbalanceCostUah: positiveCost,
    totalPositiveCostUah: positiveCost,
    averagePositiveImbalancePriceUsedUah: averagePositivePrice,
    averagePositivePriceUsed: averagePositivePrice,
    netImbalanceResultUah,
    mmsKnessToStationMwh: input.mmsKnessToStationMwh,
    mmsStationToKnessMwh: input.mmsStationToKnessMwh,
    mmsBalanceMwh: input.mmsBalanceMwh,
    source: input.source,
  };
}

export function calculateImbalancesDraft(state: ProjectReportState, period: ReportPeriod): ImbalancesDraftResult {
  const periodState = state.periods[period];
  if (!periodState) {
    throw new Error(`В ProjectReportState нет периода ${period}.`);
  }

  const priceRows = periodState.marketPrices?.result?.rows ?? [];
  if (priceRows.length === 0) {
    throw new Error(`Для периода ${period} нет сохраненных цен небалансов. Сначала сохраните цены Укренерго в месячный отчет.`);
  }

  const fallbackAverageNegativePriceUsed = roundMoney(average(priceRows.map((row) => Math.max(row.negativeImbalancePriceUah, row.rdnPriceUah * 1.05))));
  const fallbackAveragePositivePriceUsed = roundMoney(average(priceRows.map((row) => Math.min(row.positiveImbalancePriceUah, row.rdnPriceUah * 0.95))));
  const warnings: string[] = [];
  const stations: Partial<Record<StationId, ImbalanceStationResult>> = {};
  const summary = emptySummary();
  let calculationMode: ImbalancesDraftResult['calculationMode'] = 'hourly';

  for (const stationId of stationIds) {
    const stationState = periodState.stations?.[stationId];
    const rdnVdr = stationState?.rdnVdr;
    const rdnVdrResult = rdnVdr?.result;
    const datahub = stationState?.datahub;
    const datahubResult = datahub?.result;

    if (!rdnVdrResult || !datahubResult) {
      warnings.push(`${stationNames[stationId]}: нет сохраненных данных RDN/VDR или DataHub за ${period}. Станция пропущена.`);
      continue;
    }

    const rdnVdrHourlyRows = rdnVdr?.hourlyRows ?? rdnVdr?.parsedData?.hourlyRows ?? [];
    const dataHubHourlyRows = datahub?.hourlyRows ?? datahub?.parsedData?.hourlyRows ?? [];
    const purchaseVolumeMwh = (rdnVdrResult.markets.rdn.purchaseVolumeMwh ?? 0) + (rdnVdrResult.markets.vdr.purchaseVolumeMwh ?? 0);
    const saleVolumeMwh = (rdnVdrResult.markets.rdn.saleVolumeMwh ?? 0) + (rdnVdrResult.markets.vdr.saleVolumeMwh ?? 0);
    const mms = stationState?.mms?.result;

    if (!mms) {
      warnings.push(`${stationNames[stationId]}: MMS/KNESS не сохранен за ${period}. Контрольные объемы показаны как 0.`);
    }

    let stationResult: ImbalanceStationResult;
    if (rdnVdrHourlyRows.length > 0 && dataHubHourlyRows.length > 0) {
      const hourly = calculateHourly(dataHubHourlyRows, rdnVdrHourlyRows, priceRows);
      if (hourly.missingHours > 0) {
        warnings.push(`${stationNames[stationId]}: не найдены цены или сопоставление для ${hourly.missingHours} часов.`);
      }

      stationResult = buildStationResult({
        period,
        stationId,
        stationName: datahubResult.stationName || stationNames[stationId],
        calculationMode: 'hourly',
        hourlyRowsUsed: hourly.rowsUsed,
        missingHours: hourly.missingHours,
        negativeVolume: hourly.negativeVolume,
        negativeCost: hourly.negativeCost,
        averageNegativePrice: hourly.negativeVolume > 0 ? hourly.negativeWeightedPrice / hourly.negativeVolume : 0,
        positiveVolume: hourly.positiveVolume,
        positiveCost: hourly.positiveCost,
        averagePositivePrice: hourly.positiveVolume > 0 ? hourly.positiveWeightedPrice / hourly.positiveVolume : 0,
        mmsKnessToStationMwh: mms?.knessToStationMwh ?? 0,
        mmsStationToKnessMwh: mms?.stationToKnessMwh ?? 0,
        mmsBalanceMwh: mms?.saldoMwh ?? 0,
        source: {
          dataHubInMwh: datahubResult.totalInMwh ?? 0,
          dataHubOutMwh: datahubResult.totalOutMwh ?? 0,
          purchaseVolumeMwh: roundEnergy(purchaseVolumeMwh),
          saleVolumeMwh: roundEnergy(saleVolumeMwh),
          marketPriceRowsRead: priceRows.length,
          dataHubHourlyRowsRead: dataHubHourlyRows.length,
          rdnRowsRead: (rdnVdrResult.markets.rdn.rowsRead ?? 0),
          vdrRowsRead: (rdnVdrResult.markets.vdr.rowsRead ?? 0),
        },
      });
    } else {
      calculationMode = 'monthlyApprox';
      warnings.push(
        `${stationNames[stationId]}: в ProjectReportState нет сохраненных hourlyRows RDN/VDR или DataHub. `
        + 'Это старые месячные данные. Заново загрузите файлы в разделах RDN / VDR и DataHub, нажмите расчет и сохраните их в месячный отчет.',
      );
      const negativeVolume = Math.max((datahubResult.totalOutMwh ?? 0) - purchaseVolumeMwh, 0);
      const positiveVolume = Math.max((datahubResult.totalInMwh ?? 0) - saleVolumeMwh, 0);
      stationResult = buildStationResult({
        period,
        stationId,
        stationName: datahubResult.stationName || stationNames[stationId],
        calculationMode: 'monthlyApprox',
        hourlyRowsUsed: 0,
        missingHours: priceRows.length,
        negativeVolume,
        negativeCost: negativeVolume * fallbackAverageNegativePriceUsed,
        averageNegativePrice: fallbackAverageNegativePriceUsed,
        positiveVolume,
        positiveCost: positiveVolume * fallbackAveragePositivePriceUsed,
        averagePositivePrice: fallbackAveragePositivePriceUsed,
        mmsKnessToStationMwh: mms?.knessToStationMwh ?? 0,
        mmsStationToKnessMwh: mms?.stationToKnessMwh ?? 0,
        mmsBalanceMwh: mms?.saldoMwh ?? 0,
        source: {
          dataHubInMwh: datahubResult.totalInMwh ?? 0,
          dataHubOutMwh: datahubResult.totalOutMwh ?? 0,
          purchaseVolumeMwh: roundEnergy(purchaseVolumeMwh),
          saleVolumeMwh: roundEnergy(saleVolumeMwh),
          marketPriceRowsRead: priceRows.length,
          dataHubHourlyRowsRead: datahubResult.hourlyRowsRead ?? 0,
          rdnRowsRead: (rdnVdrResult.markets.rdn.rowsRead ?? 0),
          vdrRowsRead: (rdnVdrResult.markets.vdr.rowsRead ?? 0),
        },
      });
    }

    stations[stationId] = stationResult;
    summary.negativeImbalanceVolumeMwh += stationResult.negativeImbalanceVolumeMwh;
    summary.negativeImbalanceCostUah += stationResult.negativeImbalanceCostUah;
    summary.positiveImbalanceVolumeMwh += stationResult.positiveImbalanceVolumeMwh;
    summary.positiveImbalanceCostUah += stationResult.positiveImbalanceCostUah;
    summary.netImbalanceResultUah += stationResult.netImbalanceResultUah;
    summary.hourlyRowsUsed += stationResult.hourlyRowsUsed;
    summary.missingHours += stationResult.missingHours;
  }

  if (Object.keys(stations).length === 0) {
    throw new Error(`Для периода ${period} нет сохраненных данных RDN/VDR и DataHub ни по одной станции.`);
  }

  return {
    period,
    calculationMode,
    hourlyRowsUsed: summary.hourlyRowsUsed,
    missingHours: summary.missingHours,
    stations,
    summary: {
      negativeImbalanceVolumeMwh: roundEnergy(summary.negativeImbalanceVolumeMwh),
      negativeImbalanceCostUah: roundMoney(summary.negativeImbalanceCostUah),
      positiveImbalanceVolumeMwh: roundEnergy(summary.positiveImbalanceVolumeMwh),
      positiveImbalanceCostUah: roundMoney(summary.positiveImbalanceCostUah),
      netImbalanceResultUah: roundMoney(summary.netImbalanceResultUah),
      hourlyRowsUsed: summary.hourlyRowsUsed,
      missingHours: summary.missingHours,
    },
    warnings,
    calculatedAt: new Date().toISOString(),
  };
}
