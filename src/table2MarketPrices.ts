import type { MarketPriceHourlyRow, ReportPeriod } from './state/projectReportState';

export type MarketPricesColumns = {
  date: string;
  hour: string;
  rdnPriceUah: string;
  positiveImbalancePriceUah: string;
  negativeImbalancePriceUah: string;
  actualImbalancePriceUah: string | null;
};

export type MarketPricesDraftResult = {
  period: ReportPeriod;
  fileName: string;
  sheetName: string;
  rows: MarketPriceHourlyRow[];
  rowsCount: number;
  firstDate: string;
  lastDate: string;
  averageRdnPriceUah: number;
  averagePositiveImbalancePriceUah: number;
  averageNegativeImbalancePriceUah: number;
  averageActualImbalancePriceUah: number | null;
  columns: MarketPricesColumns;
  warnings: string[];
};

function getCellText(value: unknown) {
  if (value === null || value === undefined) {
    return '';
  }

  if (value instanceof Date) {
    return formatDate(value);
  }

  if (typeof value === 'object' && 'text' in value && typeof value.text === 'string') {
    return value.text;
  }

  return String(value).trim();
}

function toNumber(value: unknown) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : NaN;
  }

  if (typeof value === 'string') {
    const parsed = Number(value.replace(/\s/g, '').replace(',', '.'));
    return Number.isFinite(parsed) ? parsed : NaN;
  }

  return NaN;
}

function formatDate(date: Date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function parseDateCell(value: unknown) {
  if (value instanceof Date) {
    return formatDate(value);
  }

  const text = getCellText(value);
  const isoMatch = text.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
  }

  const dotMatch = text.match(/(\d{2})\.(\d{2})\.(\d{4})/);
  if (dotMatch) {
    return `${dotMatch[3]}-${dotMatch[2]}-${dotMatch[1]}`;
  }

  return '';
}

function parsePeriodFromFileName(fileName: string): ReportPeriod | null {
  const match = fileName.match(/(\d{2})-(\d{2})\.(\d{2})\.(\d{4})/);
  if (!match) {
    return null;
  }

  const [, , , month, year] = match;
  return `${year}-${month}` as ReportPeriod;
}

function average(values: number[]) {
  if (values.length === 0) {
    return 0;
  }

  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function roundPrice(value: number) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

export async function parseMarketPricesFile(file: File): Promise<MarketPricesDraftResult> {
  const ExcelJS = await import('exceljs');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());

  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error(`В файле "${file.name}" не найден лист с ценами.`);
  }

  const headerRow = worksheet.getRow(1);
  const headers = new Map<string, number>();
  headerRow.eachCell((cell, columnNumber) => {
    headers.set(getCellText(cell.value).toLowerCase(), columnNumber);
  });

  const findHeader = (patterns: RegExp[]) => {
    for (const [header, columnNumber] of headers) {
      if (patterns.some((pattern) => pattern.test(header))) {
        return columnNumber;
      }
    }

    return 0;
  };

  const dateColumn = 1;
  const hourColumn = 2;
  const actualColumn = findHeader([/фактична ціна небалансу/i, /actual imbalance price/i]);
  const rdnColumn = findHeader([/ціна рдн/i, /dam price/i]);
  const positiveColumn = findHeader([/позитивний небаланс/i, /positive imbalance/i]);
  const negativeColumn = findHeader([/негативний небаланс/i, /negative imbalance/i]);

  if (!rdnColumn || !positiveColumn || !negativeColumn) {
    throw new Error('Не удалось найти обязательные колонки цен РДН, позитивного и негативного небаланса.');
  }

  const rows: MarketPriceHourlyRow[] = [];
  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const date = parseDateCell(row.getCell(dateColumn).value);
    const hour = getCellText(row.getCell(hourColumn).value);
    const rdnPriceUah = toNumber(row.getCell(rdnColumn).value);
    const positiveImbalancePriceUah = toNumber(row.getCell(positiveColumn).value);
    const negativeImbalancePriceUah = toNumber(row.getCell(negativeColumn).value);
    const actualImbalancePriceUah = actualColumn ? toNumber(row.getCell(actualColumn).value) : NaN;

    if (!date || !hour || !Number.isFinite(rdnPriceUah) || !Number.isFinite(positiveImbalancePriceUah) || !Number.isFinite(negativeImbalancePriceUah)) {
      continue;
    }

    rows.push({
      date,
      hour,
      rdnPriceUah,
      positiveImbalancePriceUah,
      negativeImbalancePriceUah,
      actualImbalancePriceUah: Number.isFinite(actualImbalancePriceUah) ? actualImbalancePriceUah : null,
    });
  }

  if (rows.length === 0) {
    throw new Error('В файле цен не найдены почасовые строки.');
  }

  const period = parsePeriodFromFileName(file.name) ?? rows[0].date.slice(0, 7) as ReportPeriod;
  const rowsInPeriod = rows.filter((row) => row.date.startsWith(period));
  if (rowsInPeriod.length === 0) {
    throw new Error(`В файле цен не найдены строки за период ${period}.`);
  }

  const actualValues = rowsInPeriod
    .map((row) => row.actualImbalancePriceUah)
    .filter((value): value is number => typeof value === 'number' && Number.isFinite(value));

  return {
    period,
    fileName: file.name,
    sheetName: worksheet.name,
    rows: rowsInPeriod,
    rowsCount: rowsInPeriod.length,
    firstDate: rowsInPeriod[0].date,
    lastDate: rowsInPeriod[rowsInPeriod.length - 1].date,
    averageRdnPriceUah: roundPrice(average(rowsInPeriod.map((row) => row.rdnPriceUah))),
    averagePositiveImbalancePriceUah: roundPrice(average(rowsInPeriod.map((row) => row.positiveImbalancePriceUah))),
    averageNegativeImbalancePriceUah: roundPrice(average(rowsInPeriod.map((row) => row.negativeImbalancePriceUah))),
    averageActualImbalancePriceUah: actualValues.length > 0 ? roundPrice(average(actualValues)) : null,
    columns: {
      date: 'A',
      hour: 'B',
      rdnPriceUah: headerRow.getCell(rdnColumn).text,
      positiveImbalancePriceUah: headerRow.getCell(positiveColumn).text,
      negativeImbalancePriceUah: headerRow.getCell(negativeColumn).text,
      actualImbalancePriceUah: actualColumn ? headerRow.getCell(actualColumn).text : null,
    },
    warnings: rowsInPeriod.length < 700 ? ['Похоже, файл цен содержит неполный месяц.'] : [],
  };
}
