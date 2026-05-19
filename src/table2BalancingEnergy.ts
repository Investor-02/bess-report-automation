import type { BalancingEnergyDirectionSummary, ReportPeriod, StationId } from './state/projectReportState';

export type BalancingEnergyFileInput = {
  file: File;
  stationId: StationId;
  stationName: string;
};

export type BalancingEnergyColumnMap = {
  date: string;
  period: string;
  direction: string;
  volume: string;
  price: string;
  amount: string;
};

export type BalancingEnergyDraftStation = {
  stationId: StationId;
  stationName: string;
  sourceFileName: string;
  period: ReportPeriod;
  sheetName: string;
  headerRowNumber: number;
  firstDataRowNumber: number;
  rowsRead: number;
  purchase: BalancingEnergyDirectionSummary;
  sale: BalancingEnergyDirectionSummary;
  columns: BalancingEnergyColumnMap;
  diagnostics: {
    unknownDirections: Array<{ direction: string; rows: number; volumeMwh: number; amountWithoutVatUah: number }>;
    directions: Array<{ direction: string; rows: number; volumeMwh: number; amountWithoutVatUah: number }>;
  };
  warnings: string[];
};

export type BalancingEnergyDraftResult = {
  period: ReportPeriod;
  stations: Partial<Record<StationId, BalancingEnergyDraftStation>>;
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
    return value.text.trim();
  }

  if (typeof value === 'object' && 'result' in value) {
    return getCellText(value.result);
  }

  return String(value).trim();
}

function toNumber(value: unknown) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : NaN;
  }

  const text = getCellText(value).replace(/\s/g, '').replace(',', '.');
  if (!text) {
    return NaN;
  }

  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : NaN;
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

function normalizeText(value: string) {
  return value
    .toLowerCase()
    .replace(/[’'`ʼ]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeDirection(value: string) {
  return normalizeText(value).replace(/ґ/g, 'г');
}

function roundEnergy(value: number) {
  return Math.round((value + Number.EPSILON) * 1000000) / 1000000;
}

function roundMoney(value: number) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function averagePrice(amount: number, volume: number) {
  return volume > 0 ? roundMoney(amount / volume) : 0;
}

function summarize(volumeMwh: number, amountWithoutVatUah: number): BalancingEnergyDirectionSummary {
  const roundedAmount = roundMoney(amountWithoutVatUah);
  return {
    volumeMwh: roundEnergy(volumeMwh),
    averagePriceUahMwh: averagePrice(amountWithoutVatUah, volumeMwh),
    amountWithoutVatUah: roundedAmount,
    amountWithVatUah: roundMoney(roundedAmount * 1.2),
  };
}

function columnLetter(columnNumber: number) {
  let dividend = columnNumber;
  let columnName = '';
  while (dividend > 0) {
    const modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }
  return columnName;
}

function findHeaderRow(worksheet: import('exceljs').Worksheet) {
  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const headers = new Map<string, number>();
    row.eachCell((cell, columnNumber) => {
      const text = normalizeText(getCellText(cell.value));
      if (text) {
        headers.set(text, columnNumber);
      }
    });

    const findHeader = (patterns: RegExp[]) => {
      for (const [header, column] of headers) {
        if (patterns.some((pattern) => pattern.test(header))) {
          return column;
        }
      }
      return 0;
    };

    const dateColumn = findHeader([/^дата$/, /date/]);
    const periodColumn = findHeader([/розрахунковий період/, /settlement period/]);
    const directionColumn = findHeader([/напрямок/, /direction/]);
    const volumeColumn = findHeader([/instq/, /обсяг балансування/, /мвт/]);
    const priceColumn = findHeader([/msp/, /labeo/, /ціна балансування/, /грн\/мвт/]);
    const amountColumn = findHeader([/списання\/нарахування/, /балансуючу електричну енергію/, /amount/]);

    if (dateColumn && periodColumn && directionColumn && volumeColumn && priceColumn && amountColumn) {
      return {
        rowNumber,
        dateColumn,
        periodColumn,
        directionColumn,
        volumeColumn,
        priceColumn,
        amountColumn,
      };
    }
  }

  return null;
}

function parsePeriodFromMetadata(worksheet: import('exceljs').Worksheet) {
  for (let rowNumber = 1; rowNumber <= Math.min(worksheet.rowCount, 20); rowNumber += 1) {
    const text = `${getCellText(worksheet.getRow(rowNumber).getCell(1).value)} ${getCellText(worksheet.getRow(rowNumber).getCell(2).value)}`;
    const match = text.match(/(\d{2})\.(\d{2})\.(\d{4})\s*-\s*(\d{2})\.(\d{2})\.(\d{4})/);
    if (match) {
      return `${match[3]}-${match[2]}` as ReportPeriod;
    }
  }

  return null;
}

function stationMismatchWarning(stationId: StationId, fileName: string) {
  const normalized = fileName.toUpperCase();
  if (stationId === 'oleksandriya' && /ЗНАМ|ZNAM/.test(normalized)) {
    return `Возможно, файл "${fileName}" загружен не в слот Олександрії.`;
  }
  if (stationId === 'znamyanka' && /ОЛЕКС|OLEKS/.test(normalized)) {
    return `Возможно, файл "${fileName}" загружен не в слот Знаменки.`;
  }
  return null;
}

export async function parseBalancingEnergyFile(input: BalancingEnergyFileInput): Promise<BalancingEnergyDraftStation> {
  const ExcelJS = await import('exceljs');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await input.file.arrayBuffer());

  const candidateWorksheet = workbook.worksheets.find((worksheet) => {
    for (let rowNumber = 1; rowNumber <= Math.min(worksheet.rowCount, 15); rowNumber += 1) {
      const label = getCellText(worksheet.getRow(rowNumber).getCell(1).value);
      const value = getCellText(worksheet.getRow(rowNumber).getCell(2).value);
      if (/тип звіту/i.test(label) && /балансування/i.test(value)) {
        return true;
      }
    }
    return false;
  }) ?? workbook.worksheets[0];

  if (!candidateWorksheet) {
    throw new Error(`В файле "${input.file.name}" не найден лист балансирующей энергии.`);
  }

  const header = findHeaderRow(candidateWorksheet);
  if (!header) {
    throw new Error(`В файле "${input.file.name}" не найдены колонки Дата, Розрахунковий період, Напрямок, INSTQ/MSP/Списання.`);
  }

  const directionStats = new Map<string, { rows: number; volumeMwh: number; amountWithoutVatUah: number }>();
  let firstDate = '';
  let lastDate = '';
  let rowsRead = 0;

  for (let rowNumber = header.rowNumber + 1; rowNumber <= candidateWorksheet.rowCount; rowNumber += 1) {
    const row = candidateWorksheet.getRow(rowNumber);
    const date = parseDateCell(row.getCell(header.dateColumn).value);
    const period = getCellText(row.getCell(header.periodColumn).value);
    const direction = getCellText(row.getCell(header.directionColumn).value);
    const volume = toNumber(row.getCell(header.volumeColumn).value);
    const amount = toNumber(row.getCell(header.amountColumn).value);

    if (!date || !period || !direction || !Number.isFinite(volume) || !Number.isFinite(amount)) {
      continue;
    }

    firstDate ||= date;
    lastDate = date;
    rowsRead += 1;
    const normalizedDirection = normalizeDirection(direction);
    const current = directionStats.get(normalizedDirection) ?? { rows: 0, volumeMwh: 0, amountWithoutVatUah: 0 };
    current.rows += 1;
    current.volumeMwh += volume;
    current.amountWithoutVatUah += amount;
    directionStats.set(normalizedDirection, current);
  }

  if (rowsRead === 0) {
    throw new Error(`В файле "${input.file.name}" не найдены реальные строки периода.`);
  }

  const purchaseStats = directionStats.get('вниз') ?? { rows: 0, volumeMwh: 0, amountWithoutVatUah: 0 };
  const saleStats = directionStats.get('вгору') ?? { rows: 0, volumeMwh: 0, amountWithoutVatUah: 0 };
  const period = parsePeriodFromMetadata(candidateWorksheet) ?? firstDate.slice(0, 7) as ReportPeriod;
  const warnings: string[] = [];
  const mismatchWarning = stationMismatchWarning(input.stationId, input.file.name);
  if (mismatchWarning) {
    warnings.push(mismatchWarning);
  }
  if (rowsRead < 700) {
    warnings.push(`Файл "${input.file.name}" содержит меньше 700 строк периода.`);
  }

  const directions = [...directionStats.entries()].map(([direction, stats]) => ({
    direction,
    rows: stats.rows,
    volumeMwh: roundEnergy(stats.volumeMwh),
    amountWithoutVatUah: roundMoney(stats.amountWithoutVatUah),
  }));
  const unknownDirections = directions.filter((direction) => direction.direction !== 'вниз' && direction.direction !== 'вгору');
  if (unknownDirections.length > 0) {
    warnings.push(`Найдены неизвестные направления: ${unknownDirections.map((direction) => direction.direction).join(', ')}.`);
  }

  return {
    stationId: input.stationId,
    stationName: input.stationName,
    sourceFileName: input.file.name,
    period,
    sheetName: candidateWorksheet.name,
    headerRowNumber: header.rowNumber,
    firstDataRowNumber: header.rowNumber + 1,
    rowsRead,
    purchase: summarize(purchaseStats.volumeMwh, purchaseStats.amountWithoutVatUah),
    sale: summarize(saleStats.volumeMwh, saleStats.amountWithoutVatUah),
    columns: {
      date: columnLetter(header.dateColumn),
      period: columnLetter(header.periodColumn),
      direction: columnLetter(header.directionColumn),
      volume: columnLetter(header.volumeColumn),
      price: columnLetter(header.priceColumn),
      amount: columnLetter(header.amountColumn),
    },
    diagnostics: {
      unknownDirections,
      directions,
    },
    warnings,
  };
}

export function buildBalancingEnergyDraft(files: BalancingEnergyDraftStation[]): BalancingEnergyDraftResult {
  const stations: Partial<Record<StationId, BalancingEnergyDraftStation>> = {};
  for (const file of files) {
    stations[file.stationId] = file;
  }

  const periods = [...new Set(files.map((file) => file.period))];
  const period = periods[0] ?? '2026-04';
  const warnings = files.flatMap((file) => file.warnings);
  if (periods.length > 1) {
    warnings.push(`Периоды файлов балансирующей энергии не совпадают: ${periods.join(', ')}.`);
  }

  return {
    period,
    stations,
    warnings,
  };
}
