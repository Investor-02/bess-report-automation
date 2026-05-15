import * as XLSX from 'xlsx';

export type FcrMonitoringDebug = {
  totalCellsRead: number;
  addressedCellsRead: number;
  trueFound: number;
  falseFound: number;
  totalFound: number;
  readRange: string;
  trueFalseRange: string;
  dateColumnsWithValues: number;
  firstDateHeader: string;
  lastDateHeader: string;
};

export type FcrMonitoringResult = {
  trueHours: number;
  falseHours: number;
  totalHours: number;
  debug: FcrMonitoringDebug;
};

function getExplicitBooleanValue(cellValue: unknown): boolean | null {
  if (typeof cellValue === 'boolean') {
    return cellValue;
  }

  if (typeof cellValue === 'string') {
    const normalizedValue = cellValue.trim().toUpperCase();
    if (normalizedValue === 'TRUE') {
      return true;
    }
    if (normalizedValue === 'FALSE') {
      return false;
    }
  }

  return null;
}

function formatExcelDate(cellValue: unknown): string {
  if (typeof cellValue !== 'number') {
    return '';
  }

  const parsedDate = XLSX.SSF.parse_date_code(cellValue);
  if (!parsedDate) {
    return '';
  }

  return `${parsedDate.y}-${String(parsedDate.m).padStart(2, '0')}-${String(parsedDate.d).padStart(2, '0')}`;
}

export function calculateFcrMonitoringFromWorkbook(workbook: XLSX.WorkBook): FcrMonitoringResult {
  const sheet = workbook.Sheets.FCR;

  if (!sheet) {
    throw new Error('Лист "FCR" не найден в выбранном Excel-файле.');
  }

  const readRange = sheet['!ref'];
  if (!readRange) {
    throw new Error('Лист "FCR" пустой.');
  }

  const range = XLSX.utils.decode_range(readRange);
  const totalCellsRead = (range.e.r - range.s.r + 1) * (range.e.c - range.s.c + 1);
  let addressedCellsRead = 0;
  let trueHours = 0;
  let falseHours = 0;
  let minBooleanRow = Number.POSITIVE_INFINITY;
  let minBooleanColumn = Number.POSITIVE_INFINITY;
  let maxBooleanRow = 0;
  let maxBooleanColumn = 0;
  const booleanColumns = new Set<number>();

  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (let column = range.s.c; column <= range.e.c; column += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: column });
      const cell = sheet[address];
      if (cell) {
        addressedCellsRead += 1;
      }

      const booleanValue = getExplicitBooleanValue(cell?.v);
      if (booleanValue === null) {
        continue;
      }

      minBooleanRow = Math.min(minBooleanRow, row);
      minBooleanColumn = Math.min(minBooleanColumn, column);
      maxBooleanRow = Math.max(maxBooleanRow, row);
      maxBooleanColumn = Math.max(maxBooleanColumn, column);
      booleanColumns.add(column);

      if (booleanValue) {
        trueHours += 1;
      } else {
        falseHours += 1;
      }
    }
  }

  const totalHours = trueHours + falseHours;

  if (totalHours === 0) {
    throw new Error('На листе "FCR" не найдены значения TRUE/FALSE для подсчета часов.');
  }

  const sortedBooleanColumns = [...booleanColumns].sort((a, b) => a - b);
  const firstDateHeader = formatExcelDate(sheet[XLSX.utils.encode_cell({ r: 0, c: sortedBooleanColumns[0] })]?.v);
  const lastDateHeader = formatExcelDate(
    sheet[XLSX.utils.encode_cell({ r: 0, c: sortedBooleanColumns[sortedBooleanColumns.length - 1] })]?.v,
  );

  return {
    trueHours,
    falseHours,
    totalHours,
    debug: {
      totalCellsRead,
      addressedCellsRead,
      trueFound: trueHours,
      falseFound: falseHours,
      totalFound: totalHours,
      readRange,
      trueFalseRange: XLSX.utils.encode_range({
        s: { r: minBooleanRow, c: minBooleanColumn },
        e: { r: maxBooleanRow, c: maxBooleanColumn },
      }),
      dateColumnsWithValues: booleanColumns.size,
      firstDateHeader,
      lastDateHeader,
    },
  };
}

export async function calculateFcrMonitoringFromFile(file: File): Promise<FcrMonitoringResult> {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', sheetStubs: true });
  return calculateFcrMonitoringFromWorkbook(workbook);
}
