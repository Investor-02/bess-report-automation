import type { DataHubHourlyRow, ReportPeriod, StationId } from './state/projectReportState';

export type DataHubFileInput = {
  file: File;
  stationId: StationId;
  stationName: string;
};

export type DataHubParsedFile = {
  stationId: StationId;
  stationName: string;
  fileName: string;
  period: ReportPeriod;
  totalInKwh: number;
  totalOutKwh: number;
  totalInMwh: number;
  totalOutMwh: number;
  saldoMwh: number;
  hourlyRowsRead: number;
  hourlyRows: DataHubHourlyRow[];
  warnings: string[];
};

export type DataHubDraftResult = {
  period: ReportPeriod;
  stations: Record<StationId, DataHubParsedFile>;
  summary: {
    totalInMwh: number;
    totalOutMwh: number;
    saldoMwh: number;
  };
  warnings: string[];
};

const sheetName = 'Група А';

function toNumber(value: unknown) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }

  if (typeof value === 'string') {
    const normalized = value.replace(/\s/g, '').replace(',', '.');
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  return 0;
}

function getCellText(value: unknown) {
  if (value === null || value === undefined) {
    return '';
  }

  if (typeof value === 'object' && 'text' in value && typeof value.text === 'string') {
    return value.text;
  }

  return String(value).trim();
}

function roundEnergy(value: number) {
  return Math.round((value + Number.EPSILON) * 1000) / 1000;
}

function parseDateHour(value: unknown) {
  const text = getCellText(value);
  const match = text.match(/(\d{2})\.(\d{2})\.(\d{4})\s+(\d{1,2}):\d{2}/);
  if (!match) {
    return null;
  }

  return {
    date: `${match[3]}-${match[2]}-${match[1]}`,
    hour: match[4].padStart(2, '0'),
  };
}

function parsePeriodFromFileName(fileName: string): ReportPeriod | null {
  const match = fileName.match(/(\d{2})_(\d{2})_(\d{4})_(\d{2})_(\d{2})_(\d{4})/);
  if (!match) {
    return null;
  }

  const [, , month, year] = match;
  return `${year}-${month}` as ReportPeriod;
}

function getStationWarnings(fileName: string, stationId: StationId, period: ReportPeriod) {
  const warnings: string[] = [];
  const normalizedFileName = fileName.toUpperCase();

  const hasZnamyankaMarker = /ЗНАМЕН|ZNAMEN|ZNAMYANKA/.test(normalizedFileName);
  const hasOleksandriyaMarker = /ОЛЕКСАНДР|ОЛЕКСАНДРІ|OLEKSAND|OLEKSANDRI/.test(normalizedFileName);

  if (stationId === 'oleksandriya' && hasZnamyankaMarker) {
    warnings.push('Возможно, файл DataHub относится к Знаменке, а загружен в слот Олександрии.');
  }

  if (stationId === 'znamyanka' && hasOleksandriyaMarker) {
    warnings.push('Возможно, файл DataHub относится к Олександрии, а загружен в слот Знаменки.');
  }

  if (!period) {
    warnings.push('Не удалось определить период DataHub из имени файла.');
  }

  return warnings;
}

export async function parseDataHubFile(input: DataHubFileInput): Promise<DataHubParsedFile> {
  const ExcelJS = await import('exceljs');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await input.file.arrayBuffer());

  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) {
    throw new Error(`В файле "${input.file.name}" не найден лист "${sheetName}".`);
  }

  let inColumn = 0;
  let outColumn = 0;
  let directionRowNumber = 0;

  worksheet.eachRow((row, rowNumber) => {
    if (directionRowNumber) {
      return;
    }

    row.eachCell((cell, columnNumber) => {
      const text = getCellText(cell.value).toUpperCase();
      if (text === 'IN') {
        inColumn = columnNumber;
      }
      if (text === 'OUT') {
        outColumn = columnNumber;
      }
    });

    if (inColumn && outColumn) {
      directionRowNumber = rowNumber;
    }
  });

  if (!inColumn || !outColumn || !directionRowNumber) {
    throw new Error(`В листе "${sheetName}" файла "${input.file.name}" не найдены колонки IN/OUT.`);
  }

  let totalInKwh = 0;
  let totalOutKwh = 0;
  let hourlyRowsRead = 0;
  const hourlyRows: DataHubHourlyRow[] = [];

  for (let rowNumber = directionRowNumber + 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const rowTitle = getCellText(row.getCell(1).value).toUpperCase();

    if (rowNumber === directionRowNumber + 1 || !rowTitle || rowTitle.includes('СУМА ЗА ПЕРІОД')) {
      continue;
    }

    const inValue = toNumber(row.getCell(inColumn).value);
    const outValue = toNumber(row.getCell(outColumn).value);
    const dateHour = parseDateHour(row.getCell(1).value);

    if (!dateHour) {
      continue;
    }

    totalInKwh += inValue;
    totalOutKwh += outValue;
    hourlyRowsRead += 1;
    const inMwh = roundEnergy(inValue / 1000);
    const outMwh = roundEnergy(outValue / 1000);
    hourlyRows.push({
      date: dateHour.date,
      hour: dateHour.hour,
      inMwh,
      outMwh,
      balanceMwh: roundEnergy(inMwh - outMwh),
    });
  }

  if (hourlyRowsRead === 0) {
    throw new Error(`В листе "${sheetName}" файла "${input.file.name}" не найдены почасовые значения DataHub.`);
  }

  const period = parsePeriodFromFileName(input.file.name);
  if (!period) {
    throw new Error(`Не удалось определить период из имени файла "${input.file.name}".`);
  }

  const totalInMwh = roundEnergy(totalInKwh / 1000);
  const totalOutMwh = roundEnergy(totalOutKwh / 1000);

  return {
    stationId: input.stationId,
    stationName: input.stationName,
    fileName: input.file.name,
    period,
    totalInKwh,
    totalOutKwh,
    totalInMwh,
    totalOutMwh,
    saldoMwh: roundEnergy(totalInMwh - totalOutMwh),
    hourlyRowsRead,
    hourlyRows,
    warnings: getStationWarnings(input.file.name, input.stationId, period),
  };
}

export function buildDataHubDraft(files: DataHubParsedFile[]): DataHubDraftResult {
  const oleksandriya = files.find((file) => file.stationId === 'oleksandriya');
  const znamyanka = files.find((file) => file.stationId === 'znamyanka');

  if (!oleksandriya || !znamyanka) {
    throw new Error('Для расчета DataHub нужны 2 файла: Олександрія и Знаменка.');
  }

  const period = oleksandriya.period;
  const warnings = files.flatMap((file) => file.warnings);

  if (znamyanka.period !== period) {
    warnings.push(`Периоды DataHub не совпадают: Олександрія ${period}, Знаменка ${znamyanka.period}.`);
  }

  return {
    period,
    stations: {
      oleksandriya,
      znamyanka,
    },
    summary: {
      totalInMwh: roundEnergy(oleksandriya.totalInMwh + znamyanka.totalInMwh),
      totalOutMwh: roundEnergy(oleksandriya.totalOutMwh + znamyanka.totalOutMwh),
      saldoMwh: roundEnergy(oleksandriya.saldoMwh + znamyanka.saldoMwh),
    },
    warnings,
  };
}
