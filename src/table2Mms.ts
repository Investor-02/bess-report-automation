import type { MmsDirectionDiagnostic, MmsKnessColumnDiagnostic, ReportPeriod, StationId } from './state/projectReportState';

export type MmsFileInput = {
  file: File;
  stationId: StationId;
  stationName: string;
};

export type MmsParsedFile = {
  stationId: StationId;
  stationName: string;
  fileName: string;
  period: ReportPeriod;
  knessToStationKwh: number;
  stationToKnessKwh: number;
  knessToStationMwh: number;
  stationToKnessMwh: number;
  operatorToStationMwh: number;
  stationToOperatorMwh: number;
  naToStationMwh: number;
  stationToNaMwh: number;
  otherBalancingToStationMwh: number;
  stationToOtherBalancingMwh: number;
  directionDiagnostics: MmsDirectionDiagnostic[];
  knessColumnDiagnostics: MmsKnessColumnDiagnostic[];
  saldoMwh: number;
  rowsRead: number;
  firstDate: string;
  lastDate: string;
  warnings: string[];
};

export type MmsDraftResult = {
  period: ReportPeriod;
  stations: Record<StationId, MmsParsedFile>;
  summary: {
    knessToStationsMwh: number;
    stationsToKnessMwh: number;
    saldoMwh: number;
  };
  warnings: string[];
};

function parseCsvLine(line: string) {
  const values: string[] = [];
  let currentValue = '';
  let insideQuotes = false;

  for (let index = 0; index < line.length; index += 1) {
    const character = line[index];
    if (character === '"') {
      insideQuotes = !insideQuotes;
      continue;
    }

    if (character === ';' && !insideQuotes) {
      values.push(currentValue.trim());
      currentValue = '';
      continue;
    }

    currentValue += character;
  }

  values.push(currentValue.trim());
  return values;
}

function parseNumber(value: string, stationId: StationId) {
  const trimmedValue = value.trim();
  if (trimmedValue.includes(',') && trimmedValue.includes('.')) {
    if (stationId === 'oleksandriya') {
      const normalizedValue = trimmedValue.replace(',', '.');
      const [integerPart, ...fractionParts] = normalizedValue.split('.');
      return Number(`${integerPart}.${fractionParts.join('')}`);
    }

    return Number(trimmedValue.replace(/,/g, ''));
  }

  return Number(trimmedValue.replace(',', '.'));
}

function parseDate(value: string) {
  const match = value.match(/(\d{2})\.(\d{2})\.(\d{4})/);
  if (!match) {
    return '';
  }

  return `${match[3]}-${match[2]}-${match[1]}`;
}

function parsePeriodFromFileName(fileName: string): ReportPeriod | null {
  const match = fileName.match(/(\d{2})_(\d{4})/);
  if (!match) {
    return null;
  }

  return `${match[2]}-${match[1]}` as ReportPeriod;
}

function roundEnergy(value: number) {
  return Math.round((value + Number.EPSILON) * 1000) / 1000;
}

function normalizeDirectionName(value: string) {
  return value
    .toUpperCase()
    .replace(/[’'`ʼ]/g, '')
    .replace(/І/g, 'И')
    .replace(/Ї/g, 'И')
    .replace(/Є/g, 'Е')
    .replace(/\s+/g, ' ')
    .trim();
}

function classifyDirection(header: string): MmsDirectionDiagnostic['group'] {
  const normalized = normalizeDirectionName(header);
  if (normalized.includes('KNESS')) {
    return 'KNESS';
  }
  if (/OPERATOR|OPERATOR_RYNKU|OPERATOR RYNKU|ОПЕРАТОР|RYNKU|РИНКУ/.test(normalized)) {
    return 'Operator';
  }
  if (/(^|[^A-ZА-Я])NA[_ -]|NEK|НЕК|УКРЕНЕРГО|UKRENERGO/.test(normalized)) {
    return 'NA';
  }
  if (/IMBALANCE|BALANCING|НЕБАЛАНС|БАЛАНС/.test(normalized)) {
    return 'Other';
  }
  return 'Unknown';
}

function getDirectionFlow(header: string, stationCode: string): MmsDirectionDiagnostic['flow'] | null {
  if (!header.includes(stationCode)) {
    return null;
  }

  const stationIndex = header.indexOf(stationCode);
  const separatorIndex = header.lastIndexOf('_UA-IPS_', stationIndex);
  if (separatorIndex < 0) {
    return null;
  }

  const leftPart = header.slice(0, separatorIndex);
  const rightPart = header.slice(separatorIndex + '_UA-IPS_'.length);
  if (leftPart.includes(stationCode)) {
    return 'fromStation';
  }
  if (rightPart.includes(stationCode)) {
    return 'toStation';
  }

  return null;
}

function getStationCodes(stationId: StationId) {
  if (stationId === 'oleksandriya') {
    return {
      stationCode: 'OLEKSAND_BESS',
      stationMarkers: /OLEKSAND|ОЛЕКСАНДР/i,
      wrongStationMarkers: /ZNAMENSKA|ЗНАМЕН/i,
    };
  }

  return {
    stationCode: 'ZNAMENSKA_BESS',
    stationMarkers: /ZNAMENSKA|ЗНАМЕН/i,
    wrongStationMarkers: /OLEKSAND|ОЛЕКСАНДР/i,
  };
}

export async function parseMmsFile(input: MmsFileInput): Promise<MmsParsedFile> {
  const csvText = (await input.file.text()).replace(/^\uFEFF/, '');
  const lines = csvText.split(/\r?\n/).filter((line) => line.trim().length > 0);
  if (lines.length < 2) {
    throw new Error(`CSV-файл "${input.file.name}" пустой или не содержит почасовых строк.`);
  }

  const headers = parseCsvLine(lines[0]);
  const { stationCode, wrongStationMarkers } = getStationCodes(input.stationId);
  const knessToStationColumn = headers.findIndex((header) => (
    header.includes('KNESS_ENERGY')
    && header.includes(stationCode)
    && header.indexOf('KNESS_ENERGY') < header.indexOf(stationCode)
  ));
  const stationToKnessColumn = headers.findIndex((header) => (
    header.includes('KNESS_ENERGY')
    && header.includes(stationCode)
    && header.indexOf(stationCode) < header.indexOf('KNESS_ENERGY')
  ));

  if (knessToStationColumn < 0 || stationToKnessColumn < 0) {
    throw new Error(`В файле "${input.file.name}" не найдены направления KNESS ↔ станция для ${input.stationName}.`);
  }

  let knessToStationKwh = 0;
  let stationToKnessKwh = 0;
  const directionColumns = headers
    .map((header, index) => ({
      header,
      index,
      flow: getDirectionFlow(header, stationCode),
      group: classifyDirection(header),
      totalKwh: 0,
    }))
    .filter((direction) => direction.flow);
  let rowsRead = 0;
  let firstDate = '';
  let lastDate = '';

  for (const line of lines.slice(1)) {
    const cells = parseCsvLine(line);
    const date = parseDate(cells[0] ?? '');
    const toStation = parseNumber(cells[knessToStationColumn] ?? '', input.stationId);
    const toKness = parseNumber(cells[stationToKnessColumn] ?? '', input.stationId);

    if (!date || !Number.isFinite(toStation) || !Number.isFinite(toKness)) {
      continue;
    }

    firstDate ||= date;
    lastDate = date;
    rowsRead += 1;
    knessToStationKwh += toStation;
    stationToKnessKwh += toKness;
    for (const direction of directionColumns) {
      const value = parseNumber(cells[direction.index] ?? '', input.stationId);
      if (Number.isFinite(value)) {
        direction.totalKwh += value;
      }
    }
  }

  if (rowsRead === 0) {
    throw new Error(`В файле "${input.file.name}" не найдены почасовые строки MMS.`);
  }

  const period = parsePeriodFromFileName(input.file.name) ?? firstDate.slice(0, 7) as ReportPeriod;
  const warnings: string[] = [];
  if (wrongStationMarkers.test(input.file.name)) {
    warnings.push(`Возможно, файл "${input.file.name}" загружен не в тот слот станции.`);
  }
  if (rowsRead < 700) {
    warnings.push(`Файл "${input.file.name}" содержит меньше 700 почасовых строк.`);
  }

  const knessToStationMwh = roundEnergy(knessToStationKwh / 1000);
  const stationToKnessMwh = roundEnergy(stationToKnessKwh / 1000);
  const knessColumnDiagnostics: MmsKnessColumnDiagnostic[] = [
    {
      direction: 'KNESS_TO_STATION',
      header: headers[knessToStationColumn],
      columnIndex: knessToStationColumn + 1,
      totalKwh: knessToStationKwh,
      totalMwh: knessToStationKwh / 1000,
      rowsUsed: rowsRead,
    },
    {
      direction: 'STATION_TO_KNESS',
      header: headers[stationToKnessColumn],
      columnIndex: stationToKnessColumn + 1,
      totalKwh: stationToKnessKwh,
      totalMwh: stationToKnessKwh / 1000,
      rowsUsed: rowsRead,
    },
  ];
  const directionDiagnostics: MmsDirectionDiagnostic[] = directionColumns.map((direction) => ({
    directionName: direction.header,
    group: direction.group,
    flow: direction.flow ?? 'toStation',
    volumeMwh: roundEnergy(direction.totalKwh / 1000),
  }));
  const sumBy = (group: MmsDirectionDiagnostic['group'], flow: MmsDirectionDiagnostic['flow']) => roundEnergy(
    directionDiagnostics
      .filter((direction) => direction.group === group && direction.flow === flow)
      .reduce((sum, direction) => sum + direction.volumeMwh, 0),
  );
  const operatorToStationMwh = sumBy('Operator', 'toStation');
  const stationToOperatorMwh = sumBy('Operator', 'fromStation');
  const naToStationMwh = sumBy('NA', 'toStation');
  const stationToNaMwh = sumBy('NA', 'fromStation');
  const otherBalancingToStationMwh = sumBy('Other', 'toStation');
  const stationToOtherBalancingMwh = sumBy('Other', 'fromStation');

  return {
    stationId: input.stationId,
    stationName: input.stationName,
    fileName: input.file.name,
    period,
    knessToStationKwh,
    stationToKnessKwh,
    knessToStationMwh,
    stationToKnessMwh,
    operatorToStationMwh,
    stationToOperatorMwh,
    naToStationMwh,
    stationToNaMwh,
    otherBalancingToStationMwh,
    stationToOtherBalancingMwh,
    directionDiagnostics,
    knessColumnDiagnostics,
    saldoMwh: roundEnergy(knessToStationMwh - stationToKnessMwh),
    rowsRead,
    firstDate,
    lastDate,
    warnings,
  };
}

export function buildMmsDraft(files: MmsParsedFile[]): MmsDraftResult {
  const oleksandriya = files.find((file) => file.stationId === 'oleksandriya');
  const znamyanka = files.find((file) => file.stationId === 'znamyanka');

  if (!oleksandriya || !znamyanka) {
    throw new Error('Для расчета MMS нужны 2 CSV-файла: Олександрія и Знаменка.');
  }

  const period = oleksandriya.period;
  const warnings = files.flatMap((file) => file.warnings);
  if (znamyanka.period !== period) {
    warnings.push(`Периоды MMS не совпадают: Олександрія ${period}, Знаменка ${znamyanka.period}.`);
  }

  return {
    period,
    stations: {
      oleksandriya,
      znamyanka,
    },
    summary: {
      knessToStationsMwh: roundEnergy(oleksandriya.knessToStationMwh + znamyanka.knessToStationMwh),
      stationsToKnessMwh: roundEnergy(oleksandriya.stationToKnessMwh + znamyanka.stationToKnessMwh),
      saldoMwh: roundEnergy(oleksandriya.saldoMwh + znamyanka.saldoMwh),
    },
    warnings,
  };
}
