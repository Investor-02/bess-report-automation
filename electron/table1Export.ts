import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';

type ReportPeriod = `${number}-${string}`;
type StationId = 'oleksandriya' | 'znamyanka';
type Table1PaymentRecord = {
  id: string;
  paymentDate: string;
  forPeriod: ReportPeriod;
  amountUah: number;
  createdAt: string;
};
type ProjectReportState = {
  periods: Record<
    ReportPeriod,
    {
      stations: Record<
        StationId,
        {
          table0Fcr: {
            result: {
              serviceVolume: number;
              monthlyPriceUah: number;
              costWithVat: number;
            } | null;
          };
          table1Payments: {
            manualInputs: {
              payments?: Table1PaymentRecord[];
            };
            result: {
              paidAmount: number;
              debtAmount: number;
              payoutPercent: number;
              payments: Table1PaymentRecord[];
            } | null;
          };
        }
      >;
    }
  >;
};

type Table1ExportInput = {
  state: ProjectReportState;
  exportPeriod: ReportPeriod;
};

type StationSheetConfig = {
  stationId: StationId;
  sheetName: string;
  generalColumn: 2 | 3;
};

export type Table1ExportResult = {
  fileName: string;
  outputPath: string;
  templateSource: string;
  exportPeriod: ReportPeriod;
  updatedStationRows: Array<{
    stationId: StationId;
    sheetName: string;
    period: ReportPeriod;
    rowNumber: number;
  }>;
};

const templateFileName = '!!!__1_РОЗРАХУНКИ_УКРЕНЕРГО.xlsx';
const stationSheets: StationSheetConfig[] = [
  { stationId: 'znamyanka', sheetName: 'ЗНАМЕНКА', generalColumn: 2 },
  { stationId: 'oleksandriya', sheetName: 'ОЛЕКСАНДРІЯ', generalColumn: 3 },
];
const monthNamesLower = [
  'січень',
  'лютий',
  'березень',
  'квітень',
  'травень',
  'червень',
  'липень',
  'серпень',
  'вересень',
  'жовтень',
  'листопад',
  'грудень',
];
const monthNamesUpper = monthNamesLower.map((monthName) => monthName.toUpperCase());

function getCellText(cell: ExcelJS.Cell) {
  const value = cell.value;
  if (value === null || value === undefined) {
    return '';
  }
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === 'object' && 'richText' in value && Array.isArray(value.richText)) {
    return value.richText.map((part) => part.text).join('');
  }
  if (typeof value === 'object' && 'result' in value) {
    return String(value.result ?? '');
  }
  return String(value);
}

function getPeriodParts(period: ReportPeriod) {
  const [yearText, monthText] = period.split('-');
  const year = Number(yearText);
  const month = Number(monthText);

  if (!year || month < 1 || month > 12) {
    throw new Error('Период Таблицы_1 должен быть в формате YYYY-MM.');
  }

  return { year, month, monthIndex: month - 1 };
}

function getMonthLabel(period: ReportPeriod, uppercase = false) {
  const { year, monthIndex } = getPeriodParts(period);
  const monthName = uppercase ? monthNamesUpper[monthIndex] : monthNamesLower[monthIndex];

  return `${monthName} ${year}`;
}

function getPeriodFilePart(period: ReportPeriod) {
  const { year, month } = getPeriodParts(period);
  return `${String(month).padStart(2, '0')}_${year}`;
}

function getTemplatePath() {
  const projectRoot = path.join(__dirname, '..');
  const rootTemplatePath = path.join(projectRoot, templateFileName);
  const publicTemplatePath = path.join(projectRoot, 'public', 'templates', templateFileName);

  if (fs.existsSync(rootTemplatePath)) {
    return rootTemplatePath;
  }
  if (fs.existsSync(publicTemplatePath)) {
    return publicTemplatePath;
  }

  throw new Error('Шаблон Таблицы_1 не найден.');
}

function getPeriodEndDateText(period: ReportPeriod) {
  const { year, month } = getPeriodParts(period);
  const lastDay = new Date(year, month, 0).getDate();

  return `станом на ${String(lastDay).padStart(2, '0')}.${String(month).padStart(2, '0')}.${year} року`;
}

function cloneCellStyle(sourceCell: ExcelJS.Cell, targetCell: ExcelJS.Cell) {
  targetCell.style = JSON.parse(JSON.stringify(sourceCell.style ?? {}));
}

function cloneRowStyle(sourceRow: ExcelJS.Row, targetRow: ExcelJS.Row, maxColumn: number) {
  targetRow.height = sourceRow.height;
  for (let column = 1; column <= maxColumn; column += 1) {
    cloneCellStyle(sourceRow.getCell(column), targetRow.getCell(column));
  }
}

function findYearTotalRow(worksheet: ExcelJS.Worksheet, year: number) {
  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const text = getCellText(worksheet.getCell(rowNumber, 1)).toLowerCase();
    if (text.includes('всього') && text.includes(String(year))) {
      return rowNumber;
    }
  }

  return 0;
}

function findStationMonthRow(worksheet: ExcelJS.Worksheet, period: ReportPeriod) {
  const label = getMonthLabel(period).toLowerCase();
  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    if (getCellText(worksheet.getCell(rowNumber, 1)).trim().toLowerCase() === label) {
      return rowNumber;
    }
  }

  return 0;
}

function findGeneralMonthTitleRow(worksheet: ExcelJS.Worksheet, period: ReportPeriod) {
  const label = getMonthLabel(period, true);
  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    if (getCellText(worksheet.getCell(rowNumber, 1)).replace(/\s+/g, ' ').trim().toUpperCase() === label.toUpperCase()) {
      return rowNumber;
    }
  }

  return 0;
}

function isEmptyStationMonthRow(worksheet: ExcelJS.Worksheet, rowNumber: number) {
  for (let column = 1; column <= 12; column += 1) {
    if (getCellText(worksheet.getCell(rowNumber, column)).trim() !== '') {
      return false;
    }
  }

  return true;
}

function getStationYearMonthRows(worksheet: ExcelJS.Worksheet, year: number, totalRowNumber: number) {
  const rows: number[] = [];
  for (let rowNumber = 1; rowNumber < totalRowNumber; rowNumber += 1) {
    const text = getCellText(worksheet.getCell(rowNumber, 1)).trim().toLowerCase();
    if (monthNamesLower.some((monthName) => text === `${monthName} ${year}`)) {
      rows.push(rowNumber);
    }
  }

  return rows;
}

function ensureStationMonthRow(worksheet: ExcelJS.Worksheet, period: ReportPeriod) {
  const { year } = getPeriodParts(period);
  let rowNumber = findStationMonthRow(worksheet, period);
  let totalRowNumber = findYearTotalRow(worksheet, year);

  if (rowNumber) {
    return { rowNumber, totalRowNumber };
  }
  if (!totalRowNumber) {
    throw new Error(`На листе "${worksheet.name}" не найдена итоговая строка за ${year} год.`);
  }

  const monthRows = getStationYearMonthRows(worksheet, year, totalRowNumber);
  const lastMonthRow = monthRows.length ? Math.max(...monthRows) : totalRowNumber - 1;
  const emptyRowBeforeTotal = totalRowNumber > lastMonthRow + 1 && isEmptyStationMonthRow(worksheet, totalRowNumber - 1)
    ? totalRowNumber - 1
    : 0;

  if (emptyRowBeforeTotal) {
    rowNumber = emptyRowBeforeTotal;
    cloneRowStyle(worksheet.getRow(lastMonthRow), worksheet.getRow(rowNumber), 12);
  } else {
    rowNumber = totalRowNumber;
    worksheet.spliceRows(rowNumber, 0, []);
    cloneRowStyle(worksheet.getRow(lastMonthRow), worksheet.getRow(rowNumber), 12);
    totalRowNumber += 1;
  }

  worksheet.getCell(rowNumber, 1).value = getMonthLabel(period);

  return { rowNumber, totalRowNumber };
}

function updateStationTotalRow(worksheet: ExcelJS.Worksheet, year: number) {
  const totalRowNumber = findYearTotalRow(worksheet, year);
  if (!totalRowNumber) {
    return;
  }

  const monthRows = getStationYearMonthRows(worksheet, year, totalRowNumber);
  if (monthRows.length === 0) {
    return;
  }

  const firstRow = Math.min(...monthRows);
  const lastRow = Math.max(...monthRows);
  const row = worksheet.getRow(totalRowNumber);
  row.getCell(2).value = { formula: `SUM(B${firstRow}:B${lastRow})` };
  row.getCell(4).value = { formula: `SUM(D${firstRow}:D${lastRow})` };
  row.getCell(5).value = { formula: `SUM(E${firstRow}:E${lastRow})` };
  row.getCell(6).value = { formula: `SUM(F${firstRow}:F${lastRow})` };
  row.getCell(7).value = { formula: `IF(ISERROR(E${totalRowNumber}/D${totalRowNumber}*100%),"",E${totalRowNumber}/D${totalRowNumber}*100%)` };
  row.getCell(12).value = { formula: `SUM(L${firstRow}:L${lastRow})` };
}

function getSavedStationPeriodData(state: ProjectReportState, period: ReportPeriod, stationId: StationId) {
  const stationState = state.periods[period]?.stations[stationId];
  const table0 = stationState?.table0Fcr.result;
  if (!table0) {
    return null;
  }

  const table1 = stationState?.table1Payments.result;
  const manualPayments = stationState?.table1Payments.manualInputs.payments?.filter((payment) => payment.forPeriod === period) ?? [];
  const resultPayments = table1?.payments?.filter((payment) => payment.forPeriod === period) ?? [];
  const payments = manualPayments.length > 0 ? manualPayments : resultPayments;
  const paidAmount = payments.length > 0
    ? payments.reduce((sum, payment) => sum + payment.amountUah, 0)
    : (table1?.paidAmount ?? 0);
  const accruedWithVat = table0.costWithVat;
  const debtAmount = Math.max(0, accruedWithVat - paidAmount);
  const payoutPercent = accruedWithVat > 0 ? paidAmount / accruedWithVat : 0;

  return {
    serviceVolume: table0.serviceVolume,
    monthlyPriceUah: table0.monthlyPriceUah,
    accruedWithVat,
    paidAmount,
    debtAmount,
    payoutPercent,
    payments,
  };
}

function getLatestPayment(payments: Table1PaymentRecord[]) {
  if (payments.length === 0) {
    return null;
  }

  return [...payments].sort((a, b) => a.paymentDate.localeCompare(b.paymentDate)).at(-1) ?? null;
}

function fillStationRow(worksheet: ExcelJS.Worksheet, period: ReportPeriod, rowNumber: number, data: NonNullable<ReturnType<typeof getSavedStationPeriodData>>) {
  const latestPayment = getLatestPayment(data.payments);
  const row = worksheet.getRow(rowNumber);
  row.getCell(1).value = getMonthLabel(period);
  row.getCell(2).value = data.serviceVolume;
  row.getCell(3).value = data.monthlyPriceUah;
  row.getCell(4).value = { formula: `ROUND(ROUND(B${rowNumber}*C${rowNumber},2)*1.2,2)`, result: data.accruedWithVat };
  row.getCell(5).value = data.paidAmount;
  row.getCell(6).value = data.debtAmount;
  row.getCell(7).value = data.payoutPercent;
  row.getCell(10).value = latestPayment?.paymentDate ? new Date(`${latestPayment.paymentDate}T00:00:00`) : null;
  row.getCell(11).value = latestPayment ? getPeriodParts(latestPayment.forPeriod).month.toString().padStart(2, '0') : null;
  row.getCell(12).value = data.paidAmount || null;
}

function getGeneralRowsForPeriod(worksheet: ExcelJS.Worksheet, period: ReportPeriod) {
  const titleRow = findGeneralMonthTitleRow(worksheet, period);
  if (!titleRow) {
    throw new Error(`В общей таблице не найдена строка месяца ${getMonthLabel(period, true)}.`);
  }

  return {
    titleRow,
    volumeRow: titleRow + 1,
    accruedRow: titleRow + 2,
    paidRow: titleRow + 3,
    debtRow: titleRow + 4,
    percentRow: titleRow + 5,
  };
}

function fillGeneralMetricRow(
  worksheet: ExcelJS.Worksheet,
  rowNumber: number,
  znamyankaValue: number | null,
  oleksandriyaValue: number | null,
  cumulativeRow?: number,
) {
  const rowTotal = (znamyankaValue ?? 0) + (oleksandriyaValue ?? 0);
  const row = worksheet.getRow(rowNumber);
  row.getCell(2).value = znamyankaValue;
  row.getCell(3).value = oleksandriyaValue;
  row.getCell(4).value = { formula: `SUM(B${rowNumber}:C${rowNumber})`, result: rowTotal };
  row.getCell(5).value = cumulativeRow
    ? { formula: `E${cumulativeRow}+D${rowNumber}`, result: rowTotal }
    : { formula: `+D${rowNumber}`, result: rowTotal };
}

function findPreviousGeneralMetricRow(worksheet: ExcelJS.Worksheet, rowNumber: number, metricLabel: string) {
  for (let currentRow = rowNumber - 1; currentRow >= 1; currentRow -= 1) {
    if (getCellText(worksheet.getCell(currentRow, 1)).trim() === metricLabel) {
      return currentRow;
    }
  }

  return 0;
}

function updateGeneralForPeriod(worksheet: ExcelJS.Worksheet, state: ProjectReportState, period: ReportPeriod) {
  const rows = getGeneralRowsForPeriod(worksheet, period);
  const znamyanka = getSavedStationPeriodData(state, period, 'znamyanka');
  const oleksandriya = getSavedStationPeriodData(state, period, 'oleksandriya');

  fillGeneralMetricRow(
    worksheet,
    rows.volumeRow,
    znamyanka ? znamyanka.serviceVolume : null,
    oleksandriya ? oleksandriya.serviceVolume : null,
    findPreviousGeneralMetricRow(worksheet, rows.volumeRow, 'дРПЧ_с, МВт'),
  );
  fillGeneralMetricRow(
    worksheet,
    rows.accruedRow,
    znamyanka ? znamyanka.accruedWithVat : null,
    oleksandriya ? oleksandriya.accruedWithVat : null,
    findPreviousGeneralMetricRow(worksheet, rows.accruedRow, 'Нараховано, Гривень'),
  );
  fillGeneralMetricRow(
    worksheet,
    rows.paidRow,
    znamyanka ? znamyanka.paidAmount : null,
    oleksandriya ? oleksandriya.paidAmount : null,
    findPreviousGeneralMetricRow(worksheet, rows.paidRow, 'Виплачено, Гривень'),
  );
  fillGeneralMetricRow(
    worksheet,
    rows.debtRow,
    znamyanka ? znamyanka.debtAmount : null,
    oleksandriya ? oleksandriya.debtAmount : null,
    findPreviousGeneralMetricRow(worksheet, rows.debtRow, 'Заборгованість, Гривень'),
  );
  worksheet.getCell(rows.percentRow, 2).value = {
    formula: `IF(ISERROR(B${rows.paidRow}/B${rows.accruedRow}),"",B${rows.paidRow}/B${rows.accruedRow})`,
  };
  worksheet.getCell(rows.percentRow, 3).value = {
    formula: `IF(ISERROR(C${rows.paidRow}/C${rows.accruedRow}),"",C${rows.paidRow}/C${rows.accruedRow})`,
  };
  worksheet.getCell(rows.percentRow, 4).value = {
    formula: `IF(ISERROR(D${rows.paidRow}/D${rows.accruedRow}),"",D${rows.paidRow}/D${rows.accruedRow})`,
  };
  worksheet.getCell(rows.percentRow, 5).value = {
    formula: `IF(ISERROR(E${rows.paidRow}/E${rows.accruedRow}),"",E${rows.paidRow}/E${rows.accruedRow})`,
  };
}

function updateGeneralSummaryRows(worksheet: ExcelJS.Worksheet) {
  worksheet.getCell('B190').value = { formula: 'B170' };
  worksheet.getCell('C190').value = { formula: 'C170' };
  worksheet.getCell('D190').value = { formula: 'SUM(B190:C190)' };
  worksheet.getCell('E190').value = { formula: 'IF(D190=0,"",100%)' };
  worksheet.getCell('B197').value = { formula: 'B171' };
  worksheet.getCell('C197').value = { formula: 'C171' };
  worksheet.getCell('D197').value = { formula: 'SUM(B197:C197)' };
  worksheet.getCell('E197').value = { formula: 'IF(ISERROR(D197/D190),"",D197/D190)' };
  worksheet.getCell('B204').value = { formula: 'B190-B197' };
  worksheet.getCell('C204').value = { formula: 'C190-C197' };
  worksheet.getCell('D204').value = { formula: 'SUM(B204:C204)' };
  worksheet.getCell('E204').value = { formula: 'IF(ISERROR(D204/D190),"",D204/D190)' };
  for (const rowNumber of [191, 198, 205]) {
    worksheet.getCell(rowNumber, 2).value = { formula: `SUM(B${rowNumber - 2}:B${rowNumber - 1})` };
    worksheet.getCell(rowNumber, 3).value = { formula: `SUM(C${rowNumber - 2}:C${rowNumber - 1})` };
    worksheet.getCell(rowNumber, 4).value = { formula: `SUM(D${rowNumber - 2}:D${rowNumber - 1})` };
  }
  worksheet.getCell('E198').value = { formula: 'IF(ISERROR(D198/D191),"",D198/D191)' };
  worksheet.getCell('E205').value = { formula: 'IF(ISERROR(D205/D191),"",D205/D191)' };
}

function getPeriodsWithTable0(state: ProjectReportState) {
  return Object.keys(state.periods)
    .filter((period): period is ReportPeriod => /^\d{4}-\d{2}$/.test(period))
    .filter((period) => stationSheets.some(({ stationId }) => Boolean(state.periods[period].stations[stationId].table0Fcr.result)))
    .sort();
}

export async function exportTable1(input: Table1ExportInput, requestedOutputPath?: string): Promise<Table1ExportResult> {
  const workbook = new ExcelJS.Workbook();
  const templateSource = getTemplatePath();
  await workbook.xlsx.readFile(templateSource);
  const periods = getPeriodsWithTable0(input.state);
  if (periods.length === 0) {
    throw new Error('В памяти проекта нет сохраненной Таблицы_0 для экспорта Таблицы_1.');
  }

  const updatedStationRows: Table1ExportResult['updatedStationRows'] = [];
  const latestPeriod = periods.includes(input.exportPeriod) ? input.exportPeriod : periods.at(-1) ?? input.exportPeriod;
  const generalWorksheet = workbook.getWorksheet('ЗАГАЛЬНА ТАБЛИЦЯ');
  if (!generalWorksheet) {
    throw new Error('В шаблоне не найден лист "ЗАГАЛЬНА ТАБЛИЦЯ".');
  }
  generalWorksheet.getCell('A2').value = getPeriodEndDateText(latestPeriod);
  generalWorksheet.getCell('A89').value = { formula: 'A2', result: getPeriodEndDateText(latestPeriod) };
  generalWorksheet.getCell('A185').value = { formula: 'A2', result: getPeriodEndDateText(latestPeriod) };

  for (const period of periods) {
    const { year } = getPeriodParts(period);
    for (const stationConfig of stationSheets) {
      const data = getSavedStationPeriodData(input.state, period, stationConfig.stationId);
      if (!data) {
        continue;
      }

      const worksheet = workbook.getWorksheet(stationConfig.sheetName);
      if (!worksheet) {
        throw new Error(`В шаблоне не найден лист "${stationConfig.sheetName}".`);
      }
      const { rowNumber } = ensureStationMonthRow(worksheet, period);
      fillStationRow(worksheet, period, rowNumber, data);
      updateStationTotalRow(worksheet, year);
      updatedStationRows.push({
        stationId: stationConfig.stationId,
        sheetName: stationConfig.sheetName,
        period,
        rowNumber,
      });
    }
    updateGeneralForPeriod(generalWorksheet, input.state, period);
  }
  updateGeneralSummaryRows(generalWorksheet);

  const fileName = `Таблица_1_Розрахунки_Укренерго_${getPeriodFilePart(latestPeriod)}.xlsx`;
  const outputPath = requestedOutputPath || path.join(path.dirname(templateSource), fileName);
  await workbook.xlsx.writeFile(outputPath);

  return {
    fileName,
    outputPath,
    templateSource,
    exportPeriod: latestPeriod,
    updatedStationRows,
  };
}
