import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';

export type Table0ExportRecord = {
  station: 'Олександрійська БЕСС' | 'Знаменська БЕСС';
  firstDateHeader: string;
  certifiedPowerMw: number;
  trueHours: number;
  falseHours: number;
  serviceVolume: number;
  fcrTariffEur: number;
  eurRate: string;
  monthlyPriceUah: number;
  costWithoutVat: number;
  vat: number;
  costWithVat: number;
};

export type Table0ExportInput = Table0ExportRecord & {
  periods?: Table0ExportRecord[];
};

export type Table0ExportResult = {
  outputPath: string;
  templateSource: string;
  monthLabel: string;
  rowNumber: number;
  totalRowNumber: number;
  totalFormulaRange: string;
  mode: 'updated' | 'filled-empty' | 'inserted-before-total';
};

const monthNames = [
  'Січень',
  'Лютий',
  'Березень',
  'Квітень',
  'Травень',
  'Червень',
  'Липень',
  'Серпень',
  'Вересень',
  'Жовтень',
  'Листопад',
  'Грудень',
];

const stationFileNames: Record<Table0ExportRecord['station'], string> = {
  'Олександрійська БЕСС': 'Олександрія',
  'Знаменська БЕСС': 'Знамянка',
};

function getTemplatePath() {
  const projectRoot = path.join(__dirname, '..');
  const publicTemplatePath = path.join(projectRoot, 'public', 'templates', 'table_0_rpch_template.xlsx');

  if (fs.existsSync(publicTemplatePath)) {
    return publicTemplatePath;
  }

  throw new Error('Шаблон public/templates/table_0_rpch_template.xlsx не найден.');
}

function getMonthLabel(firstDateHeader: string) {
  const [yearText, monthText] = firstDateHeader.split('-');
  const year = Number(yearText);
  const monthIndex = Number(monthText) - 1;

  if (!year || monthIndex < 0 || monthIndex > 11) {
    throw new Error('Не получилось определить отчетный месяц по данным FCR monitoring.');
  }

  return `${monthNames[monthIndex]} ${year}`;
}

function getMonthFilePart(firstDateHeader: string) {
  const [yearText, monthText] = firstDateHeader.split('-');
  const year = Number(yearText);
  const month = Number(monthText);

  if (!year || month < 1 || month > 12) {
    throw new Error('Не получилось определить отчетный месяц по данным FCR monitoring.');
  }

  return `${String(month).padStart(2, '0')}_${year}`;
}

function getPeriodCellValue(firstDateHeader: string) {
  const [yearText, monthText] = firstDateHeader.split('-');
  const year = Number(yearText);
  const month = Number(monthText);

  if (!year || month < 1 || month > 12) {
    throw new Error('Не получилось определить отчетный месяц по данным FCR monitoring.');
  }

  const lastDay = new Date(year, month, 0).getDate();

  return `на ${String(lastDay).padStart(2, '0')}.${String(month).padStart(2, '0')}.${year}`;
}

function getSafeFilePart(value: string) {
  return value.replace(/[<>:"/\\|?*]+/g, '').replace(/\s+/g, '_');
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

function isEmptyTemplateMonthRow(worksheet: ExcelJS.Worksheet, rowNumber: number) {
  for (let column = 1; column <= 11; column += 1) {
    if (getCellText(worksheet.getCell(rowNumber, column)).trim() !== '') {
      return false;
    }
  }

  return true;
}

function getCellText(cell: ExcelJS.Cell) {
  const value = cell.value;
  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'object' && 'richText' in value && Array.isArray(value.richText)) {
    return value.richText.map((part) => part.text).join('');
  }
  if (typeof value === 'object' && 'result' in value) {
    return String(value.result ?? '');
  }
  return String(value);
}

function worksheetContainsText(workbook: ExcelJS.Workbook, searchText: string) {
  for (const worksheet of workbook.worksheets) {
    for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
      const row = worksheet.getRow(rowNumber);
      for (let column = 1; column <= worksheet.columnCount; column += 1) {
        if (getCellText(row.getCell(column)).includes(searchText)) {
          return true;
        }
      }
    }
  }

  return false;
}

function selectedWorksheetContainsText(worksheet: ExcelJS.Worksheet, searchText: string) {
  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    for (let column = 1; column <= worksheet.columnCount; column += 1) {
      if (getCellText(row.getCell(column)).includes(searchText)) {
        return true;
      }
    }
  }

  return false;
}

function keepOnlySelectedWorksheet(workbook: ExcelJS.Workbook, selectedWorksheet: ExcelJS.Worksheet) {
  const selectedWorksheetId = selectedWorksheet.id;

  for (const worksheet of [...workbook.worksheets]) {
    if (worksheet.id !== selectedWorksheetId) {
      workbook.removeWorksheet(worksheet.id);
    }
  }
}

function escapeRegExp(value: string) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function formulaReferencesSheet(formula: string, worksheetName: string) {
  const escapedName = escapeRegExp(worksheetName);

  return new RegExp(`(^|[^A-Za-zА-Яа-яІіЇїЄєҐґ0-9_])'${escapedName}'!`).test(formula)
    || new RegExp(`(^|[^A-Za-zА-Яа-яІіЇїЄєҐґ0-9_])${escapedName}!`).test(formula);
}

function clearBrokenCrossSheetReferences(
  worksheet: ExcelJS.Worksheet,
  removedWorksheetNames: string[],
  firstDateHeader: string,
) {
  worksheet.getCell('K2').value = getPeriodCellValue(firstDateHeader);

  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    for (let column = 1; column <= worksheet.columnCount; column += 1) {
      const cell = row.getCell(column);
      const value = cell.value;

      if (!value || typeof value !== 'object' || !('formula' in value)) {
        continue;
      }

      const formula = String(value.formula ?? '');
      const referencesRemovedSheet = removedWorksheetNames.some((worksheetName) => formulaReferencesSheet(formula, worksheetName));
      const referencesExternalWorkbook = formula.includes('[') || formula.toUpperCase().includes('#REF!');

      if (referencesRemovedSheet || referencesExternalWorkbook) {
        cell.value = value.result ?? null;
      }
    }
  }

  worksheet.getCell('K2').value = getPeriodCellValue(firstDateHeader);
}

function findTotalRow(worksheet: ExcelJS.Worksheet) {
  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    if (getCellText(worksheet.getCell(rowNumber, 1)).trim().toUpperCase().includes('ВСЬОГО')) {
      return rowNumber;
    }
  }

  return 0;
}

function findMonthRows(worksheet: ExcelJS.Worksheet) {
  const monthRows: number[] = [];

  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const text = getCellText(worksheet.getCell(rowNumber, 1)).trim();
    if (monthNames.some((monthName) => text.startsWith(`${monthName} `))) {
      monthRows.push(rowNumber);
    }
  }

  return monthRows;
}

function setMonthRow(worksheet: ExcelJS.Worksheet, rowNumber: number, monthLabel: string, input: Table0ExportRecord) {
  const row = worksheet.getRow(rowNumber);
  const rate = Number(input.eurRate.replace(',', '.'));

  row.getCell(1).value = monthLabel;
  row.getCell(2).value = input.certifiedPowerMw;
  row.getCell(3).value = input.trueHours;
  row.getCell(4).value = input.falseHours;
  row.getCell(5).value = { formula: `B${rowNumber}*C${rowNumber}`, result: input.serviceVolume };
  row.getCell(6).value = input.fcrTariffEur;
  row.getCell(7).value = rate;
  row.getCell(8).value = { formula: `ROUND(F${rowNumber}*G${rowNumber},2)`, result: input.monthlyPriceUah };
  row.getCell(9).value = { formula: `H${rowNumber}*E${rowNumber}`, result: input.costWithoutVat };
  row.getCell(10).value = { formula: `ROUND(I${rowNumber}*20/100,2)`, result: input.vat };
  row.getCell(11).value = { formula: `I${rowNumber}*1.2`, result: input.costWithVat };
  row.commit();
}

function getPeriodKey(firstDateHeader: string) {
  return firstDateHeader.slice(0, 7);
}

function getExportRecords(input: Table0ExportInput): Table0ExportRecord[] {
  const records = (input.periods?.length ? input.periods : [input]).filter((record) => record.station === input.station);
  const uniqueRecords = new Map<string, Table0ExportRecord>();

  for (const record of records) {
    uniqueRecords.set(getPeriodKey(record.firstDateHeader), record);
  }

  return [...uniqueRecords.values()].sort((first, second) => getPeriodKey(first.firstDateHeader).localeCompare(getPeriodKey(second.firstDateHeader)));
}

function upsertMonthRow(
  worksheet: ExcelJS.Worksheet,
  input: Table0ExportRecord,
  totalRowNumber: number,
) {
  const monthLabel = getMonthLabel(input.firstDateHeader);
  let monthRows = findMonthRows(worksheet).filter((monthRow) => monthRow < totalRowNumber);
  if (monthRows.length === 0) {
    throw new Error('В шаблоне не найдены строки месяцев для копирования оформления.');
  }

  const existingRowNumber = monthRows.find((monthRow) => getCellText(worksheet.getCell(monthRow, 1)).trim() === monthLabel);
  let rowNumber = existingRowNumber ?? 0;
  let mode: Table0ExportResult['mode'] = 'updated';
  let nextTotalRowNumber = totalRowNumber;

  if (!rowNumber) {
    const lastMonthRow = Math.max(...monthRows);
    const emptyRowBeforeTotal = nextTotalRowNumber > lastMonthRow + 1 && isEmptyTemplateMonthRow(worksheet, nextTotalRowNumber - 1)
      ? nextTotalRowNumber - 1
      : 0;

    if (emptyRowBeforeTotal) {
      rowNumber = emptyRowBeforeTotal;
      cloneRowStyle(worksheet.getRow(lastMonthRow), worksheet.getRow(rowNumber), 11);
      mode = 'filled-empty';
    } else {
      rowNumber = nextTotalRowNumber;
      worksheet.spliceRows(rowNumber, 0, []);
      cloneRowStyle(worksheet.getRow(lastMonthRow), worksheet.getRow(rowNumber), 11);
      nextTotalRowNumber += 1;
      mode = 'inserted-before-total';
    }
  }

  setMonthRow(worksheet, rowNumber, monthLabel, input);

  return {
    monthLabel,
    rowNumber,
    totalRowNumber: nextTotalRowNumber,
    mode,
  };
}

function refreshTotalRow(worksheet: ExcelJS.Worksheet, totalRowNumber: number, firstMonthRow: number, lastMonthRow: number) {
  const row = worksheet.getRow(totalRowNumber);

  row.getCell(3).value = { formula: `SUM(C${firstMonthRow}:C${lastMonthRow})` };
  row.getCell(4).value = { formula: `SUM(D${firstMonthRow}:D${lastMonthRow})` };
  row.getCell(5).value = { formula: `SUM(E${firstMonthRow}:E${lastMonthRow})` };
  row.getCell(9).value = { formula: `SUM(I${firstMonthRow}:I${lastMonthRow})` };
  row.getCell(10).value = { formula: `SUM(J${firstMonthRow}:J${lastMonthRow})` };
  row.getCell(11).value = { formula: `SUM(K${firstMonthRow}:K${lastMonthRow})` };
  row.commit();

  return `C${firstMonthRow}:C${lastMonthRow}; D${firstMonthRow}:D${lastMonthRow}; E${firstMonthRow}:E${lastMonthRow}; I${firstMonthRow}:I${lastMonthRow}; J${firstMonthRow}:J${lastMonthRow}; K${firstMonthRow}:K${lastMonthRow}`;
}

export async function exportTable0Rpch(input: Table0ExportInput, requestedOutputPath?: string): Promise<Table0ExportResult> {
  const workbook = new ExcelJS.Workbook();
  const templatePath = getTemplatePath();
  await workbook.xlsx.readFile(templatePath);

  const worksheet = workbook.getWorksheet(input.station);
  if (!worksheet) {
    throw new Error(`Лист "${input.station}" не найден в шаблоне Таблицы_0.`);
  }
  const selectedSheetHadTestMarker = selectedWorksheetContainsText(worksheet, 'ТЕСТОВЫЙ ШАБЛОН');
  const exportRecords = getExportRecords(input);
  if (exportRecords.length === 0) {
    throw new Error('В памяти проекта нет сохраненных периодов Таблицы_0 для выбранной станции.');
  }
  const latestRecord = exportRecords.at(-1) ?? input;

  let totalRowNumber = findTotalRow(worksheet);
  if (!totalRowNumber) {
    throw new Error('В шаблоне не найдена строка "ВСЬОГО".');
  }

  let lastUpdatedRow = {
    monthLabel: getMonthLabel(latestRecord.firstDateHeader),
    rowNumber: 0,
    totalRowNumber,
    mode: 'updated' as Table0ExportResult['mode'],
  };
  for (const record of exportRecords) {
    lastUpdatedRow = upsertMonthRow(worksheet, record, totalRowNumber);
    totalRowNumber = lastUpdatedRow.totalRowNumber;
  }

  const monthRows = findMonthRows(worksheet).filter((monthRow) => monthRow < totalRowNumber);
  const firstMonthRow = Math.min(...monthRows);
  const lastMonthRow = totalRowNumber - 1;
  const totalFormulaRange = refreshTotalRow(worksheet, totalRowNumber, firstMonthRow, lastMonthRow);

  const removedWorksheetNames = workbook.worksheets
    .filter((sheet) => sheet.id !== worksheet.id)
    .map((sheet) => sheet.name);
  clearBrokenCrossSheetReferences(worksheet, removedWorksheetNames, latestRecord.firstDateHeader);
  keepOnlySelectedWorksheet(workbook, worksheet);

  const stationFileName = stationFileNames[input.station] ?? input.station;
  const defaultFileName = `Таблица_0_РПЧ_${getSafeFilePart(stationFileName)}_${getMonthFilePart(latestRecord.firstDateHeader)}.xlsx`;
  const outputPath = requestedOutputPath || path.join(path.dirname(templatePath), defaultFileName);

  if (selectedSheetHadTestMarker && !worksheetContainsText(workbook, 'ТЕСТОВЫЙ ШАБЛОН')) {
    throw new Error('Проверка шаблона не пройдена: текст "ТЕСТОВЫЙ ШАБЛОН" не сохранился после экспорта.');
  }

  await workbook.xlsx.writeFile(outputPath);

  return {
    outputPath,
    templateSource: templatePath,
    monthLabel: lastUpdatedRow.monthLabel,
    rowNumber: lastUpdatedRow.rowNumber,
    totalRowNumber,
    totalFormulaRange,
    mode: lastUpdatedRow.mode,
  };
}
