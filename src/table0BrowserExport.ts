import type ExcelJS from 'exceljs';

export type BrowserTable0ExportInput = {
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

export type BrowserTable0ExportResult = {
  fileName: string;
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

function getMonthLabel(firstDateHeader: string) {
  const [yearText, monthText] = firstDateHeader.split('-');
  const year = Number(yearText);
  const monthIndex = Number(monthText) - 1;

  if (!year || monthIndex < 0 || monthIndex > 11) {
    throw new Error('Не получилось определить отчетный месяц по данным FCR monitoring.');
  }

  return `${monthNames[monthIndex]} ${year}`;
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

function setMonthRow(worksheet: ExcelJS.Worksheet, rowNumber: number, monthLabel: string, input: BrowserTable0ExportInput) {
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
}

function refreshTotalRow(worksheet: ExcelJS.Worksheet, totalRowNumber: number, firstMonthRow: number, lastMonthRow: number) {
  const row = worksheet.getRow(totalRowNumber);

  row.getCell(3).value = { formula: `SUM(C${firstMonthRow}:C${lastMonthRow})` };
  row.getCell(4).value = { formula: `SUM(D${firstMonthRow}:D${lastMonthRow})` };
  row.getCell(5).value = { formula: `SUM(E${firstMonthRow}:E${lastMonthRow})` };
  row.getCell(9).value = { formula: `SUM(I${firstMonthRow}:I${lastMonthRow})` };
  row.getCell(10).value = { formula: `SUM(J${firstMonthRow}:J${lastMonthRow})` };
  row.getCell(11).value = { formula: `SUM(K${firstMonthRow}:K${lastMonthRow})` };

  return `C${firstMonthRow}:C${lastMonthRow}; D${firstMonthRow}:D${lastMonthRow}; E${firstMonthRow}:E${lastMonthRow}; I${firstMonthRow}:I${lastMonthRow}; J${firstMonthRow}:J${lastMonthRow}; K${firstMonthRow}:K${lastMonthRow}`;
}

function downloadWorkbook(buffer: ExcelJS.Buffer, fileName: string) {
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

export async function exportTable0RpchInBrowser(input: BrowserTable0ExportInput): Promise<BrowserTable0ExportResult> {
  const monthLabel = getMonthLabel(input.firstDateHeader);
  const ExcelJSModule = await import('exceljs');
  const workbook = new ExcelJSModule.Workbook();
  const templateSource = `/templates/table_0_rpch_template.xlsx?v=${Date.now()}`;
  const templateResponse = await fetch(templateSource, { cache: 'no-store' });

  if (!templateResponse.ok) {
    throw new Error('Шаблон public/templates/table_0_rpch_template.xlsx не найден.');
  }

  await workbook.xlsx.load(await templateResponse.arrayBuffer());
  const templateHadTestMarker = worksheetContainsText(workbook, 'ТЕСТОВЫЙ ШАБЛОН');

  const worksheet = workbook.getWorksheet(input.station);
  if (!worksheet) {
    throw new Error(`Лист "${input.station}" не найден в шаблоне Таблицы_0.`);
  }

  let totalRowNumber = findTotalRow(worksheet);
  if (!totalRowNumber) {
    throw new Error('В шаблоне не найдена строка "ВСЬОГО".');
  }

  let monthRows = findMonthRows(worksheet).filter((monthRow) => monthRow < totalRowNumber);
  if (monthRows.length === 0) {
    throw new Error('В шаблоне не найдены строки месяцев для копирования оформления.');
  }

  const existingRowNumber = monthRows.find((rowNumber) => getCellText(worksheet.getCell(rowNumber, 1)).trim() === monthLabel);
  let rowNumber = existingRowNumber ?? 0;
  let mode: BrowserTable0ExportResult['mode'] = 'updated';

  if (!rowNumber) {
    const lastMonthRow = Math.max(...monthRows);
    const emptyRowBeforeTotal = totalRowNumber > lastMonthRow + 1 && isEmptyTemplateMonthRow(worksheet, totalRowNumber - 1)
      ? totalRowNumber - 1
      : 0;

    if (emptyRowBeforeTotal) {
      rowNumber = emptyRowBeforeTotal;
      cloneRowStyle(worksheet.getRow(lastMonthRow), worksheet.getRow(rowNumber), 11);
      mode = 'filled-empty';
    } else {
      rowNumber = totalRowNumber;
      worksheet.spliceRows(rowNumber, 0, []);
      cloneRowStyle(worksheet.getRow(lastMonthRow), worksheet.getRow(rowNumber), 11);
      totalRowNumber += 1;
      mode = 'inserted-before-total';
    }
  }

  setMonthRow(worksheet, rowNumber, monthLabel, input);

  monthRows = findMonthRows(worksheet).filter((monthRow) => monthRow < totalRowNumber);
  const firstMonthRow = Math.min(...monthRows);
  const lastMonthRow = totalRowNumber - 1;
  const totalFormulaRange = refreshTotalRow(worksheet, totalRowNumber, firstMonthRow, lastMonthRow);

  if (templateHadTestMarker && !worksheetContainsText(workbook, 'ТЕСТОВЫЙ ШАБЛОН')) {
    throw new Error('Проверка шаблона не пройдена: текст "ТЕСТОВЫЙ ШАБЛОН" не сохранился после экспорта.');
  }

  const [monthName, year] = monthLabel.split(' ');
  const fileName = `Таблица_0_РПЧ_${getSafeFilePart(input.station)}_${monthName}_${year}.xlsx`;
  const buffer = await workbook.xlsx.writeBuffer();
  downloadWorkbook(buffer, fileName);

  return {
    fileName,
    templateSource,
    monthLabel,
    rowNumber,
    totalRowNumber,
    totalFormulaRange,
    mode,
  };
}
