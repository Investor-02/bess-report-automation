import { app, BrowserWindow, ipcMain, dialog } from 'electron';
import path from 'node:path';
import { calculateFcrMonitoringFromPath } from './fcrMonitoring.js';
import { exportTable0Rpch, type Table0ExportInput } from './table0Export.js';
import { exportTable1 } from './table1Export.js';

const isDev = Boolean(process.env.VITE_DEV_SERVER_URL);

function getTable0DefaultFileName(input: Table0ExportInput) {
  const stationFileNames: Record<Table0ExportInput['station'], string> = {
    'Олександрійська БЕСС': 'Олександрія',
    'Знаменська БЕСС': 'Знамянка',
  };
  const latestRecord = input.periods?.length
    ? [...input.periods].sort((first, second) => first.firstDateHeader.localeCompare(second.firstDateHeader)).at(-1) ?? input
    : input;
  const [year, month] = latestRecord.firstDateHeader.split('-');
  const period = month && year ? `${month}_${year}` : latestRecord.firstDateHeader.slice(0, 7).replace('-', '_');

  return `Таблица_0_РПЧ_${stationFileNames[input.station]}_${period}.xlsx`;
}

function getTable1DefaultFileName(exportPeriod: string) {
  const [year, month] = exportPeriod.split('-');
  const period = month && year ? `${month}_${year}` : exportPeriod.replace('-', '_');

  return `Таблица_1_Розрахунки_Укренерго_${period}.xlsx`;
}

function createWindow() {
  const mainWindow = new BrowserWindow({
    width: 1180,
    height: 760,
    minWidth: 980,
    minHeight: 640,
    title: 'UZE Report Automation',
    backgroundColor: '#f5f7fb',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  if (isDev && process.env.VITE_DEV_SERVER_URL) {
    mainWindow.loadURL(process.env.VITE_DEV_SERVER_URL);
  } else {
    mainWindow.loadFile(path.join(__dirname, '../dist/index.html'));
  }
}

app.whenReady().then(() => {
  ipcMain.handle('dialog:open-fcr-file', async () => {
    const result = await dialog.showOpenDialog({
      title: 'Выберите FCR_monitoring.xlsx',
      properties: ['openFile'],
      filters: [{ name: 'Excel files', extensions: ['xlsx'] }],
    });

    return result.canceled ? null : result.filePaths[0];
  });

  ipcMain.handle('fcr:calculate-monitoring', (_event, filePath: string) => {
    return calculateFcrMonitoringFromPath(filePath);
  });

  ipcMain.handle('table0:export-rpch', async (_event, input: Table0ExportInput) => {
    const result = await dialog.showSaveDialog({
      title: 'Сохранить Таблицу_0 РПЧ',
      defaultPath: getTable0DefaultFileName(input),
      filters: [{ name: 'Excel files', extensions: ['xlsx'] }],
    });

    if (result.canceled || !result.filePath) {
      return null;
    }

    return exportTable0Rpch(input, result.filePath);
  });

  ipcMain.handle('table1:export-ukrenergo', async (_event, input: Parameters<typeof exportTable1>[0]) => {
    const result = await dialog.showSaveDialog({
      title: 'Сохранить Таблицу_1 Розрахунки Укренерго',
      defaultPath: getTable1DefaultFileName(input.exportPeriod),
      filters: [{ name: 'Excel files', extensions: ['xlsx'] }],
    });

    if (result.canceled || !result.filePath) {
      return null;
    }

    return exportTable1(input, result.filePath);
  });

  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
