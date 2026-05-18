import { contextBridge, ipcRenderer } from 'electron';

contextBridge.exposeInMainWorld('uzeApp', {
  openFcrFile: () => ipcRenderer.invoke('dialog:open-fcr-file') as Promise<string | null>,
  calculateFcrMonitoring: (filePath: string) =>
    ipcRenderer.invoke('fcr:calculate-monitoring', filePath) as Promise<{
      trueHours: number;
      falseHours: number;
      totalHours: number;
      debug: {
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
    }>,
  exportTable0Rpch: (input: {
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
    periods?: Array<{
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
    }>;
  }) =>
    ipcRenderer.invoke('table0:export-rpch', input) as Promise<{
      outputPath: string;
      templateSource: string;
      monthLabel: string;
      rowNumber: number;
      totalRowNumber: number;
      totalFormulaRange: string;
      mode: 'updated' | 'filled-empty' | 'inserted-before-total';
    } | null>,
  exportTable1Ukrenergo: (input: unknown) =>
    ipcRenderer.invoke('table1:export-ukrenergo', input) as Promise<{
      fileName: string;
      outputPath: string;
      templateSource: string;
      exportPeriod: string;
      updatedStationRows: Array<{
        stationId: 'oleksandriya' | 'znamyanka';
        sheetName: string;
        period: string;
        rowNumber: number;
      }>;
    } | null>,
});
