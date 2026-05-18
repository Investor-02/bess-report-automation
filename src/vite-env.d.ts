/// <reference types="vite/client" />

interface Window {
  uzeApp?: {
    openFcrFile: () => Promise<string | null>;
    calculateFcrMonitoring: (filePath: string) => Promise<{
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
    }>;
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
    }) => Promise<{
      outputPath: string;
      templateSource: string;
      monthLabel: string;
      rowNumber: number;
      totalRowNumber: number;
      totalFormulaRange: string;
      mode: 'updated' | 'filled-empty' | 'inserted-before-total';
    } | null>;
    exportTable1Ukrenergo: (input: unknown) => Promise<{
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
    } | null>;
  };
}
