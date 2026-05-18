import { useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertCircle,
  AlertTriangle,
  BarChart3,
  Calculator,
  Database,
  Download,
  FileSpreadsheet,
  FileText,
  FolderOpen,
  Gauge,
  LineChart,
  Plug,
  Settings,
  Trash2,
  Zap,
} from 'lucide-react';
import { calculateFcrMonitoringFromFile, type FcrMonitoringResult } from './fcrMonitoring';
import { exportTable0RpchInBrowser, type BrowserTable0ExportRecord } from './table0BrowserExport';
import { exportTable1InBrowser, type Table1BrowserExportResult } from './table1BrowserExport';
import {
  loadReportState,
  projectReportStateStorageKey,
  updateStationModule,
  type ProjectReportState,
  type ReportPeriod,
  type StationId,
  type Table1PaymentRecord,
} from './state/projectReportState';
import { clearDraftState, loadDraftState, saveDraftState } from './state/draftReportState';

type Station = 'Олександрійська БЕСС' | 'Знаменська БЕСС';

type StationConfig = {
  certifiedPowerMw: number;
  fcrTariffEur: number;
};

type PaymentCalculation = {
  station: Station;
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

type PersistedAppState = {
  station: Station;
  eurRate: string;
  fileName: string;
  result: FcrMonitoringResult | null;
  paymentCalculation: PaymentCalculation | null;
};

type ExportStatus = {
  fileName: string;
  station: Station;
  monthLabel: string;
  action: 'Обновлена строка' | 'Заполнена пустая строка' | 'Добавлена строка перед ВСЬОГО';
  rowNumber: number;
  totalRowNumber: number;
  totalFormulaRange: string;
  templateSource: string;
};

type Table1ExportStatus = Table1BrowserExportResult & { outputPath?: string };

type FcrSubTab = 'table0' | 'table1';

const stationConfig: Record<Station, StationConfig> = {
  'Олександрійська БЕСС': {
    certifiedPowerMw: 10,
    fcrTariffEur: 13.99830431,
  },
  'Знаменська БЕСС': {
    certifiedPowerMw: 8,
    fcrTariffEur: 15.41138234,
  },
};

const stationLabels: Record<StationId, Station> = {
  oleksandriya: 'Олександрійська БЕСС',
  znamyanka: 'Знаменська БЕСС',
};

const stationMemoryOrder: StationId[] = ['oleksandriya', 'znamyanka'];

const modules = [
  { title: 'РПЧ / FCR', description: 'Начисление и оплата', status: 'Активно', icon: Gauge, enabled: true },
  { title: 'РДН / ВДР', description: 'Почасовые рынки', status: 'Скоро', icon: LineChart, enabled: false },
  { title: 'Небалансы', description: 'Отклонения и сверки', status: 'Скоро', icon: Zap, enabled: false },
  { title: 'DataHub', description: 'Импорт данных', status: 'Скоро', icon: Database, enabled: false },
  { title: 'Итоговый отчет', description: 'Сводный файл', status: 'Скоро', icon: FileText, enabled: false },
];

const moneyFormatter = new Intl.NumberFormat('ru-RU', {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

const numberFormatter = new Intl.NumberFormat('ru-RU', {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

function parseRate(rate: string) {
  return Number(rate.replace(',', '.').trim());
}

function roundToTwoDecimals(value: number) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function isStation(value: unknown): value is Station {
  return value === 'Олександрійська БЕСС' || value === 'Знаменська БЕСС';
}

function getStationId(stationName: Station): StationId {
  return stationName === 'Олександрійська БЕСС' ? 'oleksandriya' : 'znamyanka';
}

function getReportPeriod(firstDateHeader: string) {
  return firstDateHeader.slice(0, 7);
}

function getDaysInReportPeriod(period: string) {
  if (!isReportPeriod(period)) {
    return 0;
  }

  const [year, month] = period.split('-').map(Number);
  return new Date(year, month, 0).getDate();
}

function isReportPeriod(value: string): value is ReportPeriod {
  return /^\d{4}-\d{2}$/.test(value);
}

function getPreviousPeriod(period: string) {
  if (!isReportPeriod(period)) {
    return period;
  }

  const [year, month] = period.split('-').map(Number);
  const previousMonthDate = new Date(year, month - 2, 1);

  return `${previousMonthDate.getFullYear()}-${String(previousMonthDate.getMonth() + 1).padStart(2, '0')}`;
}

function parseMoney(value: string) {
  return Number(value.replace(/\s/g, '').replace(',', '.'));
}

function roundMoney(value: number) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function getStationFileMismatchWarning(stationName: Station, uploadedFileName: string) {
  const normalizedFileName = uploadedFileName.toUpperCase();
  if (!normalizedFileName) {
    return '';
  }

  const hasZnamyankaMarker = /ZNAMENSK|ZNAMYANKA|ЗНАМЕН/.test(normalizedFileName);
  const hasOleksandriyaMarker = /OLEKSAND|OLEKSANDRIYA|ОЛЕКСАНДР/.test(normalizedFileName);

  if (stationName === 'Олександрійська БЕСС' && hasZnamyankaMarker) {
    return 'Возможно, выбранная станция не соответствует загруженному файлу. Проверьте перед расчетом.';
  }

  if (stationName === 'Знаменська БЕСС' && hasOleksandriyaMarker) {
    return 'Возможно, выбранная станция не соответствует загруженному файлу. Проверьте перед расчетом.';
  }

  return '';
}

function getFirstDateHeaderFromTable0Module(table0Fcr: ProjectReportState['periods'][ReportPeriod]['stations'][StationId]['table0Fcr']) {
  const debug = table0Fcr.parsedData?.debug;
  if (debug && typeof debug === 'object' && 'firstDateHeader' in debug && typeof debug.firstDateHeader === 'string') {
    return debug.firstDateHeader;
  }

  return '';
}

function isSameTable0Result(first: PaymentCalculation, second: PaymentCalculation | NonNullable<ProjectReportState['periods'][ReportPeriod]['stations'][StationId]['table0Fcr']['result']>) {
  return first.certifiedPowerMw === second.certifiedPowerMw
    && first.trueHours === second.trueHours
    && first.falseHours === second.falseHours
    && first.serviceVolume === second.serviceVolume
    && first.fcrTariffEur === second.fcrTariffEur
    && first.eurRate === second.eurRate
    && first.monthlyPriceUah === second.monthlyPriceUah
    && first.costWithoutVat === second.costWithoutVat
    && first.vat === second.vat
    && first.costWithVat === second.costWithVat;
}

function readPersistedState(): PersistedAppState | null {
  const parsedValue = loadDraftState<Partial<PersistedAppState>>();
  if (!parsedValue) {
    return null;
  }

  return {
    station: isStation(parsedValue.station) ? parsedValue.station : 'Олександрійська БЕСС',
    eurRate: typeof parsedValue.eurRate === 'string' ? parsedValue.eurRate : '',
    fileName: typeof parsedValue.fileName === 'string' ? parsedValue.fileName : '',
    result: parsedValue.result ?? null,
    paymentCalculation: parsedValue.paymentCalculation ?? null,
  };
}

export function App() {
  const [persistedState] = useState<PersistedAppState | null>(() => readPersistedState());
  const [station, setStation] = useState<Station>(persistedState?.station ?? 'Олександрійська БЕСС');
  const [eurRate, setEurRate] = useState(persistedState?.eurRate ?? '');
  const [fileName, setFileName] = useState(persistedState?.fileName ?? '');
  const [filePath, setFilePath] = useState('');
  const [browserFile, setBrowserFile] = useState<File | null>(null);
  const [result, setResult] = useState<FcrMonitoringResult | null>(persistedState?.result ?? null);
  const [errorMessage, setErrorMessage] = useState('');
  const [exportStatus, setExportStatus] = useState<ExportStatus | null>(null);
  const [exportErrorMessage, setExportErrorMessage] = useState('');
  const [isExporting, setIsExporting] = useState(false);
  const [isCalculating, setIsCalculating] = useState(false);
  const [projectMemoryState, setProjectMemoryState] = useState<ProjectReportState | null>(() => loadReportState());
  const [table1Period, setTable1Period] = useState<ReportPeriod>(
    (loadReportState()?.activePeriod ?? persistedState?.result?.debug.firstDateHeader?.slice(0, 7) ?? '2026-04') as ReportPeriod,
  );
  const [table1Station, setTable1Station] = useState<Station>('Олександрійська БЕСС');
  const [paymentDate, setPaymentDate] = useState('');
  const [paymentForPeriod, setPaymentForPeriod] = useState<ReportPeriod>(table1Period);
  const [paymentAmount, setPaymentAmount] = useState('');
  const [table1Message, setTable1Message] = useState('');
  const [table1ExportStatus, setTable1ExportStatus] = useState<Table1ExportStatus | null>(null);
  const [table1ExportError, setTable1ExportError] = useState('');
  const [isTable1Exporting, setIsTable1Exporting] = useState(false);
  const [activeFcrSubTab, setActiveFcrSubTab] = useState<FcrSubTab>('table0');
  const [showProjectMemoryDetails, setShowProjectMemoryDetails] = useState(false);
  const [showFcrDebugDetails, setShowFcrDebugDetails] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const normalizedRate = parseRate(eurRate);
  const hasValidRate = Number.isFinite(normalizedRate) && normalizedRate > 0;

  const canCalculate = useMemo(() => {
    return station && fileName.length > 0 && hasValidRate && !isCalculating;
  }, [fileName, hasValidRate, isCalculating, station]);

  const paymentCalculation = useMemo<PaymentCalculation | null>(() => {
    if (!result || !hasValidRate) {
      return null;
    }

    const config = stationConfig[station];
    const serviceVolume = config.certifiedPowerMw * result.trueHours;
    const monthlyPriceUah = roundToTwoDecimals(config.fcrTariffEur * normalizedRate);
    const costWithoutVat = serviceVolume * monthlyPriceUah;
    const vat = costWithoutVat * 0.2;
    const costWithVat = costWithoutVat + vat;

    return {
      station,
      certifiedPowerMw: config.certifiedPowerMw,
      trueHours: result.trueHours,
      falseHours: result.falseHours,
      serviceVolume,
      fcrTariffEur: config.fcrTariffEur,
      eurRate,
      monthlyPriceUah,
      costWithoutVat,
      vat,
      costWithVat,
    };
  }, [eurRate, hasValidRate, normalizedRate, result, station]);

  const stationFileMismatchWarning = getStationFileMismatchWarning(station, fileName);
  const memoryPeriods = projectMemoryState ? (Object.keys(projectMemoryState.periods).sort() as ReportPeriod[]) : [];
  const table1StationId = getStationId(table1Station);
  const table1StationState = projectMemoryState?.periods[table1Period]?.stations[table1StationId] ?? null;
  const table1Table0Result = table1StationState?.table0Fcr.result ?? null;
  const table1ExistingModule = table1StationState?.table1Payments ?? null;
  const table1Payments = table1ExistingModule?.manualInputs.payments ?? [];
  const table1MatchingPayments = table1Payments.filter((payment) => payment.forPeriod === table1Period);
  const table1PaidAmount = roundMoney(table1MatchingPayments.reduce((sum, payment) => sum + payment.amountUah, 0));
  const table1AccruedAmount = table1Table0Result?.costWithVat ?? 0;
  const table1DebtAmount = table1Table0Result ? Math.max(0, roundMoney(table1AccruedAmount - table1PaidAmount)) : 0;
  const table1PayoutPercent = table1Table0Result && table1AccruedAmount > 0 ? table1PaidAmount / table1AccruedAmount : 0;
  const canExportTable1 = Boolean(
    projectMemoryState && memoryPeriods.some((period) => stationMemoryOrder.some((stationId) => projectMemoryState.periods[period].stations[stationId].table0Fcr.result)),
  );

  const savedTable0ExportRecords = useMemo<BrowserTable0ExportRecord[]>(() => {
    const state = projectMemoryState;
    const stationId = getStationId(station);
    if (!state) {
      return [];
    }

    return (Object.keys(state.periods).sort() as ReportPeriod[]).flatMap((period) => {
      const table0Fcr = state.periods[period]?.stations[stationId]?.table0Fcr;
      const table0Result = table0Fcr?.result;
      const firstDateHeader = table0Fcr ? getFirstDateHeaderFromTable0Module(table0Fcr) : '';
      if (!table0Result || !firstDateHeader) {
        return [];
      }

      return [{
        station,
        firstDateHeader,
        certifiedPowerMw: table0Result.certifiedPowerMw,
        trueHours: table0Result.trueHours,
        falseHours: table0Result.falseHours,
        serviceVolume: table0Result.serviceVolume,
        fcrTariffEur: table0Result.fcrTariffEur,
        eurRate: table0Result.eurRate,
        monthlyPriceUah: table0Result.monthlyPriceUah,
        costWithoutVat: table0Result.costWithoutVat,
        vat: table0Result.vat,
        costWithVat: table0Result.costWithVat,
      }];
    });
  }, [projectMemoryState, station]);
  const currentDraftPeriod = result?.debug.firstDateHeader ? (getReportPeriod(result.debug.firstDateHeader) as ReportPeriod) : null;
  const currentDraftSavedResult = currentDraftPeriod && paymentCalculation
    ? projectMemoryState?.periods[currentDraftPeriod]?.stations[getStationId(paymentCalculation.station)]?.table0Fcr.result
    : null;
  const isCurrentDraftSavedToProject = Boolean(paymentCalculation && currentDraftSavedResult && isSameTable0Result(paymentCalculation, currentDraftSavedResult));
  const draftStatusText = paymentCalculation
    ? isCurrentDraftSavedToProject
      ? 'Сохранено в месячный отчет'
      : 'Черновик расчета, не сохранен в месячный отчет'
    : '';
  const expectedHoursInDraftMonth = currentDraftPeriod ? getDaysInReportPeriod(currentDraftPeriod) * 24 : 0;
  const hasIncompleteFcrPeriod = Boolean(
    result
      && expectedHoursInDraftMonth
      && (result.debug.dateColumnsWithValues < getDaysInReportPeriod(currentDraftPeriod ?? '') || result.totalHours < expectedHoursInDraftMonth),
  );

  useEffect(() => {
    const isEmptyDefaultState =
      station === 'Олександрійська БЕСС' && eurRate === '' && fileName === '' && !result && !paymentCalculation;

    if (isEmptyDefaultState) {
      clearDraftState();
      return;
    }

    const stateToPersist: PersistedAppState = {
      station,
      eurRate,
      fileName,
      result,
      paymentCalculation,
    };

    saveDraftState(stateToPersist);
  }, [eurRate, fileName, paymentCalculation, result, station]);

  const monthWarning =
    result && hasIncompleteFcrPeriod
      ? `Файл содержит неполный период. Расчет можно использовать для анализа, но перед сохранением в месячный отчет проверьте данные.${
          result.debug.firstDateHeader && result.debug.lastDateHeader
            ? ` Диапазон дат: ${result.debug.firstDateHeader} — ${result.debug.lastDateHeader}.`
            : ''
        }`
      : '';

  async function handlePickFile() {
    setResult(null);
    setErrorMessage('');
    setExportStatus(null);
    setExportErrorMessage('');

    if (window.uzeApp?.openFcrFile) {
      const selectedFilePath = await window.uzeApp.openFcrFile();
      if (selectedFilePath) {
        setFilePath(selectedFilePath);
        setBrowserFile(null);
        setFileName(selectedFilePath.split(/[\\/]/).pop() ?? selectedFilePath);
      }
      return;
    }

    fileInputRef.current?.click();
  }

  function handleBrowserFileChange(file: File | undefined) {
    setResult(null);
    setErrorMessage('');
    setExportStatus(null);
    setExportErrorMessage('');
    setFilePath('');
    setBrowserFile(file ?? null);
    setFileName(file?.name ?? '');
  }

  function handleClearData() {
    clearDraftState();
    window.localStorage.removeItem(projectReportStateStorageKey);
    setProjectMemoryState(null);
    setStation('Олександрійська БЕСС');
    setEurRate('');
    setFileName('');
    setFilePath('');
    setBrowserFile(null);
    setResult(null);
    setErrorMessage('');
    setExportStatus(null);
    setExportErrorMessage('');
    setTable1ExportStatus(null);
    setTable1ExportError('');
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  }

  function handleRefreshProjectMemory() {
    setProjectMemoryState(loadReportState());
  }

  function handleSaveTable0ToProject() {
    if (!result || !paymentCalculation || !result.debug.firstDateHeader) {
      setExportErrorMessage('Сначала выполните расчет Таблицы_0.');
      return;
    }

    const reportPeriod = getReportPeriod(result.debug.firstDateHeader) as ReportPeriod;
    const updatedModule = updateStationModule(reportPeriod, getStationId(paymentCalculation.station), 'table0Fcr', {
      manualInputs: {
        stationName: paymentCalculation.station,
        eurRate,
        certifiedPowerMw: paymentCalculation.certifiedPowerMw,
        fcrTariffEur: paymentCalculation.fcrTariffEur,
      },
      uploadedFiles: fileName
        ? [
            {
              fileName,
              source: filePath ? 'electron' : 'browser',
              path: filePath || undefined,
              uploadedAt: new Date().toISOString(),
            },
          ]
        : [],
      parsedData: {
        trueHours: result.trueHours,
        falseHours: result.falseHours,
        totalHours: result.totalHours,
        debug: result.debug,
      },
      result: paymentCalculation,
      validationErrors: errorMessage
        ? [
            {
              moduleName: 'table0Fcr',
              message: errorMessage,
              createdAt: new Date().toISOString(),
            },
          ]
        : [],
    });

    if (updatedModule) {
      setProjectMemoryState(loadReportState());
      setTable1Period(reportPeriod);
      setPaymentForPeriod(reportPeriod);
      setExportErrorMessage('');
    }
  }

  async function handleExportTable1() {
    setTable1ExportStatus(null);
    setTable1ExportError('');
    const state = loadReportState();

    if (!state) {
      setTable1ExportError('В памяти проекта нет данных для экспорта Таблицы_1.');
      return;
    }

    setIsTable1Exporting(true);
    try {
      const exportInput = {
        state,
        exportPeriod: table1Period,
      };
      const exportResult = window.uzeApp?.exportTable1Ukrenergo
        ? await window.uzeApp.exportTable1Ukrenergo(exportInput)
        : await exportTable1InBrowser(exportInput);
      if (!exportResult) {
        setTable1ExportError('Экспорт отменен.');
        return;
      }
      setProjectMemoryState(loadReportState());
      setTable1ExportStatus({
        ...exportResult,
        exportPeriod: exportResult.exportPeriod as ReportPeriod,
        fileName: 'outputPath' in exportResult ? exportResult.outputPath.split(/[\\/]/).pop() || exportResult.fileName : exportResult.fileName,
        updatedStationRows: exportResult.updatedStationRows.map((row) => ({
          ...row,
          period: row.period as ReportPeriod,
        })),
      });
    } catch (error) {
      setTable1ExportError(error instanceof Error ? error.message : 'Не удалось экспортировать Таблицу_1.');
    } finally {
      setIsTable1Exporting(false);
    }
  }

  function saveTable1Payments(nextPayments: Table1PaymentRecord[]) {
    if (!table1Table0Result) {
      setTable1Message('Сначала рассчитайте и сохраните Таблицу_0 для этой станции и периода.');
      return;
    }

    const paidAmount = roundMoney(nextPayments
      .filter((payment) => payment.forPeriod === table1Period)
      .reduce((sum, payment) => sum + payment.amountUah, 0));
    const debtAmount = Math.max(0, roundMoney(table1Table0Result.costWithVat - paidAmount));
    const payoutPercent = table1Table0Result.costWithVat > 0 ? paidAmount / table1Table0Result.costWithVat : 0;

    const updatedState = updateStationModule(table1Period, table1StationId, 'table1Payments', {
      manualInputs: {
        stationName: table1Station,
        accrualPeriod: table1Period,
        payments: nextPayments,
      },
      parsedData: null,
      result: {
        period: table1Period,
        stationId: table1StationId,
        stationName: table1Station,
        serviceVolume: table1Table0Result.serviceVolume,
        averagePriceUah: table1Table0Result.monthlyPriceUah,
        accruedWithVat: table1Table0Result.costWithVat,
        paidAmount,
        debtAmount,
        payoutPercent,
        payments: nextPayments.filter((payment) => payment.forPeriod === table1Period),
      },
      validationErrors: [],
    });

    if (updatedState) {
      setProjectMemoryState(loadReportState());
    }
    setTable1Message('Таблица_1 сохранена в памяти проекта.');
  }

  function handleSaveTable1WithoutPayment() {
    saveTable1Payments(table1Payments);
  }

  function handleAddTable1Payment() {
    setTable1Message('');

    if (!paymentDate || !paymentForPeriod || !paymentAmount.trim()) {
      setTable1Message('Заполните дату оплаты, месяц оплаты и сумму.');
      return;
    }

    if (!isReportPeriod(paymentForPeriod)) {
      setTable1Message('Поле "За який місяць" должно быть в формате YYYY-MM.');
      return;
    }

    const amountUah = parseMoney(paymentAmount);
    if (!Number.isFinite(amountUah) || amountUah <= 0) {
      setTable1Message('Укажите корректную сумму оплаты.');
      return;
    }

    const nextPayments: Table1PaymentRecord[] = [
      ...table1Payments,
      {
        id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
        paymentDate,
        forPeriod: paymentForPeriod,
        amountUah: roundMoney(amountUah),
        createdAt: new Date().toISOString(),
      },
    ];

    saveTable1Payments(nextPayments);
    setPaymentDate('');
    setPaymentAmount('');
  }

  async function handleCalculate() {
    if (!hasValidRate) {
      setErrorMessage('Укажите корректный средний курс EUR/UAH за месяц.');
      return;
    }

    setIsCalculating(true);
    setResult(null);
    setErrorMessage('');
    setExportStatus(null);
    setExportErrorMessage('');

    try {
      const monitoringResult =
        filePath && window.uzeApp?.calculateFcrMonitoring
          ? await window.uzeApp.calculateFcrMonitoring(filePath)
          : browserFile
            ? await calculateFcrMonitoringFromFile(browserFile)
            : null;

      if (!monitoringResult) {
        throw new Error('Сначала загрузите Excel-файл FCR_monitoring.xlsx.');
      }

      setResult(monitoringResult);
    } catch (error) {
      setErrorMessage(error instanceof Error ? error.message : 'Не удалось прочитать Excel-файл.');
    } finally {
      setIsCalculating(false);
    }
  }

  async function handleExportTable0() {
    setExportStatus(null);
    setExportErrorMessage('');

    const latestProjectState = loadReportState();
    const latestStationId = getStationId(station);
    const storedPeriods = latestProjectState
      ? (Object.keys(latestProjectState.periods).sort() as ReportPeriod[]).flatMap((period) => {
          const table0Fcr = latestProjectState.periods[period]?.stations[latestStationId]?.table0Fcr;
          const table0Result = table0Fcr?.result;
          const firstDateHeader = table0Fcr ? getFirstDateHeaderFromTable0Module(table0Fcr) : '';
          if (!table0Result || !firstDateHeader) {
            return [];
          }

          return [{
            station,
            firstDateHeader,
            certifiedPowerMw: table0Result.certifiedPowerMw,
            trueHours: table0Result.trueHours,
            falseHours: table0Result.falseHours,
            serviceVolume: table0Result.serviceVolume,
            fcrTariffEur: table0Result.fcrTariffEur,
            eurRate: table0Result.eurRate,
            monthlyPriceUah: table0Result.monthlyPriceUah,
            costWithoutVat: table0Result.costWithoutVat,
            vat: table0Result.vat,
            costWithVat: table0Result.costWithVat,
          }];
        })
      : [];
    const exportRecords = storedPeriods.length > 0 ? storedPeriods : savedTable0ExportRecords;
    const latestRecord = exportRecords.at(-1);

    if (!latestRecord) {
      setExportErrorMessage('Сначала сохраните Таблицу_0 в месячный отчет. Экспорт строится только по финальной памяти проекта.');
      return;
    }


    setIsExporting(true);

    try {
      const exportInput = {
        ...latestRecord,
        periods: exportRecords,
      };

      const exportResult = window.uzeApp?.exportTable0Rpch
        ? await window.uzeApp.exportTable0Rpch(exportInput)
        : await exportTable0RpchInBrowser(exportInput);

      if (!exportResult) {
        setExportErrorMessage('Экспорт отменен.');
        return;
      }

      setExportStatus({
        fileName: 'fileName' in exportResult ? exportResult.fileName : exportResult.outputPath.split(/[\\/]/).pop() || exportResult.outputPath,
        station,
        monthLabel: exportResult.monthLabel,
        action:
          exportResult.mode === 'updated'
            ? 'Обновлена строка'
            : exportResult.mode === 'filled-empty'
              ? 'Заполнена пустая строка'
              : 'Добавлена строка перед ВСЬОГО',
        rowNumber: exportResult.rowNumber,
        totalRowNumber: exportResult.totalRowNumber,
        totalFormulaRange: exportResult.totalFormulaRange,
        templateSource: exportResult.templateSource,
      });
    } catch (error) {
      setExportErrorMessage(error instanceof Error ? error.message : 'Не удалось экспортировать Таблицу_0.');
    } finally {
      setIsExporting(false);
    }
  }

  const projectMemoryPanel = (
    <section className="memory-panel" aria-labelledby="memory-title">
      <div className="memory-header">
        <div>
          <p className="eyebrow">LocalStorage</p>
          <h3 id="memory-title">Память проекта</h3>
        </div>
        <div className="memory-actions">
          <button className="memory-refresh-button" type="button" onClick={handleRefreshProjectMemory}>
            Обновить состояние
          </button>
          <button className="memory-details-button" type="button" onClick={() => setShowProjectMemoryDetails((value) => !value)}>
            {showProjectMemoryDetails ? 'Скрыть подробности' : 'Подробнее'}
          </button>
        </div>
      </div>

      {memoryPeriods.length === 0 ? (
        <p className="memory-empty">Сохраненных периодов пока нет. Выполните расчет Таблицы_0 для одной из станций.</p>
      ) : (
        <div className="memory-period-list">
          {memoryPeriods.map((period) => {
            const periodState = projectMemoryState?.periods[period];

            return (
              <div className="memory-period" key={period}>
                <div className="memory-period-title">
                  <span>Период</span>
                  <strong>{period}</strong>
                </div>
                <div className="memory-station-grid">
                  {stationMemoryOrder.map((stationId) => {
                    const table0Fcr = periodState?.stations[stationId]?.table0Fcr;
                    const table1PaymentsModule = periodState?.stations[stationId]?.table1Payments;
                    const table0Result = table0Fcr?.result;
                    const hasTable0 = Boolean(table0Result);
                    const hasTable1 = Boolean(table1PaymentsModule?.result);

                    return (
                      <div className={hasTable0 ? 'memory-station memory-station-saved' : 'memory-station'} key={stationId}>
                        <div className="memory-station-topline">
                          <strong>{stationLabels[stationId]}</strong>
                          <div className="memory-status-list">
                            <span>{hasTable0 ? 'Таблица_0 ✅' : 'Таблица_0 нет данных'}</span>
                            <span>{hasTable1 ? 'Таблица_1 ✅' : 'Таблица_1 нет данных'}</span>
                          </div>
                        </div>
                        {showProjectMemoryDetails && table0Result ? (
                          <div className="memory-facts">
                            <div>
                              <span>Месяц</span>
                              <strong>{period}</strong>
                            </div>
                            <div>
                              <span>Станция</span>
                              <strong>{table0Result.station ?? stationLabels[stationId]}</strong>
                            </div>
                            <div>
                              <span>Часы оказания</span>
                              <strong>{table0Result.trueHours}</strong>
                            </div>
                            <div>
                              <span>Часы неоказания</span>
                              <strong>{table0Result.falseHours}</strong>
                            </div>
                            <div>
                              <span>Сумма с НДС</span>
                              <strong>{moneyFormatter.format(table0Result.costWithVat)} грн</strong>
                            </div>
                          </div>
                        ) : showProjectMemoryDetails ? (
                          <p>Для этой станции данные Таблицы_0 еще не сохранены.</p>
                        ) : null}
                      </div>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </section>
  );

  return (
    <main className="app-shell">
      <aside className="sidebar" aria-label="Модули приложения">
        <div className="brand-block">
          <div className="brand-mark" aria-hidden="true">
            <BarChart3 size={24} />
          </div>
          <div>
            <p className="eyebrow">Локальное приложение</p>
            <h1>UZE Report Automation</h1>
          </div>
        </div>

        <nav className="module-list">
          {modules.map((module) => {
            const Icon = module.icon;
            return (
              <button
                className={module.enabled ? 'module-item module-item-active' : 'module-item'}
                disabled={!module.enabled}
                key={module.title}
                type="button"
                title={module.enabled ? module.title : `${module.title}: будет добавлено позже`}
              >
                <Icon size={18} />
                <span>
                  <strong>{module.title}</strong>
                  <small>{module.description}</small>
                </span>
                <em>{module.status}</em>
              </button>
            );
          })}
        </nav>

        <div className="sidebar-footer">
          <Settings size={17} />
          <span>Структура готова для новых модулей</span>
        </div>
      </aside>

      <section className="workspace">
        <header className="topbar">
          <div>
            <p className="eyebrow">Первый модуль</p>
            <h2>РПЧ / FCR</h2>
          </div>
          <div className="topbar-status">
            <Plug size={16} />
            <span>Чтение Excel и расчет оплаты РПЧ подключены для первого модуля</span>
          </div>
        </header>

        <div className="subtab-bar" role="tablist" aria-label="РПЧ / FCR">
          <button
            className={activeFcrSubTab === 'table0' ? 'subtab-button subtab-button-active' : 'subtab-button'}
            type="button"
            role="tab"
            aria-selected={activeFcrSubTab === 'table0'}
            onClick={() => setActiveFcrSubTab('table0')}
          >
            Таблица_0 — расчет начисления
          </button>
          <button
            className={activeFcrSubTab === 'table1' ? 'subtab-button subtab-button-active' : 'subtab-button'}
            type="button"
            role="tab"
            aria-selected={activeFcrSubTab === 'table1'}
            onClick={() => setActiveFcrSubTab('table1')}
          >
            Таблица_1 — оплата Укренерго
          </button>
        </div>

        {activeFcrSubTab === 'table0' && (
        <div className="content-grid">
          <section className="form-panel" aria-labelledby="fcr-form-title">
              <div className="section-heading">
                <FileSpreadsheet size={22} />
                <div>
                  <h3 id="fcr-form-title">Данные для расчета</h3>
                  <p>Выберите станцию, укажите курс и загрузите файл мониторинга FCR.</p>
                </div>
              </div>

              <div className="field-group">
                <label htmlFor="station">Станция</label>
                <select
                  id="station"
                  value={station}
                  onChange={(event) => {
                    setStation(event.target.value as Station);
                    setResult(null);
                    setErrorMessage('');
                    setExportStatus(null);
                    setExportErrorMessage('');
                  }}
                >
                  <option value="Олександрійська БЕСС">Олександрійська БЕСС</option>
                  <option value="Знаменська БЕСС">Знаменська БЕСС</option>
                </select>
              </div>

              <div className="field-group">
                <label htmlFor="eur-rate">Средний курс EUR/UAH за месяц</label>
                <input
                  id="eur-rate"
                  inputMode="decimal"
                  placeholder="Например, 43,20"
                  type="text"
                  value={eurRate}
                  onChange={(event) => {
                    setResult(null);
                    setErrorMessage('');
                    setEurRate(event.target.value);
                  }}
                />
              </div>

              <div className="upload-row">
                <button className="secondary-button" type="button" onClick={handlePickFile}>
                  <FolderOpen size={18} />
                  Загрузить Excel-файл
                </button>
                <input
                  accept=".xlsx"
                  ref={fileInputRef}
                  type="file"
                  hidden
                  onChange={(event) => handleBrowserFileChange(event.target.files?.[0])}
                />
                <span className={fileName ? 'file-chip' : 'file-chip file-chip-empty'}>
                  {fileName || 'FCR_monitoring.xlsx не выбран'}
                </span>
              </div>

              <div className="form-actions">
                <button className="primary-button" disabled={!canCalculate} type="button" onClick={handleCalculate}>
                  <Calculator size={19} />
                  {isCalculating ? 'Чтение файла...' : 'Рассчитать'}
                </button>
                <button className="clear-button" type="button" onClick={handleClearData}>
                  <Trash2 size={18} />
                  Очистить данные
                </button>
              </div>
          </section>

          <aside className="summary-panel" aria-label="Состояние модуля">
            <div className="summary-card">
              <p className="eyebrow">Текущий этап</p>
              <h3>Чтение FCR и расчет оплаты</h3>
              <p>
                Приложение читает лист FCR, считает значения TRUE/FALSE и рассчитывает оплату РПЧ без формирования
                итогового Excel-отчета.
              </p>
            </div>

            <div className="summary-list">
              <div>
                <span>Станция</span>
                <strong>{station}</strong>
              </div>
              <div>
                <span>Курс EUR/UAH</span>
                <strong>{eurRate || 'Не указан'}</strong>
              </div>
              <div>
                <span>Excel-файл</span>
                <strong>{fileName || 'Не выбран'}</strong>
              </div>
            </div>
          </aside>
        </div>
        )}

        <div className="results-stack">
            {activeFcrSubTab === 'table0' && stationFileMismatchWarning && (
              <section className="message-panel message-panel-warning" aria-live="polite">
                <AlertTriangle size={20} />
                <div>
                  <h3>Проверьте станцию</h3>
                  <p>{stationFileMismatchWarning}</p>
                </div>
              </section>
            )}

            {activeFcrSubTab === 'table0' && projectMemoryPanel}

            {activeFcrSubTab === 'table1' && (
            <section className="table1-panel" aria-labelledby="table1-title">
              <div className="section-heading result-heading">
                <FileText size={22} />
                <div>
                  <p className="eyebrow">Розрахунки Укренерго</p>
                  <h3 id="table1-title">Таблица_1 — оплата РПЧ</h3>
                  <p>Начисление берется из сохраненной Таблицы_0, оплаты добавляются вручную по актам Укренерго.</p>
                </div>
              </div>

              <div className="table1-controls">
                <div className="field-group">
                  <label htmlFor="table1-period">Период начисления</label>
                  <input
                    id="table1-period"
                    type="month"
                    value={table1Period}
                    onChange={(event) => {
                      const nextPeriod = event.target.value as ReportPeriod;
                      setTable1Period(nextPeriod);
                      setPaymentForPeriod(nextPeriod);
                      setTable1Message('');
                    }}
                  />
                </div>
                <div className="field-group">
                  <label htmlFor="table1-station">Станция</label>
                  <select
                    id="table1-station"
                    value={table1Station}
                    onChange={(event) => {
                      setTable1Station(event.target.value as Station);
                      setTable1Message('');
                    }}
                  >
                    <option value="Олександрійська БЕСС">Олександрійська БЕСС</option>
                    <option value="Знаменська БЕСС">Знаменська БЕСС</option>
                  </select>
                </div>
              </div>

              {!table1Table0Result ? (
                <div className="table1-empty-state">
                  Сначала рассчитайте и сохраните Таблицу_0 для этой станции и периода.
                </div>
              ) : (
                <>
                  <div className="table1-kpi-grid">
                    <div>
                      <span>ОБСЯГ ПОСЛУГ, МВт</span>
                      <strong>{numberFormatter.format(table1Table0Result.serviceVolume)}</strong>
                    </div>
                    <div>
                      <span>ЦІНА середньозважена</span>
                      <strong>{moneyFormatter.format(table1Table0Result.monthlyPriceUah)} грн</strong>
                    </div>
                    <div>
                      <span>НАРАХОВАНО</span>
                      <strong>{moneyFormatter.format(table1AccruedAmount)} грн</strong>
                    </div>
                    <div>
                      <span>ВИПЛАЧЕНО</span>
                      <strong>{moneyFormatter.format(table1PaidAmount)} грн</strong>
                    </div>
                    <div>
                      <span>ЗАБОРГОВАНІСТЬ</span>
                      <strong>{moneyFormatter.format(table1DebtAmount)} грн</strong>
                    </div>
                    <div>
                      <span>%% виплат</span>
                      <strong>{(table1PayoutPercent * 100).toFixed(2)}%</strong>
                    </div>
                  </div>

                  <div className="table1-payment-form">
                    <div className="field-group">
                      <label htmlFor="payment-date">Дата оплаты</label>
                      <input id="payment-date" type="date" value={paymentDate} onChange={(event) => setPaymentDate(event.target.value)} />
                    </div>
                    <div className="field-group">
                      <label htmlFor="payment-for-period">За який місяць</label>
                      <input
                        id="payment-for-period"
                        type="month"
                        value={paymentForPeriod}
                        onChange={(event) => setPaymentForPeriod(event.target.value as ReportPeriod)}
                      />
                    </div>
                    <div className="field-group">
                      <label htmlFor="payment-amount">Сума оплати, грн</label>
                      <input
                        id="payment-amount"
                        inputMode="decimal"
                        placeholder="Например, 6162421,68"
                        type="text"
                        value={paymentAmount}
                        onChange={(event) => setPaymentAmount(event.target.value)}
                      />
                    </div>
                    <div className="table1-actions">
                      <button className="primary-button" type="button" onClick={handleAddTable1Payment}>
                        Добавить оплату
                      </button>
                      <button className="clear-button" type="button" onClick={handleSaveTable1WithoutPayment}>
                        Сохранить расчет без оплат
                      </button>
                    </div>
                  </div>

                  <div className="table1-payments-list">
                    <h4>Оплаты по выбранной станции</h4>
                    {table1Payments.length === 0 ? (
                      <p>Оплат пока нет. ВИПЛАЧЕНО = 0, задолженность равна начислению.</p>
                    ) : (
                      <div className="table1-payments-table-wrap">
                        <table className="table1-payments-table">
                          <thead>
                            <tr>
                              <th>Дата оплаты</th>
                              <th>За який місяць</th>
                              <th>Сума оплати, грн</th>
                              <th>Учитывается сейчас</th>
                            </tr>
                          </thead>
                          <tbody>
                            {table1Payments.map((payment) => (
                              <tr key={payment.id}>
                                <td>{payment.paymentDate}</td>
                                <td>{payment.forPeriod}</td>
                                <td>{moneyFormatter.format(payment.amountUah)}</td>
                                <td>{payment.forPeriod === table1Period ? 'Да' : 'Нет'}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                </>
              )}

              <div className="table1-export-actions">
                <button className="export-button" disabled={!canExportTable1 || isTable1Exporting} type="button" onClick={handleExportTable1}>
                  <Download size={18} />
                  {isTable1Exporting ? 'Экспорт...' : 'Скачать Таблицу_1 Excel'}
                </button>
                {table1ExportError && <span className="export-status export-status-error">{table1ExportError}</span>}
              </div>

              {table1ExportStatus && (
                <div className="export-summary" aria-live="polite">
                  <h4>Статус экспорта Таблицы_1</h4>
                  <div className="export-summary-grid">
                    <div>
                      <span>Файл</span>
                      <strong>{table1ExportStatus.fileName}</strong>
                    </div>
                    <div>
                      <span>Период</span>
                      <strong>{table1ExportStatus.exportPeriod}</strong>
                    </div>
                    <div>
                      <span>Обновлено строк</span>
                      <strong>{table1ExportStatus.updatedStationRows.length}</strong>
                    </div>
                    <div className="export-summary-wide">
                      <span>Источник шаблона</span>
                      <strong>{table1ExportStatus.templateSource}</strong>
                    </div>
                  </div>
                </div>
              )}

              {table1Message && <p className="table1-message">{table1Message}</p>}
            </section>
            )}

            {activeFcrSubTab === 'table1' && projectMemoryPanel}

            {activeFcrSubTab === 'table0' && errorMessage && (
              <section className="message-panel message-panel-error" aria-live="polite">
                <AlertCircle size={20} />
                <div>
                  <h3>Ошибка расчета</h3>
                  <p>{errorMessage}</p>
                </div>
              </section>
            )}

            {activeFcrSubTab === 'table0' && monthWarning && (
              <section className="message-panel message-panel-warning" aria-live="polite">
                <AlertTriangle size={20} />
                <div>
                  <h3>Предупреждение</h3>
                  <p>{monthWarning}</p>
                </div>
              </section>
            )}

            {activeFcrSubTab === 'table0' && result && (
              <section className="result-panel" aria-labelledby="fcr-result-title">
                <div className="section-heading result-heading">
                  <BarChart3 size={22} />
                  <div>
                    <p className="eyebrow">Проверка Excel</p>
                    <h3 id="fcr-result-title">Результат мониторинга FCR</h3>
                  </div>
                </div>

                <div className="result-grid">
                  <div className="result-file">
                    <span>Файл</span>
                    <strong>{fileName}</strong>
                  </div>
                  <div>
                    <span>Часы оказания РПЧ</span>
                    <strong>{result.trueHours}</strong>
                  </div>
                  <div>
                    <span>Часы неоказания РПЧ</span>
                    <strong>{result.falseHours}</strong>
                  </div>
                  <div>
                    <span>Всего проверенных часов</span>
                    <strong>{result.totalHours}</strong>
                  </div>
                </div>

                <button className="technical-toggle-button" type="button" onClick={() => setShowFcrDebugDetails((value) => !value)}>
                  {showFcrDebugDetails ? 'Скрыть технические детали' : 'Показать технические детали'}
                </button>

                {showFcrDebugDetails && (
                <div className="debug-block">
                  <h4>Debug чтения листа FCR</h4>
                  <div className="debug-grid">
                    <div>
                      <span>Всего ячеек прочитано</span>
                      <strong>{result.debug.totalCellsRead}</strong>
                    </div>
                    <div>
                      <span>Адресованных ячеек</span>
                      <strong>{result.debug.addressedCellsRead}</strong>
                    </div>
                    <div>
                      <span>Найдено TRUE</span>
                      <strong>{result.debug.trueFound}</strong>
                    </div>
                    <div>
                      <span>Найдено FALSE</span>
                      <strong>{result.debug.falseFound}</strong>
                    </div>
                    <div>
                      <span>Итог TRUE + FALSE</span>
                      <strong>{result.debug.totalFound}</strong>
                    </div>
                    <div>
                      <span>Диапазон листа</span>
                      <strong>{result.debug.readRange}</strong>
                    </div>
                    <div>
                      <span>Диапазон TRUE/FALSE</span>
                      <strong>{result.debug.trueFalseRange}</strong>
                    </div>
                    <div>
                      <span>Дней с TRUE/FALSE</span>
                      <strong>{result.debug.dateColumnsWithValues}</strong>
                    </div>
                    <div>
                      <span>Первая дата</span>
                      <strong>{result.debug.firstDateHeader || 'Не найдена'}</strong>
                    </div>
                    <div>
                      <span>Последняя дата</span>
                      <strong>{result.debug.lastDateHeader || 'Не найдена'}</strong>
                    </div>
                  </div>
                </div>
                )}
              </section>
            )}

            {activeFcrSubTab === 'table0' && paymentCalculation && (
              <section className="payment-panel" aria-labelledby="payment-title">
                <div className="section-heading result-heading">
                  <Calculator size={22} />
                  <div>
                    <p className="eyebrow">Финансовый расчет</p>
                    <h3 id="payment-title">Таблица_0 — расчет оплаты РПЧ</h3>
                  </div>
                </div>

                <div className="payment-table-wrap">
                  <table className="payment-table">
                    <thead>
                      <tr>
                        <th>Станция</th>
                        <th>Сертифицированная мощность РПЧ, МВт</th>
                        <th>Количество часов оказания РПЧ</th>
                        <th>Количество часов неоказания РПЧ</th>
                        <th>Расчетный объем услуг, МВт*год</th>
                        <th>Аукционная цена, EUR/МВт/час</th>
                        <th>Средний курс EUR/UAH</th>
                        <th>Цена за отчетный месяц, грн</th>
                        <th>Стоимость без НДС, грн</th>
                        <th>НДС, грн</th>
                        <th>Стоимость с НДС, грн</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr>
                        <td>{paymentCalculation.station}</td>
                        <td>{paymentCalculation.certifiedPowerMw}</td>
                        <td>{paymentCalculation.trueHours}</td>
                        <td>{paymentCalculation.falseHours}</td>
                        <td>{numberFormatter.format(paymentCalculation.serviceVolume)}</td>
                        <td>{paymentCalculation.fcrTariffEur.toFixed(8)}</td>
                        <td>{paymentCalculation.eurRate}</td>
                        <td>{moneyFormatter.format(paymentCalculation.monthlyPriceUah)}</td>
                        <td>{moneyFormatter.format(paymentCalculation.costWithoutVat)}</td>
                        <td>{moneyFormatter.format(paymentCalculation.vat)}</td>
                        <td>{moneyFormatter.format(paymentCalculation.costWithVat)}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
                <div className={isCurrentDraftSavedToProject ? 'draft-status draft-status-saved' : 'draft-status draft-status-unsaved'} aria-live="polite">
                  <div>
                    <span>{isCurrentDraftSavedToProject ? 'Финальная память' : 'Черновик'}</span>
                    <strong>{draftStatusText}</strong>
                    {hasIncompleteFcrPeriod && (
                      <p>Файл содержит неполный период. Расчет можно использовать для анализа, но перед сохранением в месячный отчет проверьте данные.</p>
                    )}
                  </div>
                  <button className="primary-button draft-save-button" type="button" onClick={handleSaveTable0ToProject}>
                    Сохранить в месячный отчет
                  </button>
                </div>
                <div className="export-actions">
                  <button className="export-button" disabled={isExporting} type="button" onClick={handleExportTable0}>
                    <Download size={18} />
                    {isExporting ? 'Экспорт...' : 'Экспорт Таблицы_0 в Excel'}
                  </button>
                  {exportErrorMessage && <span className="export-status export-status-error">{exportErrorMessage}</span>}
                </div>
                {exportStatus && (
                  <div className="export-summary" aria-live="polite">
                    <h4>Статус экспорта</h4>
                    <div className="export-summary-grid">
                      <div>
                        <span>Файл</span>
                        <strong>{exportStatus.fileName}</strong>
                      </div>
                      <div>
                        <span>Станция</span>
                        <strong>{exportStatus.station}</strong>
                      </div>
                      <div>
                        <span>Месяц</span>
                        <strong>{exportStatus.monthLabel}</strong>
                      </div>
                      <div>
                        <span>Действие</span>
                        <strong>{exportStatus.action}</strong>
                      </div>
                      <div>
                        <span>Строка Excel</span>
                        <strong>{exportStatus.rowNumber}</strong>
                      </div>
                      <div>
                        <span>Строка ВСЬОГО</span>
                        <strong>{exportStatus.totalRowNumber}</strong>
                      </div>
                      <div>
                        <span>Диапазон формул ВСЬОГО</span>
                        <strong>{exportStatus.totalFormulaRange}</strong>
                      </div>
                      <div className="export-summary-wide">
                        <span>Источник шаблона</span>
                        <strong>{exportStatus.templateSource}</strong>
                      </div>
                    </div>
                  </div>
                )}
              </section>
            )}
        </div>
      </section>
    </main>
  );
}
