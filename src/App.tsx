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
import { buildBalancingEnergyDraft, parseBalancingEnergyFile, type BalancingEnergyDraftResult } from './table2BalancingEnergy';
import { buildDataHubDraft, parseDataHubFile, type DataHubDraftResult } from './table2DataHub';
import { buildTable2FinalReportData, exportTable2FinalReportInBrowser, type FinalReportExportResult } from './table2FinalReport';
import { parseMarketPricesFile, type MarketPricesDraftResult } from './table2MarketPrices';
import { buildMmsDraft, parseMmsFile, type MmsDraftResult } from './table2Mms';
import {
  buildHourlyImbalanceDetailRows,
  calculateImbalancesDraft,
  exportHourlyImbalancesToExcel,
  type HourlyImbalanceDetailRow,
  type HourlyImbalanceStationFilter,
  type ImbalancesDraftResult,
} from './table2Imbalances';
import { buildRdnVdrDraft, parseRdnVdrFile, type RdnVdrDraftResult, type RdnVdrMarket } from './table2RdnVdr';
import {
  deleteReportPeriod,
  loadReportState,
  updatePeriodMarketPrices,
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
  table2RdnVdrDraft?: RdnVdrDraftResult | null;
  table2DataHubDraft?: DataHubDraftResult | null;
  table2MarketPricesDraft?: MarketPricesDraftResult | null;
  table2MmsDraft?: MmsDraftResult | null;
  table2BalancingEnergyDraft?: BalancingEnergyDraftResult | null;
  table2ImbalancesDraft?: ImbalancesDraftResult | null;
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

type AppModule = 'fcr' | 'rdnVdr' | 'datahub' | 'imbalances' | 'finalReport' | 'projectState' | 'service';
type FcrSubTab = 'table0' | 'table1';
type ImbalancesSubTab = 'prices' | 'balancingEnergy' | 'mms' | 'calculation' | 'hourly';
type Table2FileKey = 'oleksandriya-rdn' | 'oleksandriya-vdr' | 'znamyanka-rdn' | 'znamyanka-vdr';
type DataHubFileKey = 'oleksandriya-datahub' | 'znamyanka-datahub';
type MmsFileKey = 'oleksandriya-mms' | 'znamyanka-mms';
type BalancingEnergyFileKey = 'oleksandriya-balancing-energy' | 'znamyanka-balancing-energy';

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
const table2FileSlots: Array<{ key: Table2FileKey; stationId: StationId; stationName: Station; market: RdnVdrMarket; label: string }> = [
  { key: 'oleksandriya-rdn', stationId: 'oleksandriya', stationName: 'Олександрійська БЕСС', market: 'РДН', label: 'Олександрія РДН' },
  { key: 'oleksandriya-vdr', stationId: 'oleksandriya', stationName: 'Олександрійська БЕСС', market: 'ВДР', label: 'Олександрія ВДР' },
  { key: 'znamyanka-rdn', stationId: 'znamyanka', stationName: 'Знаменська БЕСС', market: 'РДН', label: 'Знаменка РДН' },
  { key: 'znamyanka-vdr', stationId: 'znamyanka', stationName: 'Знаменська БЕСС', market: 'ВДР', label: 'Знаменка ВДР' },
];

const dataHubFileSlots: Array<{ key: DataHubFileKey; stationId: StationId; stationName: Station; label: string }> = [
  { key: 'oleksandriya-datahub', stationId: 'oleksandriya', stationName: stationLabels.oleksandriya, label: 'DataHub Олександрія' },
  { key: 'znamyanka-datahub', stationId: 'znamyanka', stationName: stationLabels.znamyanka, label: 'DataHub Знаменка' },
];

const mmsFileSlots: Array<{ key: MmsFileKey; stationId: StationId; stationName: Station; label: string }> = [
  { key: 'oleksandriya-mms', stationId: 'oleksandriya', stationName: stationLabels.oleksandriya, label: 'MMS Олександрія' },
  { key: 'znamyanka-mms', stationId: 'znamyanka', stationName: stationLabels.znamyanka, label: 'MMS Знаменка' },
];

const balancingEnergyFileSlots: Array<{ key: BalancingEnergyFileKey; stationId: StationId; stationName: Station; label: string }> = [
  { key: 'oleksandriya-balancing-energy', stationId: 'oleksandriya', stationName: stationLabels.oleksandriya, label: 'Балансирующая энергия Олександрія' },
  { key: 'znamyanka-balancing-energy', stationId: 'znamyanka', stationName: stationLabels.znamyanka, label: 'Балансирующая энергия Знаменка' },
];

const modules: Array<{ id: AppModule; title: string; description: string; status: string; icon: typeof Gauge; enabled: boolean }> = [
  { id: 'fcr', title: 'РПЧ / FCR', description: 'Таблица_0 и Таблица_1', status: 'Активно', icon: Gauge, enabled: true },
  { id: 'rdnVdr', title: 'РДН / ВДР', description: 'Покупка и продажа', status: 'Активно', icon: LineChart, enabled: true },
  { id: 'datahub', title: 'DataHub', description: 'Фактическая энергия', status: 'Активно', icon: Database, enabled: true },
  { id: 'imbalances', title: 'Небалансы', description: 'Цены, MMS, расчеты', status: 'Активно', icon: Zap, enabled: true },
  { id: 'finalReport', title: 'Итоговый отчет', description: 'Таблица_2', status: 'Активно', icon: FileText, enabled: true },
  { id: 'projectState', title: 'Состояние проекта', description: 'Проверка данных', status: 'Готово', icon: Database, enabled: true },
  { id: 'service', title: 'Сервис', description: 'Очистка и сброс', status: 'Готово', icon: Settings, enabled: true },
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

function formatOptionalNumber(value: number | null | undefined, suffix = '') {
  return value === null || value === undefined ? '—' : `${numberFormatter.format(value)}${suffix}`;
}

function formatOptionalMoney(value: number | null | undefined) {
  return value === null || value === undefined ? '—' : `${moneyFormatter.format(value)} грн`;
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
    table2RdnVdrDraft: parsedValue.table2RdnVdrDraft ?? null,
    table2DataHubDraft: parsedValue.table2DataHubDraft ?? null,
    table2MarketPricesDraft: parsedValue.table2MarketPricesDraft ?? null,
    table2MmsDraft: parsedValue.table2MmsDraft ?? null,
    table2BalancingEnergyDraft: parsedValue.table2BalancingEnergyDraft ?? null,
    table2ImbalancesDraft: parsedValue.table2ImbalancesDraft ?? null,
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
  const [activeModule, setActiveModule] = useState<AppModule>('fcr');
  const [activeFcrSubTab, setActiveFcrSubTab] = useState<FcrSubTab>('table0');
  const [activeImbalancesSubTab, setActiveImbalancesSubTab] = useState<ImbalancesSubTab>('prices');
  const [showProjectMemoryDetails, setShowProjectMemoryDetails] = useState(false);
  const [showFcrDebugDetails, setShowFcrDebugDetails] = useState(false);
  const [showMmsKnessDiagnostics, setShowMmsKnessDiagnostics] = useState(false);
  const [showMmsDirectionDiagnostics, setShowMmsDirectionDiagnostics] = useState(false);
  const [showFinalReportWarnings, setShowFinalReportWarnings] = useState(false);
  const [cleanupPeriod, setCleanupPeriod] = useState<ReportPeriod>((loadReportState()?.activePeriod ?? '2026-04') as ReportPeriod);
  const [cleanupMessage, setCleanupMessage] = useState('');
  const [showResetReportConfirm, setShowResetReportConfirm] = useState(false);
  const [table2Files, setTable2Files] = useState<Partial<Record<Table2FileKey, File>>>({});
  const [table2Draft, setTable2Draft] = useState<RdnVdrDraftResult | null>(persistedState?.table2RdnVdrDraft ?? null);
  const [table2Message, setTable2Message] = useState('');
  const [table2Error, setTable2Error] = useState('');
  const [isTable2Calculating, setIsTable2Calculating] = useState(false);
  const [dataHubFiles, setDataHubFiles] = useState<Partial<Record<DataHubFileKey, File>>>({});
  const [dataHubDraft, setDataHubDraft] = useState<DataHubDraftResult | null>(persistedState?.table2DataHubDraft ?? null);
  const [dataHubMessage, setDataHubMessage] = useState('');
  const [dataHubError, setDataHubError] = useState('');
  const [isDataHubCalculating, setIsDataHubCalculating] = useState(false);
  const [marketPricesFile, setMarketPricesFile] = useState<File | null>(null);
  const [marketPricesDraft, setMarketPricesDraft] = useState<MarketPricesDraftResult | null>(persistedState?.table2MarketPricesDraft ?? null);
  const [marketPricesMessage, setMarketPricesMessage] = useState('');
  const [marketPricesError, setMarketPricesError] = useState('');
  const [isMarketPricesCalculating, setIsMarketPricesCalculating] = useState(false);
  const [mmsFiles, setMmsFiles] = useState<Partial<Record<MmsFileKey, File>>>({});
  const [mmsDraft, setMmsDraft] = useState<MmsDraftResult | null>(persistedState?.table2MmsDraft ?? null);
  const [mmsMessage, setMmsMessage] = useState('');
  const [mmsError, setMmsError] = useState('');
  const [isMmsCalculating, setIsMmsCalculating] = useState(false);
  const [balancingEnergyFiles, setBalancingEnergyFiles] = useState<Partial<Record<BalancingEnergyFileKey, File>>>({});
  const [balancingEnergyDraft, setBalancingEnergyDraft] = useState<BalancingEnergyDraftResult | null>(persistedState?.table2BalancingEnergyDraft ?? null);
  const [balancingEnergyMessage, setBalancingEnergyMessage] = useState('');
  const [balancingEnergyError, setBalancingEnergyError] = useState('');
  const [isBalancingEnergyCalculating, setIsBalancingEnergyCalculating] = useState(false);
  const [imbalancesPeriod, setImbalancesPeriod] = useState<ReportPeriod>((loadReportState()?.activePeriod ?? '2026-04') as ReportPeriod);
  const [imbalancesDraft, setImbalancesDraft] = useState<ImbalancesDraftResult | null>(persistedState?.table2ImbalancesDraft ?? null);
  const [imbalancesMessage, setImbalancesMessage] = useState('');
  const [imbalancesError, setImbalancesError] = useState('');
  const [showHourlyImbalanceDetails, setShowHourlyImbalanceDetails] = useState(false);
  const [hourlyImbalanceStationFilter, setHourlyImbalanceStationFilter] = useState<HourlyImbalanceStationFilter>('all');
  const [hourlyImbalanceRows, setHourlyImbalanceRows] = useState<HourlyImbalanceDetailRow[]>([]);
  const [hourlyImbalanceMessage, setHourlyImbalanceMessage] = useState('');
  const [hourlyImbalanceError, setHourlyImbalanceError] = useState('');
  const [isHourlyImbalanceExporting, setIsHourlyImbalanceExporting] = useState(false);
  const [finalReportPeriod, setFinalReportPeriod] = useState<ReportPeriod>((loadReportState()?.activePeriod ?? '2026-04') as ReportPeriod);
  const [finalReportExportStatus, setFinalReportExportStatus] = useState<FinalReportExportResult | null>(null);
  const [finalReportExportError, setFinalReportExportError] = useState('');
  const [isFinalReportExporting, setIsFinalReportExporting] = useState(false);
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
  const table1StationState = projectMemoryState?.periods[table1Period]?.stations?.[table1StationId] ?? null;
  const table1Table0Result = table1StationState?.table0Fcr?.result ?? null;
  const table1ExistingModule = table1StationState?.table1Payments ?? null;
  const table1Payments = table1ExistingModule?.manualInputs.payments ?? [];
  const table1MatchingPayments = table1Payments.filter((payment) => payment.forPeriod === table1Period);
  const table1PaidAmount = roundMoney(table1MatchingPayments.reduce((sum, payment) => sum + payment.amountUah, 0));
  const table1AccruedAmount = table1Table0Result?.costWithVat ?? 0;
  const table1DebtAmount = table1Table0Result ? Math.max(0, roundMoney(table1AccruedAmount - table1PaidAmount)) : 0;
  const table1PayoutPercent = table1Table0Result && table1AccruedAmount > 0 ? table1PaidAmount / table1AccruedAmount : 0;
  const canExportTable1 = Boolean(
    projectMemoryState && memoryPeriods.some((period) => stationMemoryOrder.some((stationId) => projectMemoryState.periods[period]?.stations?.[stationId]?.table0Fcr?.result)),
  );

  const savedTable0ExportRecords = useMemo<BrowserTable0ExportRecord[]>(() => {
    const state = projectMemoryState;
    const stationId = getStationId(station);
    if (!state) {
      return [];
    }

    return (Object.keys(state.periods).sort() as ReportPeriod[]).flatMap((period) => {
      const table0Fcr = state.periods[period]?.stations?.[stationId]?.table0Fcr;
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
    ? projectMemoryState?.periods[currentDraftPeriod]?.stations?.[getStationId(paymentCalculation.station)]?.table0Fcr?.result
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
  const canCalculateTable2 = table2FileSlots.every((slot) => Boolean(table2Files[slot.key])) && !isTable2Calculating;
  const isTable2DraftSavedToProject = Boolean(
    table2Draft && stationMemoryOrder.every((stationId) => {
      const saved = projectMemoryState?.periods[table2Draft.period]?.stations?.[stationId]?.rdnVdr?.result;
      const stationDraft = table2Draft.stations?.[stationId];
      return saved
        && stationDraft
        && saved.markets.rdn.tradingResultUah === stationDraft.rdn.tradingResultUah
        && saved.markets.vdr.tradingResultUah === stationDraft.vdr.tradingResultUah;
    }),
  );
  const table2Rows = table2Draft
    ? stationMemoryOrder.flatMap((stationId) => {
        const stationDraft = table2Draft.stations?.[stationId];
        return stationDraft?.rdn && stationDraft?.vdr
          ? [
              { stationName: stationDraft.stationName, market: 'РДН' as const, result: stationDraft.rdn },
              { stationName: stationDraft.stationName, market: 'ВДР' as const, result: stationDraft.vdr },
            ]
          : [];
      })
    : [];

  const canCalculateDataHub = dataHubFileSlots.every((slot) => Boolean(dataHubFiles[slot.key])) && !isDataHubCalculating;
  const isDataHubDraftSavedToProject = Boolean(
    dataHubDraft && stationMemoryOrder.every((stationId) => {
      const saved = projectMemoryState?.periods[dataHubDraft.period]?.stations?.[stationId]?.datahub?.result;
      const stationDraft = dataHubDraft.stations?.[stationId];
      return saved
        && stationDraft
        && saved.totalInKwh === stationDraft.totalInKwh
        && saved.totalOutKwh === stationDraft.totalOutKwh;
    }),
  );
  const dataHubRows = dataHubDraft ? stationMemoryOrder.flatMap((stationId) => dataHubDraft.stations?.[stationId] ?? []) : [];
  const canCalculateMarketPrices = Boolean(marketPricesFile) && !isMarketPricesCalculating;
  const isMarketPricesDraftSavedToProject = Boolean(
    marketPricesDraft
      && projectMemoryState?.periods[marketPricesDraft.period]?.marketPrices?.result?.rowsCount === marketPricesDraft.rowsCount
      && projectMemoryState?.periods[marketPricesDraft.period]?.marketPrices?.result?.averageRdnPriceUah === marketPricesDraft.averageRdnPriceUah,
  );
  const canCalculateMms = mmsFileSlots.every((slot) => Boolean(mmsFiles[slot.key])) && !isMmsCalculating;
  const isMmsDraftSavedToProject = Boolean(
    mmsDraft && stationMemoryOrder.every((stationId) => {
      const saved = projectMemoryState?.periods[mmsDraft.period]?.stations?.[stationId]?.mms?.result;
      const stationDraft = mmsDraft.stations?.[stationId];
      return saved
        && stationDraft
        && saved.knessToStationMwh === stationDraft.knessToStationMwh
        && saved.stationToKnessMwh === stationDraft.stationToKnessMwh;
    }),
  );
  const mmsRows = mmsDraft ? stationMemoryOrder.flatMap((stationId) => mmsDraft.stations?.[stationId] ?? []) : [];
  const mmsDirectionRows = mmsRows.flatMap((stationRow) => (
    stationRow.directionDiagnostics ?? []
  ).map((direction) => ({
    stationId: stationRow.stationId,
    stationName: stationRow.stationName,
    ...direction,
  })));
  const mmsKnessColumnRows = mmsRows.flatMap((stationRow) => (
    stationRow.knessColumnDiagnostics ?? []
  ).map((column) => ({
    stationId: stationRow.stationId,
    stationName: stationRow.stationName,
    sourceFileName: stationRow.fileName,
    period: stationRow.period,
    rowsRead: stationRow.rowsRead,
    ...column,
  })));
  const canCalculateBalancingEnergy = balancingEnergyFileSlots.some((slot) => Boolean(balancingEnergyFiles[slot.key])) && !isBalancingEnergyCalculating;
  const balancingEnergyRows = balancingEnergyDraft
    ? stationMemoryOrder.flatMap((stationId) => balancingEnergyDraft.stations?.[stationId] ?? [])
    : [];
  const isBalancingEnergyDraftSavedToProject = Boolean(
    balancingEnergyDraft && balancingEnergyRows.length > 0 && balancingEnergyRows.every((stationDraft) => {
      const saved = projectMemoryState?.periods[stationDraft.period]?.stations?.[stationDraft.stationId]?.balancingEnergy?.result;
      return saved
        && saved.purchase.volumeMwh === stationDraft.purchase.volumeMwh
        && saved.sale.volumeMwh === stationDraft.sale.volumeMwh
        && saved.sourceFileName === stationDraft.sourceFileName;
    }),
  );
  const isImbalancesDraftSavedToProject = Boolean(
    imbalancesDraft && stationMemoryOrder
      .filter((stationId) => Boolean(imbalancesDraft.stations?.[stationId]))
      .every((stationId) => {
        const saved = projectMemoryState?.periods[imbalancesDraft.period]?.stations?.[stationId]?.imbalances?.result;
        const stationDraft = imbalancesDraft.stations?.[stationId];
        return saved
          && stationDraft
          && saved.negativeImbalanceVolumeMwh === stationDraft.negativeImbalanceVolumeMwh
          && saved.positiveImbalanceVolumeMwh === stationDraft.positiveImbalanceVolumeMwh
          && saved.netImbalanceResultUah === stationDraft.netImbalanceResultUah;
      }),
  );
  const imbalanceRows = imbalancesDraft ? stationMemoryOrder.flatMap((stationId) => imbalancesDraft.stations?.[stationId] ?? []) : [];
  const finalReportData = useMemo(
    () => buildTable2FinalReportData(projectMemoryState, finalReportPeriod),
    [finalReportPeriod, projectMemoryState],
  );
  const activeModuleMeta = modules.find((module) => module.id === activeModule) ?? modules[0];
  const ActiveModuleIcon = activeModuleMeta.icon;

  useEffect(() => {
    const isEmptyDefaultState =
      station === 'Олександрійська БЕСС' && eurRate === '' && fileName === '' && !result && !paymentCalculation && !table2Draft;

    if (isEmptyDefaultState && !dataHubDraft && !marketPricesDraft && !mmsDraft && !balancingEnergyDraft && !imbalancesDraft) {
      clearDraftState();
      return;
    }

    const stateToPersist: PersistedAppState = {
      station,
      eurRate,
      fileName,
      result,
      paymentCalculation,
      table2RdnVdrDraft: table2Draft,
      table2DataHubDraft: dataHubDraft,
      table2MarketPricesDraft: marketPricesDraft,
      table2MmsDraft: mmsDraft,
      table2BalancingEnergyDraft: balancingEnergyDraft,
      table2ImbalancesDraft: imbalancesDraft,
    };

    saveDraftState(stateToPersist);
  }, [balancingEnergyDraft, dataHubDraft, eurRate, fileName, imbalancesDraft, marketPricesDraft, mmsDraft, paymentCalculation, result, station, table2Draft]);

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

  function handleClearDraftData() {
    clearDraftState();
    setProjectMemoryState(loadReportState());
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
    setTable2Files({});
    setTable2Draft(null);
    setTable2Message('');
    setTable2Error('');
    setDataHubFiles({});
    setDataHubDraft(null);
    setDataHubMessage('');
    setDataHubError('');
    setMarketPricesFile(null);
    setMarketPricesDraft(null);
    setMarketPricesMessage('');
    setMarketPricesError('');
    setMmsFiles({});
    setMmsDraft(null);
    setMmsMessage('');
    setMmsError('');
    setBalancingEnergyFiles({});
    setBalancingEnergyDraft(null);
    setBalancingEnergyMessage('');
    setBalancingEnergyError('');
    setImbalancesDraft(null);
    setImbalancesMessage('');
    setImbalancesError('');
    setHourlyImbalanceRows([]);
    setHourlyImbalanceMessage('');
    setHourlyImbalanceError('');
    setShowHourlyImbalanceDetails(false);
    setCleanupMessage('Черновые данные очищены. Сохраненные месячные отчеты не изменены.');
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  }

  function handleResetMonthlyReport() {
    if (!isReportPeriod(cleanupPeriod)) {
      setCleanupMessage('Выберите период в формате YYYY-MM.');
      setShowResetReportConfirm(false);
      return;
    }

    const removedPeriod = cleanupPeriod;
    const nextState = deleteReportPeriod(removedPeriod);
    setProjectMemoryState(nextState);
    const nextActivePeriod = nextState?.activePeriod ?? removedPeriod;
    setCleanupPeriod(nextActivePeriod);
    setTable1Period(nextActivePeriod);
    setPaymentForPeriod(nextActivePeriod);
    setImbalancesPeriod(nextActivePeriod);
    setFinalReportPeriod(nextActivePeriod);
    setShowResetReportConfirm(false);
    setCleanupMessage(`Месячный отчет за период ${removedPeriod} очищен.`);
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

  function handleTable2FileChange(key: Table2FileKey, file: File | undefined) {
    setTable2Files((currentFiles) => ({
      ...currentFiles,
      [key]: file,
    }));
    setTable2Error('');
    setTable2Message('');
  }

  async function handleCalculateTable2() {
    setTable2Error('');
    setTable2Message('');

    if (!canCalculateTable2) {
      setTable2Error('Загрузите все 4 файла: Олександрія РДН/ВДР и Знаменка РДН/ВДР.');
      return;
    }

    setIsTable2Calculating(true);
    try {
      const parsedFiles = await Promise.all(
        table2FileSlots.map((slot) => {
          const file = table2Files[slot.key];
          if (!file) {
            throw new Error(`Не загружен файл "${slot.label}".`);
          }

          return parseRdnVdrFile({
            file,
            stationId: slot.stationId,
            stationName: slot.stationName,
            market: slot.market,
          });
        }),
      );
      const draft = buildRdnVdrDraft(parsedFiles);
      setTable2Draft(draft);
      setTable2Message('Черновик РДН/ВДР, не сохранен.');
    } catch (error) {
      setTable2Draft(null);
      setTable2Error(error instanceof Error ? error.message : 'Не удалось прочитать файлы РДН/ВДР.');
    } finally {
      setIsTable2Calculating(false);
    }
  }

  function handleSaveTable2ToProject() {
    if (!table2Draft) {
      setTable2Error('Сначала рассчитайте РДН/ВДР.');
      return;
    }

    for (const stationId of stationMemoryOrder) {
      const stationDraft = table2Draft.stations[stationId];
      updateStationModule(table2Draft.period, stationId, 'rdnVdr', {
        manualInputs: {
          period: table2Draft.period,
          stationName: stationDraft.stationName,
        },
        uploadedFiles: [
          {
            fileName: stationDraft.rdn.fileName,
            source: 'browser',
            uploadedAt: new Date().toISOString(),
          },
          {
            fileName: stationDraft.vdr.fileName,
            source: 'browser',
            uploadedAt: new Date().toISOString(),
          },
        ],
        parsedData: {
          rdn: stationDraft.rdn,
          vdr: stationDraft.vdr,
          hourlyRows: stationDraft.rdn.hourlyRows.concat(stationDraft.vdr.hourlyRows),
        },
        hourlyRows: stationDraft.rdn.hourlyRows.concat(stationDraft.vdr.hourlyRows),
        result: {
          period: table2Draft.period,
          stationId,
          stationName: stationDraft.stationName,
          markets: {
            rdn: stationDraft.rdn,
            vdr: stationDraft.vdr,
          },
          totalTradingResultUah: stationDraft.totalTradingResultUah,
        },
        validationErrors: stationDraft.rdn.warnings.concat(stationDraft.vdr.warnings).map((warning) => ({
          moduleName: 'rdnVdr',
          message: warning,
          createdAt: new Date().toISOString(),
        })),
      });
    }

    setProjectMemoryState(loadReportState());
    setTable2Error('');
    setTable2Message('РДН/ВДР сохранено в месячный отчет.');
  }

  function handleDataHubFileChange(key: DataHubFileKey, file: File | undefined) {
    setDataHubFiles((currentFiles) => ({
      ...currentFiles,
      [key]: file,
    }));
    setDataHubError('');
    setDataHubMessage('');
  }

  async function handleCalculateDataHub() {
    setDataHubError('');
    setDataHubMessage('');

    if (!canCalculateDataHub) {
      setDataHubError('Загрузите 2 файла DataHub: Олександрія и Знаменка.');
      return;
    }

    setIsDataHubCalculating(true);
    try {
      const parsedFiles = await Promise.all(
        dataHubFileSlots.map((slot) => {
          const file = dataHubFiles[slot.key];
          if (!file) {
            throw new Error(`Не загружен файл "${slot.label}".`);
          }

          return parseDataHubFile({
            file,
            stationId: slot.stationId,
            stationName: slot.stationName,
          });
        }),
      );
      const draft = buildDataHubDraft(parsedFiles);
      setDataHubDraft(draft);
      setDataHubMessage('Черновик DataHub, не сохранен.');
    } catch (error) {
      setDataHubDraft(null);
      setDataHubError(error instanceof Error ? error.message : 'Не удалось прочитать файлы DataHub.');
    } finally {
      setIsDataHubCalculating(false);
    }
  }

  function handleSaveDataHubToProject() {
    if (!dataHubDraft) {
      setDataHubError('Сначала рассчитайте DataHub.');
      return;
    }

    for (const stationId of stationMemoryOrder) {
      const stationDraft = dataHubDraft.stations[stationId];
      updateStationModule(dataHubDraft.period, stationId, 'datahub', {
        manualInputs: {
          period: dataHubDraft.period,
          stationName: stationDraft.stationName,
        },
        uploadedFiles: [
          {
            fileName: stationDraft.fileName,
            source: 'browser',
            uploadedAt: new Date().toISOString(),
          },
        ],
        parsedData: {
          totalInKwh: stationDraft.totalInKwh,
          totalOutKwh: stationDraft.totalOutKwh,
          hourlyRowsRead: stationDraft.hourlyRowsRead,
          hourlyRows: stationDraft.hourlyRows,
        },
        hourlyRows: stationDraft.hourlyRows,
        result: {
          period: dataHubDraft.period,
          stationId,
          stationName: stationDraft.stationName,
          totalInKwh: stationDraft.totalInKwh,
          totalOutKwh: stationDraft.totalOutKwh,
          totalInMwh: stationDraft.totalInMwh,
          totalOutMwh: stationDraft.totalOutMwh,
          saldoMwh: stationDraft.saldoMwh,
          hourlyRowsRead: stationDraft.hourlyRowsRead,
        },
        validationErrors: (stationDraft.warnings ?? []).map((warning) => ({
          moduleName: 'datahub',
          message: warning,
          createdAt: new Date().toISOString(),
        })),
      });
    }

    setProjectMemoryState(loadReportState());
    setDataHubError('');
    setDataHubMessage('DataHub сохранен в месячный отчет.');
  }

  function handleMarketPricesFileChange(file: File | undefined) {
    setMarketPricesFile(file ?? null);
    setMarketPricesError('');
    setMarketPricesMessage('');
  }

  async function handleCalculateMarketPrices() {
    setMarketPricesError('');
    setMarketPricesMessage('');

    if (!marketPricesFile) {
      setMarketPricesError('Загрузите файл цен небалансов Укренерго.');
      return;
    }

    setIsMarketPricesCalculating(true);
    try {
      const draft = await parseMarketPricesFile(marketPricesFile);
      setMarketPricesDraft(draft);
      setMarketPricesMessage('Черновик цен небалансов, не сохранен.');
    } catch (error) {
      setMarketPricesDraft(null);
      setMarketPricesError(error instanceof Error ? error.message : 'Не удалось прочитать файл цен небалансов.');
    } finally {
      setIsMarketPricesCalculating(false);
    }
  }

  function handleSaveMarketPricesToProject() {
    if (!marketPricesDraft) {
      setMarketPricesError('Сначала прочитайте файл цен небалансов.');
      return;
    }

    updatePeriodMarketPrices(marketPricesDraft.period, {
      manualInputs: {
        period: marketPricesDraft.period,
      },
      uploadedFiles: [
        {
          fileName: marketPricesDraft.fileName,
          source: 'browser',
          uploadedAt: new Date().toISOString(),
        },
      ],
      parsedData: {
        rows: marketPricesDraft.rows,
        columns: marketPricesDraft.columns,
      },
      result: {
        period: marketPricesDraft.period,
        rowsCount: marketPricesDraft.rowsCount,
        firstDate: marketPricesDraft.firstDate,
        lastDate: marketPricesDraft.lastDate,
        averageRdnPriceUah: marketPricesDraft.averageRdnPriceUah,
        averagePositiveImbalancePriceUah: marketPricesDraft.averagePositiveImbalancePriceUah,
        averageNegativeImbalancePriceUah: marketPricesDraft.averageNegativeImbalancePriceUah,
        averageActualImbalancePriceUah: marketPricesDraft.averageActualImbalancePriceUah,
        rows: marketPricesDraft.rows,
      },
      validationErrors: (marketPricesDraft.warnings ?? []).map((warning) => ({
        moduleName: 'imbalances',
        message: warning,
        createdAt: new Date().toISOString(),
      })),
    });

    setProjectMemoryState(loadReportState());
    setMarketPricesError('');
    setMarketPricesMessage('Цены небалансов сохранены в месячный отчет.');
  }

  function handleMmsFileChange(key: MmsFileKey, file: File | undefined) {
    setMmsFiles((currentFiles) => ({
      ...currentFiles,
      [key]: file,
    }));
    setMmsError('');
    setMmsMessage('');
  }

  async function handleCalculateMms() {
    setMmsError('');
    setMmsMessage('');

    if (!canCalculateMms) {
      setMmsError('Загрузите 2 CSV-файла MMS: Олександрія и Знаменка.');
      return;
    }

    setIsMmsCalculating(true);
    try {
      const parsedFiles = await Promise.all(
        mmsFileSlots.map((slot) => {
          const file = mmsFiles[slot.key];
          if (!file) {
            throw new Error(`Не загружен файл "${slot.label}".`);
          }

          return parseMmsFile({
            file,
            stationId: slot.stationId,
            stationName: slot.stationName,
          });
        }),
      );
      const draft = buildMmsDraft(parsedFiles);
      setMmsDraft(draft);
      setMmsMessage('Черновик MMS, не сохранен.');
    } catch (error) {
      setMmsDraft(null);
      setMmsError(error instanceof Error ? error.message : 'Не удалось прочитать CSV-файлы MMS.');
    } finally {
      setIsMmsCalculating(false);
    }
  }

  function handleSaveMmsToProject() {
    if (!mmsDraft) {
      setMmsError('Сначала рассчитайте MMS.');
      return;
    }

    for (const stationId of stationMemoryOrder) {
      const stationDraft = mmsDraft.stations[stationId];
      updateStationModule(mmsDraft.period, stationId, 'mms', {
        manualInputs: {
          period: mmsDraft.period,
          stationName: stationDraft.stationName,
        },
        uploadedFiles: [
          {
            fileName: stationDraft.fileName,
            source: 'browser',
            uploadedAt: new Date().toISOString(),
          },
        ],
        parsedData: {
          knessToStationKwh: stationDraft.knessToStationKwh,
          stationToKnessKwh: stationDraft.stationToKnessKwh,
          operatorToStationMwh: stationDraft.operatorToStationMwh,
          stationToOperatorMwh: stationDraft.stationToOperatorMwh,
          naToStationMwh: stationDraft.naToStationMwh,
          stationToNaMwh: stationDraft.stationToNaMwh,
          otherBalancingToStationMwh: stationDraft.otherBalancingToStationMwh,
          stationToOtherBalancingMwh: stationDraft.stationToOtherBalancingMwh,
          directionDiagnostics: stationDraft.directionDiagnostics,
          knessColumnDiagnostics: stationDraft.knessColumnDiagnostics,
          rowsRead: stationDraft.rowsRead,
          firstDate: stationDraft.firstDate,
          lastDate: stationDraft.lastDate,
        },
        result: {
          period: mmsDraft.period,
          stationId,
          stationName: stationDraft.stationName,
          knessToStationMwh: stationDraft.knessToStationMwh,
          stationToKnessMwh: stationDraft.stationToKnessMwh,
          operatorToStationMwh: stationDraft.operatorToStationMwh,
          stationToOperatorMwh: stationDraft.stationToOperatorMwh,
          naToStationMwh: stationDraft.naToStationMwh,
          stationToNaMwh: stationDraft.stationToNaMwh,
          otherBalancingToStationMwh: stationDraft.otherBalancingToStationMwh,
          stationToOtherBalancingMwh: stationDraft.stationToOtherBalancingMwh,
          directionDiagnostics: stationDraft.directionDiagnostics,
          knessColumnDiagnostics: stationDraft.knessColumnDiagnostics,
          saldoMwh: stationDraft.saldoMwh,
          rowsRead: stationDraft.rowsRead,
          firstDate: stationDraft.firstDate,
          lastDate: stationDraft.lastDate,
        },
        validationErrors: (stationDraft.warnings ?? []).map((warning) => ({
          moduleName: 'mms',
          message: warning,
          createdAt: new Date().toISOString(),
        })),
      });
    }

    setProjectMemoryState(loadReportState());
    setMmsError('');
    setMmsMessage('MMS сохранен в месячный отчет.');
  }

  function handleBalancingEnergyFileChange(key: BalancingEnergyFileKey, file: File | undefined) {
    setBalancingEnergyFiles((currentFiles) => ({
      ...currentFiles,
      [key]: file,
    }));
    setBalancingEnergyError('');
    setBalancingEnergyMessage('');
  }

  async function handleCalculateBalancingEnergy() {
    setBalancingEnergyError('');
    setBalancingEnergyMessage('');

    if (!canCalculateBalancingEnergy) {
      setBalancingEnergyError('Загрузите хотя бы один Excel-файл балансирующей энергии.');
      return;
    }

    setIsBalancingEnergyCalculating(true);
    try {
      const filesToParse = balancingEnergyFileSlots.filter((slot) => Boolean(balancingEnergyFiles[slot.key]));
      const parsedFiles = await Promise.all(
        filesToParse.map((slot) => parseBalancingEnergyFile({
          file: balancingEnergyFiles[slot.key] as File,
          stationId: slot.stationId,
          stationName: slot.stationName,
        })),
      );
      const draft = buildBalancingEnergyDraft(parsedFiles);
      setBalancingEnergyDraft(draft);
      setBalancingEnergyMessage('Черновик балансирующей энергии Укренерго, не сохранен.');
    } catch (error) {
      setBalancingEnergyDraft(null);
      setBalancingEnergyError(error instanceof Error ? error.message : 'Не удалось прочитать файл балансирующей энергии.');
    } finally {
      setIsBalancingEnergyCalculating(false);
    }
  }

  function handleSaveBalancingEnergyToProject() {
    if (!balancingEnergyDraft || balancingEnergyRows.length === 0) {
      setBalancingEnergyError('Сначала прочитайте файл балансирующей энергии.');
      return;
    }

    for (const stationDraft of balancingEnergyRows) {
      updateStationModule(stationDraft.period, stationDraft.stationId, 'balancingEnergy', {
        manualInputs: {
          period: stationDraft.period,
          stationName: stationDraft.stationName,
        },
        uploadedFiles: [
          {
            fileName: stationDraft.sourceFileName,
            source: 'browser',
            uploadedAt: new Date().toISOString(),
          },
        ],
        parsedData: {
          sourceFileName: stationDraft.sourceFileName,
          sheetName: stationDraft.sheetName,
          headerRowNumber: stationDraft.headerRowNumber,
          firstDataRowNumber: stationDraft.firstDataRowNumber,
          rowsRead: stationDraft.rowsRead,
          columns: stationDraft.columns,
          diagnostics: stationDraft.diagnostics,
        },
        result: {
          period: stationDraft.period,
          stationId: stationDraft.stationId,
          stationName: stationDraft.stationName,
          sourceFileName: stationDraft.sourceFileName,
          purchase: stationDraft.purchase,
          sale: stationDraft.sale,
          diagnostics: stationDraft.diagnostics,
        },
        validationErrors: (stationDraft.warnings ?? []).map((warning) => ({
          moduleName: 'balancingEnergy',
          message: warning,
          createdAt: new Date().toISOString(),
        })),
      });
    }

    setProjectMemoryState(loadReportState());
    setBalancingEnergyError('');
    setBalancingEnergyMessage('Балансирующая энергия Укренерго сохранена в месячный отчет.');
  }

  function handleCalculateImbalances() {
    setImbalancesError('');
    setImbalancesMessage('');
    setHourlyImbalanceError('');
    setHourlyImbalanceMessage('');
    setHourlyImbalanceRows([]);

    if (!isReportPeriod(imbalancesPeriod)) {
      setImbalancesError('Период небалансов должен быть в формате YYYY-MM.');
      return;
    }

    const latestState = loadReportState();
    if (!latestState) {
      setImbalancesError('В состоянии проекта нет сохраненных данных. Сначала сохраните РДН/ВДР, DataHub и цены небалансов в месячный отчет.');
      return;
    }

    try {
      const draft = calculateImbalancesDraft(latestState, imbalancesPeriod);
      setProjectMemoryState(latestState);
      setImbalancesDraft(draft);
      setImbalancesMessage('Черновик небалансов, не сохранен.');
    } catch (error) {
      setImbalancesDraft(null);
      setImbalancesError(error instanceof Error ? error.message : 'Не удалось рассчитать небалансы.');
    }
  }

  function handleSaveImbalancesToProject() {
    if (!imbalancesDraft) {
      setImbalancesError('Сначала рассчитайте небалансы.');
      return;
    }

    for (const stationId of stationMemoryOrder) {
      const stationDraft = imbalancesDraft.stations[stationId];
      if (!stationDraft) {
        continue;
      }

      updateStationModule(imbalancesDraft.period, stationId, 'imbalances', {
        manualInputs: {
          period: imbalancesDraft.period,
          stationName: stationDraft.stationName,
        },
        parsedData: {
          calculationMode: imbalancesDraft.calculationMode,
          warnings: imbalancesDraft.warnings,
        },
        result: stationDraft,
        validationErrors: (imbalancesDraft.warnings ?? []).map((warning) => ({
          moduleName: 'imbalances',
          message: warning,
          createdAt: new Date().toISOString(),
        })),
      });
    }

    setProjectMemoryState(loadReportState());
    setImbalancesError('');
    setImbalancesMessage('Небалансы сохранены в месячный отчет.');
  }

  function handleToggleHourlyImbalanceDetails() {
    setHourlyImbalanceError('');
    setHourlyImbalanceMessage('');

    const shouldShow = !showHourlyImbalanceDetails;
    setShowHourlyImbalanceDetails(shouldShow);
    if (!shouldShow) {
      return;
    }

    const latestState = loadReportState();
    if (!latestState) {
      setHourlyImbalanceRows([]);
      setHourlyImbalanceError('В ProjectReportState нет сохраненных данных для почасовой детализации.');
      return;
    }

    try {
      const rows = buildHourlyImbalanceDetailRows(latestState, imbalancesPeriod, hourlyImbalanceStationFilter);
      setProjectMemoryState(latestState);
      setHourlyImbalanceRows(rows);
      setHourlyImbalanceMessage(`Почасовая детализация загружена: ${rows.length} строк.`);
    } catch (error) {
      setHourlyImbalanceRows([]);
      setHourlyImbalanceError(error instanceof Error ? error.message : 'Не удалось построить почасовую детализацию.');
    }
  }

  function handleHourlyImbalanceStationFilterChange(nextFilter: HourlyImbalanceStationFilter) {
    setHourlyImbalanceStationFilter(nextFilter);
    setHourlyImbalanceError('');
    setHourlyImbalanceMessage('');

    if (!showHourlyImbalanceDetails) {
      return;
    }

    const latestState = loadReportState();
    if (!latestState) {
      setHourlyImbalanceRows([]);
      setHourlyImbalanceError('В ProjectReportState нет сохраненных данных для почасовой детализации.');
      return;
    }

    try {
      const rows = buildHourlyImbalanceDetailRows(latestState, imbalancesPeriod, nextFilter);
      setProjectMemoryState(latestState);
      setHourlyImbalanceRows(rows);
      setHourlyImbalanceMessage(`Почасовая детализация обновлена: ${rows.length} строк.`);
    } catch (error) {
      setHourlyImbalanceRows([]);
      setHourlyImbalanceError(error instanceof Error ? error.message : 'Не удалось построить почасовую детализацию.');
    }
  }

  async function handleExportHourlyImbalances() {
    setHourlyImbalanceError('');
    setHourlyImbalanceMessage('');

    const latestState = loadReportState();
    if (!latestState) {
      setHourlyImbalanceError('В ProjectReportState нет сохраненных данных для экспорта.');
      return;
    }

    setIsHourlyImbalanceExporting(true);
    try {
      const rows = hourlyImbalanceRows.length > 0
        ? hourlyImbalanceRows
        : buildHourlyImbalanceDetailRows(latestState, imbalancesPeriod, hourlyImbalanceStationFilter);
      const exportResult = await exportHourlyImbalancesToExcel({
        rows,
        period: imbalancesPeriod,
        stationFilter: hourlyImbalanceStationFilter,
      });
      setHourlyImbalanceRows(rows);
      setShowHourlyImbalanceDetails(true);
      setHourlyImbalanceMessage(`Excel-файл создан: ${exportResult.fileName}, строк: ${exportResult.rowsCount}.`);
    } catch (error) {
      setHourlyImbalanceError(error instanceof Error ? error.message : 'Не удалось экспортировать почасовую детализацию.');
    } finally {
      setIsHourlyImbalanceExporting(false);
    }
  }

  async function handleExportTable2FinalReport() {
    setFinalReportExportStatus(null);
    setFinalReportExportError('');

    const latestState = loadReportState();
    const report = buildTable2FinalReportData(latestState, finalReportPeriod);
    setProjectMemoryState(latestState);
    setIsFinalReportExporting(true);
    try {
      const exportResult = await exportTable2FinalReportInBrowser(report);
      setFinalReportExportStatus(exportResult);
    } catch (error) {
      setFinalReportExportError(error instanceof Error ? error.message : 'Не удалось экспортировать итоговую Таблицу_2.');
    } finally {
      setIsFinalReportExporting(false);
    }
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
          const table0Fcr = latestProjectState.periods[period]?.stations?.[latestStationId]?.table0Fcr;
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
          <p className="eyebrow">Состояние проекта</p>
          <h3 id="memory-title">Сохраненные данные</h3>
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
        <p className="memory-empty">Сохраненных периодов пока нет. Рассчитайте данные и сохраните их в месячный отчет.</p>
      ) : (
        <div className="memory-period-list">
          {memoryPeriods.map((period) => {
            const periodState = projectMemoryState?.periods[period];
            const hasMarketPrices = Boolean(periodState?.marketPrices?.result);

            return (
              <div className="memory-period" key={period}>
                <div className="memory-period-title">
                  <span>Период</span>
                  <strong>{period}</strong>
                  <small>{hasMarketPrices ? 'Цены небалансов ✅' : 'Цены небалансов нет данных'}</small>
                </div>
                <div className="memory-station-grid">
                  {stationMemoryOrder.map((stationId) => {
                    const table0Fcr = periodState?.stations?.[stationId]?.table0Fcr;
                    const table1PaymentsModule = periodState?.stations?.[stationId]?.table1Payments;
                    const rdnVdrModule = periodState?.stations?.[stationId]?.rdnVdr;
                    const dataHubModule = periodState?.stations?.[stationId]?.datahub;
                    const mmsModule = periodState?.stations?.[stationId]?.mms;
                    const balancingEnergyModule = periodState?.stations?.[stationId]?.balancingEnergy;
                    const imbalancesModule = periodState?.stations?.[stationId]?.imbalances;
                    const table0Result = table0Fcr?.result;
                    const hasTable0 = Boolean(table0Result);
                    const hasTable1 = Boolean(table1PaymentsModule?.result);
                    const hasTable2 = Boolean(rdnVdrModule?.result);
                    const hasDataHub = Boolean(dataHubModule?.result);
                    const hasMms = Boolean(mmsModule?.result);
                    const hasBalancingEnergy = Boolean(balancingEnergyModule?.result);
                    const hasImbalances = Boolean(imbalancesModule?.result);
                    const rdnVdrHourlyRowsCount = rdnVdrModule?.hourlyRows?.length ?? rdnVdrModule?.parsedData?.hourlyRows?.length ?? 0;
                    const dataHubHourlyRowsCount = dataHubModule?.hourlyRows?.length ?? dataHubModule?.parsedData?.hourlyRows?.length ?? 0;

                    return (
                      <div className={hasTable0 ? 'memory-station memory-station-saved' : 'memory-station'} key={stationId}>
                        <div className="memory-station-topline">
                          <strong>{stationLabels[stationId]}</strong>
                          <div className="memory-status-list">
                            <span>{hasTable0 ? 'РПЧ/FCR ✅' : 'РПЧ/FCR нет данных'}</span>
                            <span>{hasTable1 ? 'Оплата Укренерго ✅' : 'Оплата Укренерго нет данных'}</span>
                            <span>{hasTable2 ? 'РДН/ВДР ✅' : 'РДН/ВДР нет данных'}</span>
                            <span>{hasDataHub ? 'DataHub ✅' : 'DataHub нет данных'}</span>
                            <span>{hasMms ? 'MMS/KNESS ✅' : 'MMS/KNESS нет данных'}</span>
                            <span>{hasBalancingEnergy ? 'Балансирующая энергия ✅' : 'Балансирующая энергия нет данных'}</span>
                            <span>{hasImbalances ? 'Небалансы ✅' : 'Небалансы нет данных'}</span>
                            <span>Итоговый отчет: по сохраненным данным</span>
                            {showProjectMemoryDetails && hasTable2 && <span>РДН/ВДР почасовых строк: {rdnVdrHourlyRowsCount}</span>}
                            {showProjectMemoryDetails && hasDataHub && <span>DataHub почасовых строк: {dataHubHourlyRowsCount}</span>}
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
                className={activeModule === module.id ? 'module-item module-item-active' : 'module-item'}
                disabled={!module.enabled}
                key={module.title}
                onClick={() => setActiveModule(module.id)}
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
            <p className="eyebrow">{activeModule === 'service' || activeModule === 'projectState' ? 'Служебный раздел' : 'Рабочий раздел'}</p>
            <h2>{activeModuleMeta.title}</h2>
          </div>
          <div className="topbar-status">
            <ActiveModuleIcon size={16} />
            <span>{activeModuleMeta.description}</span>
          </div>
        </header>

        {activeModule === 'service' && (
        <section className="cleanup-panel" aria-labelledby="cleanup-title">
          <div>
            <p className="eyebrow">Очистка данных</p>
            <h3 id="cleanup-title">Черновики и месячные отчеты</h3>
            <p>Безопасная очистка убирает только текущие расчеты с экрана. Сброс месячного отчета удаляет сохраненные данные только за выбранный период.</p>
          </div>
          <div className="cleanup-controls">
            <label>
              <span>Период</span>
              <input
                type="month"
                value={cleanupPeriod}
                onChange={(event) => {
                  setCleanupPeriod(event.target.value as ReportPeriod);
                  setCleanupMessage('');
                }}
              />
            </label>
            <button className="clear-button" type="button" onClick={handleClearDraftData}>
              <Trash2 size={17} />
              Очистить черновые данные
            </button>
            <button className="danger-button" type="button" onClick={() => setShowResetReportConfirm(true)}>
              Сбросить месячный отчет
            </button>
          </div>
          {cleanupMessage && <p className="cleanup-message">{cleanupMessage}</p>}
        </section>
        )}

        {activeModule === 'fcr' && (
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
        )}

        {activeModule === 'fcr' && activeFcrSubTab === 'table0' && (
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
            {activeModule === 'fcr' && activeFcrSubTab === 'table0' && stationFileMismatchWarning && (
              <section className="message-panel message-panel-warning" aria-live="polite">
                <AlertTriangle size={20} />
                <div>
                  <h3>Проверьте станцию</h3>
                  <p>{stationFileMismatchWarning}</p>
                </div>
              </section>
            )}

            {activeModule === 'projectState' && projectMemoryPanel}

            {activeModule === 'fcr' && activeFcrSubTab === 'table1' && (
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

            {activeModule === 'rdnVdr' && (
            <section className="table1-panel table2-panel" aria-labelledby="table2-title">
              <div className="section-heading result-heading">
                <LineChart size={22} />
                <div>
                  <p className="eyebrow">Почасовые рынки</p>
                  <h3 id="table2-title">Таблица_2 — РДН/ВДР</h3>
                  <p>Загрузите результаты РДН и ВДР по двум станциям. Расчет сначала остается черновиком и не попадает в месячный отчет.</p>
                </div>
              </div>

              <div className="table2-upload-grid">
                {table2FileSlots.map((slot) => (
                  <label className="table2-upload-card" key={slot.key}>
                    <span>{slot.label}</span>
                    <strong>{table2Files[slot.key]?.name ?? 'Файл не выбран'}</strong>
                    <input
                      accept=".xlsx"
                      type="file"
                      onChange={(event) => handleTable2FileChange(slot.key, event.target.files?.[0])}
                    />
                  </label>
                ))}
              </div>

              <div className="table2-actions">
                <button className="primary-button" disabled={!canCalculateTable2} type="button" onClick={handleCalculateTable2}>
                  <Calculator size={18} />
                  {isTable2Calculating ? 'Чтение файлов...' : 'Рассчитать РДН/ВДР'}
                </button>
                <button className="clear-button" disabled={!table2Draft} type="button" onClick={handleSaveTable2ToProject}>
                  Сохранить РДН/ВДР в месячный отчет
                </button>
              </div>

              {table2Error && <p className="table1-message table-message-error">{table2Error}</p>}
              {table2Message && <p className="table1-message">{table2Message}</p>}

              {table2Draft && (
                <>
                  <div className={isTable2DraftSavedToProject ? 'draft-status draft-status-saved' : 'draft-status draft-status-unsaved'}>
                    <div>
                      <span>{isTable2DraftSavedToProject ? 'Финальная память' : 'Черновик'}</span>
                      <strong>{isTable2DraftSavedToProject ? 'РДН/ВДР сохранено в месячный отчет' : 'Черновик РДН/ВДР, не сохранен'}</strong>
                      <p>Период: {table2Draft.period}</p>
                    </div>
                  </div>

                  {(table2Draft.warnings ?? []).length > 0 && (
                    <div className="table2-warning-list">
                      {(table2Draft.warnings ?? []).map((warning) => (
                        <p key={warning}>{warning}</p>
                      ))}
                    </div>
                  )}

                  <div className="payment-table-wrap table2-results-wrap">
                    <table className="payment-table table2-results-table">
                      <thead>
                        <tr>
                          <th>Станция</th>
                          <th>Рынок</th>
                          <th>Покупка МВт⋅ч</th>
                          <th>Покупка грн</th>
                          <th>Продажа МВт⋅ч</th>
                          <th>Продажа грн</th>
                          <th>Средняя цена покупки</th>
                          <th>Средняя цена продажи</th>
                          <th>Результат</th>
                        </tr>
                      </thead>
                      <tbody>
                        {table2Rows.map((row) => (
                          <tr key={`${row.stationName}-${row.market}`}>
                            <td>{row.stationName}</td>
                            <td>{row.market}</td>
                            <td>{numberFormatter.format(row.result.purchaseVolumeMwh)}</td>
                            <td>{moneyFormatter.format(row.result.purchaseAmountUah)}</td>
                            <td>{numberFormatter.format(row.result.saleVolumeMwh)}</td>
                            <td>{moneyFormatter.format(row.result.saleAmountUah)}</td>
                            <td>{moneyFormatter.format(row.result.averagePurchasePriceUah)}</td>
                            <td>{moneyFormatter.format(row.result.averageSalePriceUah)}</td>
                            <td>{moneyFormatter.format(row.result.tradingResultUah)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  <div className="table2-summary-grid">
                    <div>
                      <span>Всего покупка РДН</span>
                      <strong>{numberFormatter.format(table2Draft.summary?.rdnPurchaseVolumeMwh ?? 0)} МВт⋅ч / {moneyFormatter.format(table2Draft.summary?.rdnPurchaseAmountUah ?? 0)} грн</strong>
                    </div>
                    <div>
                      <span>Всего продажа РДН</span>
                      <strong>{numberFormatter.format(table2Draft.summary?.rdnSaleVolumeMwh ?? 0)} МВт⋅ч / {moneyFormatter.format(table2Draft.summary?.rdnSaleAmountUah ?? 0)} грн</strong>
                    </div>
                    <div>
                      <span>Всего покупка ВДР</span>
                      <strong>{numberFormatter.format(table2Draft.summary?.vdrPurchaseVolumeMwh ?? 0)} МВт⋅ч / {moneyFormatter.format(table2Draft.summary?.vdrPurchaseAmountUah ?? 0)} грн</strong>
                    </div>
                    <div>
                      <span>Всего продажа ВДР</span>
                      <strong>{numberFormatter.format(table2Draft.summary?.vdrSaleVolumeMwh ?? 0)} МВт⋅ч / {moneyFormatter.format(table2Draft.summary?.vdrSaleAmountUah ?? 0)} грн</strong>
                    </div>
                    <div>
                      <span>Общий результат РДН/ВДР</span>
                      <strong>{moneyFormatter.format(table2Draft.summary?.totalTradingResultUah ?? 0)} грн</strong>
                    </div>
                  </div>
                </>
              )}

            </section>
            )}

            {activeModule === 'datahub' && (
            <section className="table1-panel table2-panel" aria-labelledby="datahub-title">
              <div className="table2-subsection table2-subsection-first" aria-labelledby="datahub-title">
                <div className="section-heading result-heading">
                  <Database size={22} />
                  <div>
                    <p className="eyebrow">DataHub / фактическая энергия</p>
                    <h3 id="datahub-title">Физика станции — лист “Група А”</h3>
                    <p>Загрузите DataHub по двум станциям. IN считается как відпуск в мережу, OUT — как відбір з мережі.</p>
                  </div>
                </div>

                <div className="table2-upload-grid table2-upload-grid-two">
                  {dataHubFileSlots.map((slot) => (
                    <label className="table2-upload-card" key={slot.key}>
                      <span>{slot.label}</span>
                      <strong>{dataHubFiles[slot.key]?.name ?? 'Файл не выбран'}</strong>
                      <input
                        accept=".xlsx"
                        type="file"
                        onChange={(event) => handleDataHubFileChange(slot.key, event.target.files?.[0])}
                      />
                    </label>
                  ))}
                </div>

                <div className="table2-actions">
                  <button className="primary-button" disabled={!canCalculateDataHub} type="button" onClick={handleCalculateDataHub}>
                    <Calculator size={18} />
                    {isDataHubCalculating ? 'Чтение DataHub...' : 'Рассчитать DataHub'}
                  </button>
                  <button className="clear-button" disabled={!dataHubDraft} type="button" onClick={handleSaveDataHubToProject}>
                    Сохранить DataHub в месячный отчет
                  </button>
                </div>

                {dataHubError && <p className="table1-message table-message-error">{dataHubError}</p>}
                {dataHubMessage && <p className="table1-message">{dataHubMessage}</p>}

                {dataHubDraft && (
                  <>
                    <div className={isDataHubDraftSavedToProject ? 'draft-status draft-status-saved' : 'draft-status draft-status-unsaved'}>
                      <div>
                        <span>{isDataHubDraftSavedToProject ? 'Финальная память' : 'Черновик'}</span>
                        <strong>{isDataHubDraftSavedToProject ? 'DataHub сохранен в месячный отчет' : 'Черновик DataHub, не сохранен'}</strong>
                        <p>Период: {dataHubDraft.period}</p>
                      </div>
                    </div>

                    {(dataHubDraft.warnings ?? []).length > 0 && (
                      <div className="table2-warning-list">
                        {(dataHubDraft.warnings ?? []).map((warning) => (
                          <p key={warning}>{warning}</p>
                        ))}
                      </div>
                    )}

                    <div className="payment-table-wrap table2-results-wrap">
                      <table className="payment-table table2-results-table table2-datahub-table">
                        <thead>
                          <tr>
                            <th>Станция</th>
                            <th>Відпуск в мережу / IN, МВт⋅ч</th>
                            <th>Відбір з мережі / OUT, МВт⋅ч</th>
                            <th>Сальдо, МВт⋅ч</th>
                            <th>Период</th>
                          </tr>
                        </thead>
                        <tbody>
                          {dataHubRows.map((row) => (
                            <tr key={row.stationId}>
                              <td>{row.stationName}</td>
                              <td>{numberFormatter.format(row.totalInMwh)}</td>
                              <td>{numberFormatter.format(row.totalOutMwh)}</td>
                              <td>{numberFormatter.format(row.saldoMwh)}</td>
                              <td>{row.period}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    <div className="table2-summary-grid table2-summary-grid-compact">
                      <div>
                        <span>Всього IN</span>
                        <strong>{numberFormatter.format(dataHubDraft.summary?.totalInMwh ?? 0)} МВт⋅ч</strong>
                      </div>
                      <div>
                        <span>Всього OUT</span>
                        <strong>{numberFormatter.format(dataHubDraft.summary?.totalOutMwh ?? 0)} МВт⋅ч</strong>
                      </div>
                      <div>
                        <span>Сальдо</span>
                        <strong>{numberFormatter.format(dataHubDraft.summary?.saldoMwh ?? 0)} МВт⋅ч</strong>
                      </div>
                    </div>
                  </>
                )}
              </div>

            </section>
            )}

            {activeModule === 'imbalances' && (
            <section className="table1-panel table2-panel" aria-labelledby="market-prices-title">
              <div className="subtab-bar imbalances-tabs" role="tablist" aria-label="Небалансы">
                {[
                  ['prices', 'Цены небалансов'],
                  ['balancingEnergy', 'Балансирующая энергия Укренерго'],
                  ['mms', 'MMS / КНЕСС'],
                  ['calculation', 'Расчет небалансов'],
                  ['hourly', 'Почасовая детализация'],
                ].map(([tabId, label]) => (
                  <button
                    className={activeImbalancesSubTab === tabId ? 'subtab-button subtab-button-active' : 'subtab-button'}
                    key={tabId}
                    type="button"
                    role="tab"
                    aria-selected={activeImbalancesSubTab === tabId}
                    onClick={() => setActiveImbalancesSubTab(tabId as ImbalancesSubTab)}
                  >
                    {label}
                  </button>
                ))}
              </div>

              {activeImbalancesSubTab === 'prices' && (
              <div className="table2-subsection table2-subsection-first" aria-labelledby="market-prices-title">
                <div className="section-heading result-heading">
                  <FileSpreadsheet size={22} />
                  <div>
                    <p className="eyebrow">Цены небалансов / Укренерго</p>
                    <h3 id="market-prices-title">Почасовые цены для будущего расчета небалансов</h3>
                    <p>Загрузите файл фактических цен небалансов. Данные сохраняются на уровне периода, потому что они общие для обеих станций.</p>
                  </div>
                </div>

                <div className="table2-upload-grid table2-upload-grid-one">
                  <label className="table2-upload-card">
                    <span>Файл цен Укренерго</span>
                    <strong>{marketPricesFile?.name ?? marketPricesDraft?.fileName ?? 'Файл не выбран'}</strong>
                    <input
                      accept=".xlsx"
                      type="file"
                      onChange={(event) => handleMarketPricesFileChange(event.target.files?.[0])}
                    />
                  </label>
                </div>

                <div className="table2-actions">
                  <button className="primary-button" disabled={!canCalculateMarketPrices} type="button" onClick={handleCalculateMarketPrices}>
                    <Calculator size={18} />
                    {isMarketPricesCalculating ? 'Чтение цен...' : 'Прочитать цены'}
                  </button>
                  <button className="clear-button" disabled={!marketPricesDraft} type="button" onClick={handleSaveMarketPricesToProject}>
                    Сохранить цены в месячный отчет
                  </button>
                </div>

                {marketPricesError && <p className="table1-message table-message-error">{marketPricesError}</p>}
                {marketPricesMessage && <p className="table1-message">{marketPricesMessage}</p>}

                {marketPricesDraft && (
                  <>
                    <div className={isMarketPricesDraftSavedToProject ? 'draft-status draft-status-saved' : 'draft-status draft-status-unsaved'}>
                      <div>
                        <span>{isMarketPricesDraftSavedToProject ? 'Финальная память' : 'Черновик'}</span>
                        <strong>{isMarketPricesDraftSavedToProject ? 'Цены небалансов сохранены в месячный отчет' : 'Черновик цен небалансов, не сохранен'}</strong>
                        <p>Период: {marketPricesDraft.period}</p>
                      </div>
                    </div>

                    {(marketPricesDraft.warnings ?? []).length > 0 && (
                      <div className="table2-warning-list">
                        {(marketPricesDraft.warnings ?? []).map((warning) => (
                          <p key={warning}>{warning}</p>
                        ))}
                      </div>
                    )}

                    <div className="table2-summary-grid table2-prices-summary-grid">
                      <div>
                        <span>Период</span>
                        <strong>{marketPricesDraft.period}</strong>
                      </div>
                      <div>
                        <span>Строк / часов</span>
                        <strong>{marketPricesDraft.rowsCount}</strong>
                      </div>
                      <div>
                        <span>Средняя цена РДН</span>
                        <strong>{moneyFormatter.format(marketPricesDraft.averageRdnPriceUah)} грн</strong>
                      </div>
                      <div>
                        <span>Средняя цена позитивного небаланса</span>
                        <strong>{moneyFormatter.format(marketPricesDraft.averagePositiveImbalancePriceUah)} грн</strong>
                      </div>
                      <div>
                        <span>Средняя цена негативного небаланса</span>
                        <strong>{moneyFormatter.format(marketPricesDraft.averageNegativeImbalancePriceUah)} грн</strong>
                      </div>
                      <div>
                        <span>Первая дата</span>
                        <strong>{marketPricesDraft.firstDate}</strong>
                      </div>
                      <div>
                        <span>Последняя дата</span>
                        <strong>{marketPricesDraft.lastDate}</strong>
                      </div>
                    </div>
                  </>
                )}
              </div>
              )}

              {activeImbalancesSubTab === 'balancingEnergy' && (
              <div className="table2-subsection" aria-labelledby="balancing-energy-title">
                <div className="section-heading result-heading">
                  <FileSpreadsheet size={22} />
                  <div>
                    <p className="eyebrow">Балансирующая энергия Укренерго</p>
                    <h3 id="balancing-energy-title">Місячний / декадний звіт балансуючої енергії</h3>
                    <p>Этот файл является финальным источником строк купівля/продаж балансуючої електроенергії НЕК “УКРЕНЕРГО” в Таблице_2.</p>
                  </div>
                </div>

                <div className="table2-upload-grid table2-upload-grid-two">
                  {balancingEnergyFileSlots.map((slot) => (
                    <label className="table2-upload-card" key={slot.key}>
                      <span>{slot.label}</span>
                      <strong>{balancingEnergyFiles[slot.key]?.name ?? balancingEnergyDraft?.stations?.[slot.stationId]?.sourceFileName ?? 'Файл не выбран'}</strong>
                      <input
                        accept=".xlsx"
                        type="file"
                        onChange={(event) => handleBalancingEnergyFileChange(slot.key, event.target.files?.[0])}
                      />
                    </label>
                  ))}
                </div>

                <div className="table2-actions">
                  <button className="primary-button" disabled={!canCalculateBalancingEnergy} type="button" onClick={handleCalculateBalancingEnergy}>
                    <Calculator size={18} />
                    {isBalancingEnergyCalculating ? 'Чтение файла...' : 'Прочитать балансирующую энергию'}
                  </button>
                  <button className="clear-button" disabled={!balancingEnergyDraft} type="button" onClick={handleSaveBalancingEnergyToProject}>
                    Сохранить балансирующую энергию в месячный отчет
                  </button>
                </div>

                {balancingEnergyError && <p className="table1-message table-message-error">{balancingEnergyError}</p>}
                {balancingEnergyMessage && <p className="table1-message">{balancingEnergyMessage}</p>}

                {balancingEnergyDraft && (
                  <>
                    <div className={isBalancingEnergyDraftSavedToProject ? 'draft-status draft-status-saved' : 'draft-status draft-status-unsaved'}>
                      <div>
                        <span>{isBalancingEnergyDraftSavedToProject ? 'Финальная память' : 'Черновик'}</span>
                        <strong>{isBalancingEnergyDraftSavedToProject ? 'Балансирующая энергия сохранена в месячный отчет' : 'Черновик балансирующей энергии, не сохранен'}</strong>
                        <p>Период: {balancingEnergyDraft.period}</p>
                      </div>
                    </div>

                    {(balancingEnergyDraft.warnings ?? []).length > 0 && (
                      <div className="table2-warning-list">
                        {(balancingEnergyDraft.warnings ?? []).map((warning) => (
                          <p key={warning}>{warning}</p>
                        ))}
                      </div>
                    )}

                    <div className="payment-table-wrap table2-results-wrap">
                      <table className="payment-table table2-results-table balancing-energy-table">
                        <thead>
                          <tr>
                            <th>Станция</th>
                            <th>Файл</th>
                            <th>Период</th>
                            <th>Покупка Укренерго, МВт⋅ч</th>
                            <th>Средняя цена покупки</th>
                            <th>Покупка без НДС</th>
                            <th>Покупка с НДС</th>
                            <th>Продажа Укренерго, МВт⋅ч</th>
                            <th>Средняя цена продажи</th>
                            <th>Продажа без НДС</th>
                            <th>Продажа с НДС</th>
                          </tr>
                        </thead>
                        <tbody>
                          {balancingEnergyRows.map((row) => (
                            <tr key={row.stationId}>
                              <td>{row.stationName}</td>
                              <td>{row.sourceFileName}</td>
                              <td>{row.period}</td>
                              <td>{numberFormatter.format(row.purchase.volumeMwh)}</td>
                              <td>{moneyFormatter.format(row.purchase.averagePriceUahMwh)} грн</td>
                              <td>{moneyFormatter.format(row.purchase.amountWithoutVatUah)} грн</td>
                              <td>{moneyFormatter.format(row.purchase.amountWithVatUah)} грн</td>
                              <td>{numberFormatter.format(row.sale.volumeMwh)}</td>
                              <td>{moneyFormatter.format(row.sale.averagePriceUahMwh)} грн</td>
                              <td>{moneyFormatter.format(row.sale.amountWithoutVatUah)} грн</td>
                              <td>{moneyFormatter.format(row.sale.amountWithVatUah)} грн</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    {balancingEnergyRows.some((row) => row.diagnostics.unknownDirections.length > 0) && (
                      <div className="table2-warning-list">
                        {balancingEnergyRows.flatMap((row) => row.diagnostics.unknownDirections.map((direction) => (
                          <p key={`${row.stationId}-${direction.direction}`}>
                            {row.stationName}: неизвестное направление “{direction.direction}”, строк: {direction.rows}, объем: {numberFormatter.format(direction.volumeMwh)} МВт⋅ч.
                          </p>
                        )))}
                      </div>
                    )}
                  </>
                )}
              </div>
              )}

              {activeImbalancesSubTab === 'mms' && (
              <div className="table2-subsection" aria-labelledby="mms-title">
                <div className="section-heading result-heading">
                  <Database size={22} />
                  <div>
                    <p className="eyebrow">MMS / КНЕСС</p>
                    <h3 id="mms-title">Обмен энергии между KNESS и станциями</h3>
                    <p>Загрузите CSV-экспорты MMS. Приложение читает KNESS ↔ станция и диагностирует все дополнительные направления станции для объемов Укренерго/Operator/NA.</p>
                  </div>
                </div>

                <div className="table2-upload-grid table2-upload-grid-two">
                  {mmsFileSlots.map((slot) => (
                    <label className="table2-upload-card" key={slot.key}>
                      <span>{slot.label}</span>
                      <strong>{mmsFiles[slot.key]?.name ?? 'Файл не выбран'}</strong>
                      <input
                        accept=".csv"
                        type="file"
                        onChange={(event) => handleMmsFileChange(slot.key, event.target.files?.[0])}
                      />
                    </label>
                  ))}
                </div>

                <div className="table2-actions">
                  <button className="primary-button" disabled={!canCalculateMms} type="button" onClick={handleCalculateMms}>
                    <Calculator size={18} />
                    {isMmsCalculating ? 'Чтение MMS...' : 'Рассчитать MMS'}
                  </button>
                  <button className="clear-button" disabled={!mmsDraft} type="button" onClick={handleSaveMmsToProject}>
                    Сохранить MMS в месячный отчет
                  </button>
                </div>

                {mmsError && <p className="table1-message table-message-error">{mmsError}</p>}
                {mmsMessage && <p className="table1-message">{mmsMessage}</p>}

                {mmsDraft && (
                  <>
                    <div className={isMmsDraftSavedToProject ? 'draft-status draft-status-saved' : 'draft-status draft-status-unsaved'}>
                      <div>
                        <span>{isMmsDraftSavedToProject ? 'Финальная память' : 'Черновик'}</span>
                        <strong>{isMmsDraftSavedToProject ? 'MMS сохранен в месячный отчет' : 'Черновик MMS, не сохранен'}</strong>
                        <p>Период: {mmsDraft.period}</p>
                      </div>
                    </div>

                    {(mmsDraft.warnings ?? []).length > 0 && (
                      <div className="table2-warning-list">
                        {(mmsDraft.warnings ?? []).map((warning) => (
                          <p key={warning}>{warning}</p>
                        ))}
                      </div>
                    )}

                    <div className="payment-table-wrap table2-results-wrap">
                      <table className="payment-table table2-results-table table2-mms-table">
                        <thead>
                          <tr>
                            <th>Станция</th>
                            <th>Файл</th>
                            <th>KNESS → станция, МВт⋅ч</th>
                            <th>Станция → KNESS, МВт⋅ч</th>
                            <th>Сальдо, МВт⋅ч</th>
                            <th>Строк/часов</th>
                            <th>Период</th>
                          </tr>
                        </thead>
                        <tbody>
                          {mmsRows.map((row) => (
                            <tr key={row.stationId}>
                              <td>{row.stationName}</td>
                              <td>{row.fileName}</td>
                              <td>{numberFormatter.format(row.knessToStationMwh)}</td>
                              <td>{numberFormatter.format(row.stationToKnessMwh)}</td>
                              <td>{numberFormatter.format(row.saldoMwh)}</td>
                              <td>{row.rowsRead}</td>
                              <td>{row.period}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    <div className="table2-summary-grid table2-summary-grid-compact">
                      <div>
                        <span>Всього KNESS → станции</span>
                        <strong>{numberFormatter.format(mmsDraft.summary?.knessToStationsMwh ?? 0)} МВт⋅ч</strong>
                      </div>
                      <div>
                        <span>Всього станции → KNESS</span>
                        <strong>{numberFormatter.format(mmsDraft.summary?.stationsToKnessMwh ?? 0)} МВт⋅ч</strong>
                      </div>
                      <div>
                        <span>Сальдо MMS</span>
                        <strong>{numberFormatter.format(mmsDraft.summary?.saldoMwh ?? 0)} МВт⋅ч</strong>
                      </div>
                    </div>

                    <div className="table2-diagnostics-block">
                      <div className="section-heading result-heading compact-heading">
                        <Database size={18} />
                        <div>
                          <p className="eyebrow">Колонки KNESS</p>
                          <h4>Диагностика KNESS-колонок MMS</h4>
                          <p>Здесь видно, какие CSV headers использованы для KNESS → station и station → KNESS, и сумма каждой колонки до округления.</p>
                        </div>
                      </div>
                      <button className="memory-details-button diagnostic-toggle" type="button" onClick={() => setShowMmsKnessDiagnostics((value) => !value)}>
                        {showMmsKnessDiagnostics ? 'Скрыть диагностику' : 'Показать диагностику'}
                      </button>

                      {showMmsKnessDiagnostics && (mmsKnessColumnRows.length > 0 ? (
                        <div className="payment-table-wrap table2-results-wrap">
                          <table className="payment-table table2-results-table mms-kness-columns-table">
                            <thead>
                              <tr>
                                <th>Станция</th>
                                <th>Файл</th>
                                <th>Период</th>
                                <th>Направление</th>
                                <th>CSV header</th>
                                <th>Колонка</th>
                                <th>Использовано часов</th>
                                <th>Сумма, kWh</th>
                                <th>Сумма, MWh</th>
                              </tr>
                            </thead>
                            <tbody>
                              {mmsKnessColumnRows.map((column) => (
                                <tr key={`${column.stationId}-${column.direction}-${column.header}`}>
                                  <td>{column.stationName}</td>
                                  <td>{column.sourceFileName}</td>
                                  <td>{column.period}</td>
                                  <td>{column.direction === 'KNESS_TO_STATION' ? 'KNESS → station' : 'station → KNESS'}</td>
                                  <td>{column.header}</td>
                                  <td>{column.columnIndex}</td>
                                  <td>{column.rowsUsed}</td>
                                  <td>{column.totalKwh.toFixed(6)}</td>
                                  <td>{column.totalMwh.toFixed(6)}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      ) : (
                        <p className="table1-message table-message-warning">
                          KNESS-колонки не найдены. Проверьте, что CSV содержит направления KNESS_ENERGY ↔ station.
                        </p>
                      ))}
                    </div>

                    <div className="table2-diagnostics-block">
                      <div className="section-heading result-heading compact-heading">
                        <Database size={18} />
                        <div>
                          <p className="eyebrow">Направления MMS</p>
                          <h4>Найденные направления MMS</h4>
                          <p>Диагностика показывает все направления из CSV и группу, к которой они отнесены для финальной Таблицы_2.</p>
                        </div>
                      </div>
                      <button className="memory-details-button diagnostic-toggle" type="button" onClick={() => setShowMmsDirectionDiagnostics((value) => !value)}>
                        {showMmsDirectionDiagnostics ? 'Скрыть диагностику' : 'Показать диагностику'}
                      </button>

                      {showMmsDirectionDiagnostics && (mmsDirectionRows.length > 0 ? (
                        <div className="payment-table-wrap table2-results-wrap">
                          <table className="payment-table table2-results-table mms-directions-table">
                            <thead>
                              <tr>
                                <th>Станция</th>
                                <th>Направление из CSV</th>
                                <th>Группа</th>
                                <th>Поток</th>
                                <th>Объем, МВт⋅ч</th>
                              </tr>
                            </thead>
                            <tbody>
                              {mmsDirectionRows.map((direction) => (
                                <tr key={`${direction.stationId}-${direction.directionName}`}>
                                  <td>{direction.stationName}</td>
                                  <td>{direction.directionName}</td>
                                  <td>{direction.group}</td>
                                  <td>{direction.flow === 'toStation' ? 'к станции' : 'от станции'}</td>
                                  <td>{numberFormatter.format(direction.volumeMwh)}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      ) : (
                        <p className="table1-message table-message-warning">
                          Направления MMS не найдены. Проверьте заголовки CSV-файла.
                        </p>
                      ))}
                    </div>
                  </>
                )}
              </div>
              )}

              {activeImbalancesSubTab === 'calculation' && (
              <div className="table2-subsection" aria-labelledby="imbalances-title">
                <div className="section-heading result-heading">
                  <Zap size={22} />
                  <div>
                    <p className="eyebrow">Расчет небалансов</p>
                    <h3 id="imbalances-title">Расчет небалансов по сохраненным данным</h3>
                    <p>Расчет использует сохраненные данные проекта: РДН/ВДР, DataHub и цены Укренерго. MMS/KNESS показывается только как контрольный слой.</p>
                  </div>
                </div>

                <div className="table1-controls table2-imbalance-controls">
                  <label>
                    <span>Период расчета</span>
                    <input
                      type="month"
                      value={imbalancesPeriod}
                      onChange={(event) => setImbalancesPeriod(event.target.value as ReportPeriod)}
                    />
                  </label>
                </div>

                <div className="table2-actions">
                  <button className="primary-button" type="button" onClick={handleCalculateImbalances}>
                    <Calculator size={18} />
                    Рассчитать небалансы
                  </button>
                  <button className="clear-button" disabled={!imbalancesDraft} type="button" onClick={handleSaveImbalancesToProject}>
                    Сохранить небалансы в месячный отчет
                  </button>
                </div>

                {imbalancesError && <p className="table1-message table-message-error">{imbalancesError}</p>}
                {imbalancesMessage && <p className="table1-message">{imbalancesMessage}</p>}
              </div>
              )}

              {activeImbalancesSubTab === 'hourly' && (
              <div className="table2-subsection" aria-labelledby="hourly-imbalances-title">
                <div className="section-heading result-heading">
                  <FileSpreadsheet size={22} />
                  <div>
                    <p className="eyebrow">Почасовая детализация</p>
                    <h3 id="hourly-imbalances-title">Проверка почасового расчета небалансов</h3>
                    <p>Используйте этот экран для проверки IN/OUT, РДН/ВДР, цен Укренерго и итогов по каждому часу.</p>
                  </div>
                </div>

                <div className="hourly-detail-toolbar">
                  <label>
                    <span>Станция для детализации</span>
                    <select
                      value={hourlyImbalanceStationFilter}
                      onChange={(event) => handleHourlyImbalanceStationFilterChange(event.target.value as HourlyImbalanceStationFilter)}
                    >
                      <option value="all">Все станции</option>
                      <option value="oleksandriya">Olexandriya</option>
                      <option value="znamyanka">Znamenka</option>
                    </select>
                  </label>
                  <button className="clear-button" type="button" onClick={handleToggleHourlyImbalanceDetails}>
                    {showHourlyImbalanceDetails ? 'Скрыть почасовую детализацию' : 'Показать почасовую детализацию'}
                  </button>
                  <button className="export-button" disabled={isHourlyImbalanceExporting} type="button" onClick={handleExportHourlyImbalances}>
                    <Download size={18} />
                    {isHourlyImbalanceExporting ? 'Экспорт...' : 'Экспорт почасовых небалансов в Excel'}
                  </button>
                </div>

                {hourlyImbalanceError && <p className="table1-message table-message-error">{hourlyImbalanceError}</p>}
                {hourlyImbalanceMessage && <p className="table1-message">{hourlyImbalanceMessage}</p>}

                {showHourlyImbalanceDetails && hourlyImbalanceRows.length > 0 && (
                  <div className="payment-table-wrap table2-results-wrap">
                    <table className="payment-table table2-results-table hourly-imbalances-table">
                      <thead>
                        <tr>
                          <th>Станция</th>
                          <th>Дата</th>
                          <th>Час</th>
                          <th>DataHub IN, MWh</th>
                          <th>DataHub OUT, MWh</th>
                          <th>Покупка РДН/ВДР, MWh</th>
                          <th>Продажа РДН/ВДР, MWh</th>
                          <th>Негативный небаланс, MWh</th>
                          <th>Позитивный небаланс, MWh</th>
                          <th>Цена РДН</th>
                          <th>Цена негативного небаланса</th>
                          <th>Цена позитивного небаланса</th>
                          <th>Примененная цена негативного</th>
                          <th>Примененная цена позитивного</th>
                          <th>Стоимость негативного</th>
                          <th>Стоимость позитивного</th>
                          <th>Итоговый результат</th>
                        </tr>
                      </thead>
                      <tbody>
                        {hourlyImbalanceRows.map((row) => (
                          <tr key={`${row.stationId}-${row.date}-${row.hour}`}>
                            <td>{row.stationName}</td>
                            <td>{row.date}</td>
                            <td>{row.hour}</td>
                            <td>{numberFormatter.format(row.dataHubInMwh)}</td>
                            <td>{numberFormatter.format(row.dataHubOutMwh)}</td>
                            <td>{numberFormatter.format(row.purchaseVolumeMwh)}</td>
                            <td>{numberFormatter.format(row.saleVolumeMwh)}</td>
                            <td>{numberFormatter.format(row.negativeImbalanceVolumeMwh)}</td>
                            <td>{numberFormatter.format(row.positiveImbalanceVolumeMwh)}</td>
                            <td>{moneyFormatter.format(row.rdnPriceUah)}</td>
                            <td>{moneyFormatter.format(row.negativeImbalancePriceUah)}</td>
                            <td>{moneyFormatter.format(row.positiveImbalancePriceUah)}</td>
                            <td>{moneyFormatter.format(row.negativePriceUsedUah)}</td>
                            <td>{moneyFormatter.format(row.positivePriceUsedUah)}</td>
                            <td>{moneyFormatter.format(row.negativeCostUah)}</td>
                            <td>{moneyFormatter.format(row.positiveCostUah)}</td>
                            <td>{moneyFormatter.format(row.netResultUah)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
              )}

              {activeImbalancesSubTab === 'calculation' && (
              <div className="table2-subsection">
                {imbalancesDraft && (
                  <>
                    <div className={isImbalancesDraftSavedToProject ? 'draft-status draft-status-saved' : 'draft-status draft-status-unsaved'}>
                      <div>
                        <span>{isImbalancesDraftSavedToProject ? 'Финальная память' : 'Черновик'}</span>
                        <strong>{isImbalancesDraftSavedToProject ? 'Небалансы сохранены в месячный отчет' : 'Черновик небалансов, не сохранен'}</strong>
                        <p>
                          Период: {imbalancesDraft.period} · Режим: {imbalancesDraft.calculationMode === 'hourly' ? 'Почасовой расчет' : 'Приблизительный месячный расчет'} ·
                          Использовано часов: {imbalancesDraft.hourlyRowsUsed} · Пропущенные часы: {imbalancesDraft.missingHours}
                        </p>
                      </div>
                    </div>

                    {(imbalancesDraft.warnings ?? []).length > 0 && (
                      <div className="table2-warning-list">
                        {(imbalancesDraft.warnings ?? []).map((warning) => (
                          <p key={warning}>{warning}</p>
                        ))}
                      </div>
                    )}

                    <div className="payment-table-wrap table2-results-wrap">
                      <table className="payment-table table2-results-table table2-imbalances-table">
                        <thead>
                          <tr>
                            <th>Станция</th>
                            <th>Режим</th>
                            <th>Использовано часов</th>
                            <th>Пропущенные часы</th>
                            <th>Негативный небаланс, МВт⋅ч</th>
                            <th>Стоимость негативного, грн</th>
                            <th>Средняя цена негативного, грн</th>
                            <th>Позитивный небаланс, МВт⋅ч</th>
                            <th>Стоимость позитивного, грн</th>
                            <th>Средняя цена позитивного, грн</th>
                            <th>Итоговый результат, грн</th>
                            <th>KNESS → станция, МВт⋅ч</th>
                            <th>Станция → KNESS, МВт⋅ч</th>
                            <th>MMS balance, МВт⋅ч</th>
                          </tr>
                        </thead>
                        <tbody>
                          {imbalanceRows.map((row) => (
                            <tr key={row.stationId}>
                              <td>{row.stationName}</td>
                              <td>{row.calculationMode === 'hourly' ? 'Почасовой' : 'Приблизительный месячный'}</td>
                              <td>{row.hourlyRowsUsed}</td>
                              <td>{row.missingHours}</td>
                              <td>{numberFormatter.format(row.negativeImbalanceVolumeMwh)}</td>
                              <td>{moneyFormatter.format(row.negativeImbalanceCostUah)}</td>
                              <td>{moneyFormatter.format(row.averageNegativeImbalancePriceUsedUah)}</td>
                              <td>{numberFormatter.format(row.positiveImbalanceVolumeMwh)}</td>
                              <td>{moneyFormatter.format(row.positiveImbalanceCostUah)}</td>
                              <td>{moneyFormatter.format(row.averagePositiveImbalancePriceUsedUah)}</td>
                              <td>{moneyFormatter.format(row.netImbalanceResultUah)}</td>
                              <td>{numberFormatter.format(row.mmsKnessToStationMwh)}</td>
                              <td>{numberFormatter.format(row.mmsStationToKnessMwh)}</td>
                              <td>{numberFormatter.format(row.mmsBalanceMwh)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    <div className="table2-summary-grid table2-summary-grid-compact">
                      <div>
                        <span>Режим расчета</span>
                        <strong>{imbalancesDraft.calculationMode === 'hourly' ? 'Почасовой расчет' : 'Приблизительный месячный расчет'}</strong>
                      </div>
                      <div>
                        <span>Использовано часов</span>
                        <strong>{imbalancesDraft.summary.hourlyRowsUsed}</strong>
                      </div>
                      <div>
                        <span>Пропущенные часы</span>
                        <strong>{imbalancesDraft.summary.missingHours}</strong>
                      </div>
                      <div>
                        <span>Негативный объем</span>
                        <strong>{numberFormatter.format(imbalancesDraft.summary.negativeImbalanceVolumeMwh)} МВт⋅ч</strong>
                      </div>
                      <div>
                        <span>Негативная стоимость</span>
                        <strong>{moneyFormatter.format(imbalancesDraft.summary.negativeImbalanceCostUah)} грн</strong>
                      </div>
                      <div>
                        <span>Позитивный объем</span>
                        <strong>{numberFormatter.format(imbalancesDraft.summary.positiveImbalanceVolumeMwh)} МВт⋅ч</strong>
                      </div>
                      <div>
                        <span>Позитивная стоимость</span>
                        <strong>{moneyFormatter.format(imbalancesDraft.summary.positiveImbalanceCostUah)} грн</strong>
                      </div>
                      <div>
                        <span>Общий результат</span>
                        <strong>{moneyFormatter.format(imbalancesDraft.summary.netImbalanceResultUah)} грн</strong>
                      </div>
                    </div>
                  </>
                )}
              </div>
              )}
            </section>
            )}

            {activeModule === 'finalReport' && (
              <section className="table1-panel table2-panel" aria-labelledby="final-report-title">
                <div className="section-heading result-heading">
                  <FileText size={22} />
                  <div>
                    <p className="eyebrow">Итоговый отчет</p>
                    <h3 id="final-report-title">Таблица_2 — купівля/продаж електроенергії</h3>
                    <p>Отчет формируется только из ProjectReportState за выбранный месяц. Тестовые цифры из шаблона не используются как источник данных.</p>
                  </div>
                </div>

                <div className="final-report-toolbar">
                  <label>
                    <span>Период отчета</span>
                    <input
                      type="month"
                      value={finalReportPeriod}
                      onChange={(event) => {
                        setFinalReportPeriod(event.target.value as ReportPeriod);
                        setFinalReportExportStatus(null);
                        setFinalReportExportError('');
                      }}
                    />
                  </label>
                  <div>
                    <span>Лист шаблона</span>
                    <strong>{finalReportData.monthSheetName}</strong>
                  </div>
                  <button className="export-button" disabled={isFinalReportExporting} type="button" onClick={handleExportTable2FinalReport}>
                    <Download size={18} />
                    {isFinalReportExporting ? 'Экспорт...' : 'Экспортировать Таблицу_2 в Excel'}
                  </button>
                </div>

                {finalReportExportError && <p className="table1-message table-message-error">{finalReportExportError}</p>}
                {finalReportExportStatus && (
                  <div className="export-summary" aria-live="polite">
                    <h4>Статус экспорта Таблицы_2</h4>
                    <div className="export-summary-grid">
                      <div>
                        <span>Файл</span>
                        <strong>{finalReportExportStatus.fileName}</strong>
                      </div>
                      <div>
                        <span>Лист</span>
                        <strong>{finalReportExportStatus.sheetName}</strong>
                      </div>
                      <div className="export-summary-wide">
                        <span>Источник шаблона</span>
                        <strong>{finalReportExportStatus.templateSource}</strong>
                      </div>
                      <div className="export-summary-wide">
                        <span>Найденные станции</span>
                        <strong>{finalReportExportStatus.foundStations.join('; ') || '—'}</strong>
                      </div>
                      <div className="export-summary-wide">
                        <span>Заполненные строки</span>
                        <strong>{finalReportExportStatus.filledRows.join('; ') || '—'}</strong>
                      </div>
                      <div className="export-summary-wide">
                        <span>Не заполнено</span>
                        <strong>{finalReportExportStatus.missingRows.join('; ') || '—'}</strong>
                      </div>
                    </div>
                  </div>
                )}

                {(finalReportData.warnings.length > 0 || (finalReportExportStatus?.warnings.length ?? 0) > 0) && (
                  <div className="table2-warning-list">
                    <p>Финальная Таблица_2 пока не полностью готова: проверьте отсутствующие источники и стоимость КНЕСС.</p>
                    <button className="memory-details-button diagnostic-toggle" type="button" onClick={() => setShowFinalReportWarnings((value) => !value)}>
                      {showFinalReportWarnings ? 'Скрыть диагностику' : 'Показать диагностику'}
                    </button>
                    {showFinalReportWarnings && [...finalReportData.warnings, ...(finalReportExportStatus?.warnings ?? [])].map((warning) => (
                      <p key={warning}>{warning}</p>
                    ))}
                  </div>
                )}

                <div className="final-readiness-heading">
                  <p className="eyebrow">Готовность Таблицы_2</p>
                  <h4>Проверка источников по станциям</h4>
                </div>
                <div className="final-readiness-grid">
                  {finalReportData.stations.map((stationReport) => (
                    <div className="final-readiness-card" key={stationReport.stationId}>
                      <strong>{stationReport.stationName}</strong>
                      <div>
                        <span>{stationReport.readiness.rdnVdr ? 'РДН/ВДР загружено' : 'РДН/ВДР отсутствует'}</span>
                        <span>{stationReport.readiness.datahub ? 'DataHub ✅' : 'DataHub отсутствует'}</span>
                        <span>{stationReport.readiness.imbalances ? 'Небалансы рассчитаны' : 'Небалансы отсутствуют'}</span>
                        <span>{stationReport.readiness.mms ? 'MMS/KNESS ✅' : 'MMS/KNESS отсутствует'}</span>
                        <span>{stationReport.readiness.balancingEnergy ? 'Балансирующая энергия Укренерго загружена' : 'Балансирующая энергия Укренерго отсутствует'}</span>
                        <span>Стоимость KNESS: ожидает акт</span>
                      </div>
                    </div>
                  ))}
                </div>

                {finalReportData.stations.map((stationReport) => (
                  <div className="final-report-station" key={stationReport.stationId}>
                    <h4>{stationReport.stationName}</h4>
                    <div className="payment-table-wrap table2-results-wrap">
                      <table className="payment-table final-report-table">
                        <thead>
                          <tr>
                            <th>Операція</th>
                            <th>Обсяг, МВт*год</th>
                            <th>Середньозважена ціна, грн/МВт*год</th>
                            <th>Вартість без ПДВ</th>
                            <th>Вартість з ПДВ</th>
                          </tr>
                        </thead>
                        <tbody>
                          {stationReport.rows.map((row) => (
                            <tr className={row.section === 'saldo' || row.operation.includes('електроенергії') ? 'final-report-total-row' : ''} key={`${stationReport.stationId}-${row.operation}`}>
                              <td>
                                {row.operation}
                                {row.note && <small>{row.note}</small>}
                              </td>
                              <td>{formatOptionalNumber(row.volumeMwh)}</td>
                              <td>{formatOptionalMoney(row.averagePriceUahMwh)}</td>
                              <td>{formatOptionalMoney(row.costWithoutVatUah)}</td>
                              <td>{formatOptionalMoney(row.costWithVatUah)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                ))}
              </section>
            )}

            {activeModule === 'fcr' && activeFcrSubTab === 'table0' && errorMessage && (
              <section className="message-panel message-panel-error" aria-live="polite">
                <AlertCircle size={20} />
                <div>
                  <h3>Ошибка расчета</h3>
                  <p>{errorMessage}</p>
                </div>
              </section>
            )}

            {activeModule === 'fcr' && activeFcrSubTab === 'table0' && monthWarning && (
              <section className="message-panel message-panel-warning" aria-live="polite">
                <AlertTriangle size={20} />
                <div>
                  <h3>Предупреждение</h3>
                  <p>{monthWarning}</p>
                </div>
              </section>
            )}

            {activeModule === 'fcr' && activeFcrSubTab === 'table0' && result && (
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

            {activeModule === 'fcr' && activeFcrSubTab === 'table0' && paymentCalculation && (
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
      {showResetReportConfirm && (
        <div className="modal-backdrop" role="presentation">
          <section className="confirm-modal" role="dialog" aria-modal="true" aria-labelledby="reset-report-title">
            <h3 id="reset-report-title">Сбросить месячный отчет</h3>
            <p>
              Вы уверены, что хотите сбросить месячный отчет за выбранный период?
            </p>
            <p>
              Это действие удалит все сохраненные данные за этот месяц:
            </p>
            <ul>
              <li>FCR/RPC</li>
              <li>оплата Укренерго</li>
              <li>РДН/ВДР</li>
              <li>DataHub</li>
              <li>Imbalances</li>
              <li>MMS/KNESS</li>
              <li>будущие данные итогового отчета</li>
            </ul>
            <p><strong>Это действие нельзя отменить.</strong></p>
            <div className="confirm-actions">
              <button className="clear-button" type="button" onClick={() => setShowResetReportConfirm(false)}>
                Отмена
              </button>
              <button className="danger-button" type="button" onClick={handleResetMonthlyReport}>
                Да, сбросить отчет
              </button>
            </div>
          </section>
        </div>
      )}
    </main>
  );
}
