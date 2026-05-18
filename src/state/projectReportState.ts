export type ReportPeriod = `${number}-${string}`;

export type StationId = 'oleksandriya' | 'znamyanka';

export type ReportModuleName =
  | 'table0Fcr'
  | 'table1Payments'
  | 'rdnVdr'
  | 'mms'
  | 'imbalances'
  | 'datahub'
  | 'acts'
  | 'summary';

export type UploadedFileRecord = {
  fileName: string;
  source?: 'browser' | 'electron' | 'manual';
  path?: string;
  uploadedAt: string;
};

export type ValidationErrorRecord = {
  moduleName: ReportModuleName;
  message: string;
  createdAt: string;
};

export type ReportModuleState<TManualInputs = Record<string, unknown>, TParsedData = unknown, TResult = unknown> = {
  manualInputs: TManualInputs;
  uploadedFiles: UploadedFileRecord[];
  parsedData: TParsedData | null;
  result: TResult | null;
  validationErrors: ValidationErrorRecord[];
  lastUpdatedAt: string | null;
};

export type Table0FcrModuleState = ReportModuleState<
  {
    stationName?: string;
    eurRate?: string;
    certifiedPowerMw?: number;
    fcrTariffEur?: number;
  },
  {
    trueHours: number;
    falseHours: number;
    totalHours: number;
    debug: unknown;
  },
  {
    station?: string;
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
  }
>;

export type Table1PaymentRecord = {
  id: string;
  paymentDate: string;
  forPeriod: ReportPeriod;
  amountUah: number;
  createdAt: string;
};

export type Table1PaymentsModuleState = ReportModuleState<
  {
    stationName?: string;
    accrualPeriod?: ReportPeriod;
    payments?: Table1PaymentRecord[];
  },
  null,
  {
    period: ReportPeriod;
    stationId: StationId;
    stationName: string;
    serviceVolume: number;
    averagePriceUah: number;
    accruedWithVat: number;
    paidAmount: number;
    debtAmount: number;
    payoutPercent: number;
    payments: Table1PaymentRecord[];
  }
>;

export type RdnVdrMarketResult = {
  market: 'РДН' | 'ВДР';
  purchaseVolumeMwh: number;
  purchaseAmountUah: number;
  saleVolumeMwh: number;
  saleAmountUah: number;
  averagePurchasePriceUah: number;
  averageSalePriceUah: number;
  tradingResultUah: number;
  rowsRead: number;
};

export type RdnVdrModuleState = ReportModuleState<
  {
    period?: ReportPeriod;
    stationName?: string;
  },
  {
    rdn: RdnVdrMarketResult;
    vdr: RdnVdrMarketResult;
  },
  {
    period: ReportPeriod;
    stationId: StationId;
    stationName: string;
    markets: {
      rdn: RdnVdrMarketResult;
      vdr: RdnVdrMarketResult;
    };
    totalTradingResultUah: number;
  }
>;

export type DataHubModuleState = ReportModuleState<
  {
    period?: ReportPeriod;
    stationName?: string;
  },
  {
    totalInKwh: number;
    totalOutKwh: number;
    hourlyRowsRead: number;
  },
  {
    period: ReportPeriod;
    stationId: StationId;
    stationName: string;
    totalInKwh: number;
    totalOutKwh: number;
    totalInMwh: number;
    totalOutMwh: number;
    saldoMwh: number;
    hourlyRowsRead: number;
  }
>;

export type MarketPriceHourlyRow = {
  date: string;
  hour: string;
  rdnPriceUah: number;
  positiveImbalancePriceUah: number;
  negativeImbalancePriceUah: number;
  actualImbalancePriceUah: number | null;
};

export type MarketPricesModuleState = ReportModuleState<
  {
    period?: ReportPeriod;
  },
  {
    rows: MarketPriceHourlyRow[];
    columns: {
      date: string;
      hour: string;
      rdnPriceUah: string;
      positiveImbalancePriceUah: string;
      negativeImbalancePriceUah: string;
      actualImbalancePriceUah: string | null;
    };
  },
  {
    period: ReportPeriod;
    rowsCount: number;
    firstDate: string;
    lastDate: string;
    averageRdnPriceUah: number;
    averagePositiveImbalancePriceUah: number;
    averageNegativeImbalancePriceUah: number;
    averageActualImbalancePriceUah: number | null;
    rows: MarketPriceHourlyRow[];
  }
>;

export type MmsModuleState = ReportModuleState<
  {
    period?: ReportPeriod;
    stationName?: string;
  },
  {
    knessToStationKwh: number;
    stationToKnessKwh: number;
    rowsRead: number;
    firstDate: string;
    lastDate: string;
  },
  {
    period: ReportPeriod;
    stationId: StationId;
    stationName: string;
    knessToStationMwh: number;
    stationToKnessMwh: number;
    saldoMwh: number;
    rowsRead: number;
    firstDate: string;
    lastDate: string;
  }
>;

export type StationReportState = {
  table0Fcr: Table0FcrModuleState;
  table1Payments: Table1PaymentsModuleState;
  rdnVdr: RdnVdrModuleState;
  mms: MmsModuleState;
  imbalances: ReportModuleState;
  datahub: DataHubModuleState;
  acts: ReportModuleState;
  summary: ReportModuleState;
  lastUpdatedAt: string | null;
};

export type PeriodSummary = {
  table0Fcr?: {
    stationsWithData: StationId[];
    totalTrueHours: number;
    totalFalseHours: number;
    totalServiceVolume: number;
    totalCostWithoutVat: number;
    totalVat: number;
    totalCostWithVat: number;
  };
  lastCalculatedAt: string;
};

export type PeriodReportState = {
  period: ReportPeriod;
  stations: Record<StationId, StationReportState>;
  marketPrices: MarketPricesModuleState;
  summary: PeriodSummary | null;
  lastUpdatedAt: string | null;
};

export type ProjectReportState = {
  version: 1;
  activePeriod: ReportPeriod;
  periods: Record<ReportPeriod, PeriodReportState>;
  lastUpdatedAt: string | null;
};

export const projectReportStateStorageKey = 'uze-report-automation:project-report-state:v1';

const stationIds: StationId[] = ['oleksandriya', 'znamyanka'];

function nowIso() {
  return new Date().toISOString();
}

function createEmptyModuleState(): ReportModuleState {
  return {
    manualInputs: {},
    uploadedFiles: [],
    parsedData: null,
    result: null,
    validationErrors: [],
    lastUpdatedAt: null,
  };
}

function createEmptyStationState(): StationReportState {
  return {
    table0Fcr: createEmptyModuleState() as Table0FcrModuleState,
    table1Payments: createEmptyModuleState() as Table1PaymentsModuleState,
    rdnVdr: createEmptyModuleState() as RdnVdrModuleState,
    mms: createEmptyModuleState() as MmsModuleState,
    imbalances: createEmptyModuleState(),
    datahub: createEmptyModuleState() as DataHubModuleState,
    acts: createEmptyModuleState(),
    summary: createEmptyModuleState(),
    lastUpdatedAt: null,
  };
}

function createEmptyPeriodState(period: ReportPeriod): PeriodReportState {
  return {
    period,
    stations: {
      oleksandriya: createEmptyStationState(),
      znamyanka: createEmptyStationState(),
    },
    marketPrices: createEmptyModuleState() as MarketPricesModuleState,
    summary: null,
    lastUpdatedAt: null,
  };
}

function normalizeModuleState<T extends ReportModuleState>(moduleState: T | undefined) {
  return {
    ...createEmptyModuleState(),
    ...(moduleState ?? {}),
  } as T;
}

function normalizeStationState(stationState: Partial<StationReportState> | undefined): StationReportState {
  const emptyStationState = createEmptyStationState();

  return {
    table0Fcr: normalizeModuleState(stationState?.table0Fcr ?? emptyStationState.table0Fcr),
    table1Payments: normalizeModuleState(stationState?.table1Payments ?? emptyStationState.table1Payments),
    rdnVdr: normalizeModuleState(stationState?.rdnVdr ?? emptyStationState.rdnVdr),
    mms: normalizeModuleState(stationState?.mms ?? emptyStationState.mms),
    imbalances: normalizeModuleState(stationState?.imbalances ?? emptyStationState.imbalances),
    datahub: normalizeModuleState(stationState?.datahub ?? emptyStationState.datahub),
    acts: normalizeModuleState(stationState?.acts ?? emptyStationState.acts),
    summary: normalizeModuleState(stationState?.summary ?? emptyStationState.summary),
    lastUpdatedAt: stationState?.lastUpdatedAt ?? null,
  };
}

function normalizePeriod(period: string): ReportPeriod {
  const normalizedPeriod = period.slice(0, 7);

  if (!/^\d{4}-\d{2}$/.test(normalizedPeriod)) {
    throw new Error('Период отчета должен быть в формате YYYY-MM.');
  }

  return normalizedPeriod as ReportPeriod;
}

function ensurePeriod(state: ProjectReportState, period: ReportPeriod) {
  if (!state.periods[period]) {
    state.periods[period] = createEmptyPeriodState(period);
  }

  for (const stationId of stationIds) {
    state.periods[period].stations[stationId] = normalizeStationState(state.periods[period].stations[stationId]);
  }

  state.periods[period].marketPrices ??= createEmptyModuleState() as MarketPricesModuleState;

  return state.periods[period];
}

function getStorage() {
  return typeof window === 'undefined' ? null : window.localStorage;
}

export function createEmptyReportState(period: string): ProjectReportState {
  const normalizedPeriod = normalizePeriod(period);

  return {
    version: 1,
    activePeriod: normalizedPeriod,
    periods: {
      [normalizedPeriod]: createEmptyPeriodState(normalizedPeriod),
    },
    lastUpdatedAt: null,
  };
}

export function loadReportState(): ProjectReportState | null {
  const storage = getStorage();
  if (!storage) {
    return null;
  }

  try {
    const storedValue = storage.getItem(projectReportStateStorageKey);
    if (!storedValue) {
      return null;
    }

    const parsedState = JSON.parse(storedValue) as ProjectReportState;
    if (parsedState.version !== 1 || !parsedState.activePeriod || !parsedState.periods) {
      return null;
    }

    for (const period of Object.keys(parsedState.periods) as ReportPeriod[]) {
      ensurePeriod(parsedState, period);
    }
    ensurePeriod(parsedState, parsedState.activePeriod);

    return parsedState;
  } catch {
    return null;
  }
}

export function saveReportState(state: ProjectReportState) {
  const storage = getStorage();
  if (!storage) {
    return;
  }

  const stateToSave: ProjectReportState = {
    ...state,
    lastUpdatedAt: nowIso(),
  };

  storage.setItem(projectReportStateStorageKey, JSON.stringify(stateToSave));
}

export function updateStationModule<T extends ReportModuleName>(
  period: string,
  stationId: StationId,
  moduleName: T,
  data: Partial<StationReportState[T]>,
) {
  const normalizedPeriod = normalizePeriod(period);
  const state = loadReportState() ?? createEmptyReportState(normalizedPeriod);
  const periodState = ensurePeriod(state, normalizedPeriod);
  const stationState = normalizeStationState(periodState.stations[stationId]);
  periodState.stations[stationId] = stationState;
  const updatedAt = nowIso();

  stationState[moduleName] = {
    ...stationState[moduleName],
    ...data,
    lastUpdatedAt: updatedAt,
  };
  stationState.lastUpdatedAt = updatedAt;
  periodState.lastUpdatedAt = updatedAt;
  state.activePeriod = normalizedPeriod;

  const recalculatedState = recalculateSummary(normalizedPeriod, state);
  saveReportState(recalculatedState);

  return recalculatedState.periods[normalizedPeriod].stations[stationId][moduleName];
}

export function getStationModule<T extends ReportModuleName>(
  period: string,
  stationId: StationId,
  moduleName: T,
) {
  const normalizedPeriod = normalizePeriod(period);
  const state = loadReportState();
  if (!state?.periods[normalizedPeriod]) {
    return null;
  }

  return state.periods[normalizedPeriod].stations[stationId]?.[moduleName] ?? null;
}

export function updatePeriodMarketPrices(period: string, data: Partial<MarketPricesModuleState>) {
  const normalizedPeriod = normalizePeriod(period);
  const state = loadReportState() ?? createEmptyReportState(normalizedPeriod);
  const periodState = ensurePeriod(state, normalizedPeriod);
  const updatedAt = nowIso();

  periodState.marketPrices = {
    ...periodState.marketPrices,
    ...data,
    lastUpdatedAt: updatedAt,
  };
  periodState.lastUpdatedAt = updatedAt;
  state.activePeriod = normalizedPeriod;

  const recalculatedState = recalculateSummary(normalizedPeriod, state);
  saveReportState(recalculatedState);

  return recalculatedState.periods[normalizedPeriod].marketPrices;
}

export function recalculateSummary(period: string, sourceState?: ProjectReportState) {
  const normalizedPeriod = normalizePeriod(period);
  const state = sourceState ?? loadReportState() ?? createEmptyReportState(normalizedPeriod);
  const periodState = ensurePeriod(state, normalizedPeriod);
  const stationsWithData: StationId[] = [];
  const table0Fcr = {
    stationsWithData,
    totalTrueHours: 0,
    totalFalseHours: 0,
    totalServiceVolume: 0,
    totalCostWithoutVat: 0,
    totalVat: 0,
    totalCostWithVat: 0,
  };

  for (const stationId of stationIds) {
    const result = periodState.stations[stationId]?.table0Fcr?.result;
    if (!result) {
      continue;
    }

    stationsWithData.push(stationId);
    table0Fcr.totalTrueHours += result.trueHours;
    table0Fcr.totalFalseHours += result.falseHours;
    table0Fcr.totalServiceVolume += result.serviceVolume;
    table0Fcr.totalCostWithoutVat += result.costWithoutVat;
    table0Fcr.totalVat += result.vat;
    table0Fcr.totalCostWithVat += result.costWithVat;
  }

  const updatedAt = nowIso();
  periodState.summary = {
    table0Fcr,
    lastCalculatedAt: updatedAt,
  };
  periodState.lastUpdatedAt = updatedAt;
  state.activePeriod = normalizedPeriod;
  state.lastUpdatedAt = updatedAt;

  if (!sourceState) {
    saveReportState(state);
  }

  return state;
}
