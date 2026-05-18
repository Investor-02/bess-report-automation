export type ReportPeriod = `${number}-${string}`;

export type StationId = 'oleksandriya' | 'znamyanka';

export type ReportModuleName =
  | 'table0Fcr'
  | 'table1Payments'
  | 'rdnVdr'
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

export type StationReportState = {
  table0Fcr: Table0FcrModuleState;
  table1Payments: Table1PaymentsModuleState;
  rdnVdr: ReportModuleState;
  imbalances: ReportModuleState;
  datahub: ReportModuleState;
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
    rdnVdr: createEmptyModuleState(),
    imbalances: createEmptyModuleState(),
    datahub: createEmptyModuleState(),
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
    summary: null,
    lastUpdatedAt: null,
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
    state.periods[period].stations[stationId] ??= createEmptyStationState();
  }

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
  const stationState = periodState.stations[stationId];
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
    const result = periodState.stations[stationId].table0Fcr.result;
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
