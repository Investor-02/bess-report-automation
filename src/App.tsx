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
import { exportTable0RpchInBrowser } from './table0BrowserExport';

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

const modules = [
  { title: 'РПЧ / FCR', description: 'Таблица_0', status: 'Активно', icon: Gauge, enabled: true },
  { title: 'РДН / ВДР', description: 'Почасовые рынки', status: 'Скоро', icon: LineChart, enabled: false },
  { title: 'Небалансы', description: 'Отклонения и сверки', status: 'Скоро', icon: Zap, enabled: false },
  { title: 'DataHub', description: 'Импорт данных', status: 'Скоро', icon: Database, enabled: false },
  { title: 'Итоговый отчет', description: 'Сводный файл', status: 'Скоро', icon: FileText, enabled: false },
];

const storageKey = 'uze-report-automation:fcr-table-0';

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

function readPersistedState(): PersistedAppState | null {
  if (typeof window === 'undefined') {
    return null;
  }

  try {
    const storedValue = window.localStorage.getItem(storageKey);
    if (!storedValue) {
      return null;
    }

    const parsedValue = JSON.parse(storedValue) as Partial<PersistedAppState>;

    return {
      station: isStation(parsedValue.station) ? parsedValue.station : 'Олександрійська БЕСС',
      eurRate: typeof parsedValue.eurRate === 'string' ? parsedValue.eurRate : '',
      fileName: typeof parsedValue.fileName === 'string' ? parsedValue.fileName : '',
      result: parsedValue.result ?? null,
      paymentCalculation: parsedValue.paymentCalculation ?? null,
    };
  } catch {
    return null;
  }
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

  useEffect(() => {
    const isEmptyDefaultState =
      station === 'Олександрійська БЕСС' && eurRate === '' && fileName === '' && !result && !paymentCalculation;

    if (isEmptyDefaultState) {
      window.localStorage.removeItem(storageKey);
      return;
    }

    const stateToPersist: PersistedAppState = {
      station,
      eurRate,
      fileName,
      result,
      paymentCalculation,
    };

    window.localStorage.setItem(storageKey, JSON.stringify(stateToPersist));
  }, [eurRate, fileName, paymentCalculation, result, station]);

  const monthWarning =
    result && /april|апрел/i.test(fileName) && result.totalHours < 700
      ? `Похоже, прочитан не весь месяц. В файле найдены TRUE/FALSE только за ${result.debug.dateColumnsWithValues} дней${
          result.debug.firstDateHeader && result.debug.lastDateHeader
            ? `: ${result.debug.firstDateHeader} — ${result.debug.lastDateHeader}`
            : ''
        }.`
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
    window.localStorage.removeItem(storageKey);
    setStation('Олександрійська БЕСС');
    setEurRate('');
    setFileName('');
    setFilePath('');
    setBrowserFile(null);
    setResult(null);
    setErrorMessage('');
    setExportStatus(null);
    setExportErrorMessage('');
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
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

    if (!paymentCalculation || !result) {
      setExportErrorMessage('Сначала выполните расчет Таблицы_0.');
      return;
    }

    if (!result.debug.firstDateHeader) {
      setExportErrorMessage('Не получилось определить отчетный месяц по данным FCR monitoring.');
      return;
    }

    setIsExporting(true);

    try {
      const exportInput = {
        station: paymentCalculation.station,
        firstDateHeader: result.debug.firstDateHeader,
        certifiedPowerMw: paymentCalculation.certifiedPowerMw,
        trueHours: paymentCalculation.trueHours,
        falseHours: paymentCalculation.falseHours,
        serviceVolume: paymentCalculation.serviceVolume,
        fcrTariffEur: paymentCalculation.fcrTariffEur,
        eurRate: paymentCalculation.eurRate,
        monthlyPriceUah: paymentCalculation.monthlyPriceUah,
        costWithoutVat: paymentCalculation.costWithoutVat,
        vat: paymentCalculation.vat,
        costWithVat: paymentCalculation.costWithVat,
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
        station: paymentCalculation.station,
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
            <h2>РПЧ / FCR — Таблица_0</h2>
          </div>
          <div className="topbar-status">
            <Plug size={16} />
            <span>Чтение Excel и расчет оплаты РПЧ подключены для первого модуля</span>
          </div>
        </header>

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
                <select id="station" value={station} onChange={(event) => setStation(event.target.value as Station)}>
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

        <div className="results-stack">
            {errorMessage && (
              <section className="message-panel message-panel-error" aria-live="polite">
                <AlertCircle size={20} />
                <div>
                  <h3>Ошибка расчета</h3>
                  <p>{errorMessage}</p>
                </div>
              </section>
            )}

            {monthWarning && (
              <section className="message-panel message-panel-warning" aria-live="polite">
                <AlertTriangle size={20} />
                <div>
                  <h3>Предупреждение</h3>
                  <p>{monthWarning}</p>
                </div>
              </section>
            )}

            {result && (
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
              </section>
            )}

            {paymentCalculation && (
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
