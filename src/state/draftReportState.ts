export const draftReportStateStorageKey = 'uze-report-automation:draft-state:v1';

export function loadDraftState<T>() {
  if (typeof window === 'undefined') {
    return null;
  }

  try {
    const storedValue = window.localStorage.getItem(draftReportStateStorageKey);
    return storedValue ? (JSON.parse(storedValue) as T) : null;
  } catch {
    return null;
  }
}

export function saveDraftState<T>(state: T) {
  if (typeof window === 'undefined') {
    return;
  }

  window.localStorage.setItem(draftReportStateStorageKey, JSON.stringify(state));
}

export function clearDraftState() {
  if (typeof window === 'undefined') {
    return;
  }

  window.localStorage.removeItem(draftReportStateStorageKey);
}
