/* eslint-disable @typescript-eslint/no-explicit-any */
// Core Google Sheets add-on logic for DataSetIQ.
// Custom functions must remain synchronous in Apps Script, so all fetches are blocking.

const BASE_URL = 'https://datasetiq.com';
const SERIES_PATH = '/api/public/sheets/series/';
const ME_PATH = '/api/public/sheets/me';
const SEARCH_PATH = '/api/public/search';
const HEADER_ROW = ['Date', 'Value'];
const API_KEY_PROP = 'DATASETIQ_API_KEY';
const FAVORITES_PROP = 'DATASETIQ_FAVORITES';
const RECENT_PROP = 'DATASETIQ_RECENT';

type ErrorCode =
  | 'NO_KEY'
  | 'INVALID_KEY'
  | 'REVOKED_KEY'
  | 'FREE_LIMIT'
  | 'QUOTA_EXCEEDED'
  | 'PLAN_REQUIRED'
  | 'UNKNOWN';

interface SeriesResponse {
  meta?: { id: string; etag: string };
  data?: Array<[string, number]>;
  scalar?: number;
  error?: { code: string; message: string };
}

interface MeResponse {
  email: string;
  plan: string;
  quota: { used: number; limit: number; reset: string };
  status: string;
}

interface SearchResult {
  id: string;
  title: string;
  frequency?: string;
  units?: string;
  source?: string;
}

const SOURCES = [
  { id: 'FRED', name: 'FRED (Federal Reserve)' },
  { id: 'BLS', name: 'BLS (Bureau of Labor Statistics)' },
  { id: 'OECD', name: 'OECD' },
  { id: 'EUROSTAT', name: 'Eurostat' },
  { id: 'IMF', name: 'IMF' },
  { id: 'WORLDBANK', name: 'World Bank' },
  { id: 'ECB', name: 'ECB (European Central Bank)' },
  { id: 'BOE', name: 'Bank of England' },
  { id: 'CENSUS', name: 'US Census Bureau' },
  { id: 'EIA', name: 'EIA (Energy Information)' },
];

interface SidebarStatus {
  connected: boolean;
  email?: string;
  plan?: string;
  quota?: { used: number; limit: number; reset: string };
  status?: string;
  error?: string;
}

/**
 * Adds menu entry for first-run authorization and sidebar.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('DataSetIQ')
    .addItem('Authorize', 'authorize')
    .addItem('Open Sidebar', 'showSidebar')
    .addToUi();
}

/**
 * One-time call to surface UrlFetchApp permissions.
 */
function authorize() {
  // Lightweight ping to request UrlFetchApp permission.
  UrlFetchApp.fetch(`${BASE_URL}/api/public/ping?ts=${Date.now()}`);
  return 'Authorized';
}

/**
 * Renders the sidebar UI.
 */
function showSidebar() {
  const template = HtmlService.createTemplateFromFile('sidebar');
  const html = template.evaluate().setTitle('DataSetIQ');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Custom function: returns spill array with headers, sorted newest to oldest.
 */
function DSIQ(seriesId: string, frequency?: string | null, startDate?: any) {
  const series = normalizeSeriesId(seriesId);
  const freq = normalizeOptionalString(frequency);
  const start = normalizeDateInput(startDate);
  const { response, errorMessage } = fetchSeries(series, {
    mode: 'table',
    freq,
    start,
  });
  if (errorMessage) {
    throw new Error(errorMessage);
  }
  const data = response?.data ?? [];
  return buildArrayResult(data);
}

/**
 * Custom function: latest value.
 */
function DSIQ_LATEST(seriesId: string) {
  return handleScalar(seriesId, { mode: 'latest' });
}

/**
 * Custom function: value on or before the given date.
 */
function DSIQ_VALUE(seriesId: string, date: any) {
  const normalizedDate = normalizeDateInput(date);
  if (!normalizedDate) {
    throw new Error('Date is required for DSIQ_VALUE.');
  }
  return handleScalar(seriesId, { mode: 'value', date: normalizedDate });
}

/**
 * Custom function: YoY.
 */
function DSIQ_YOY(seriesId: string) {
  return handleScalar(seriesId, { mode: 'yoy' });
}

/**
 * Custom function: metadata lookup.
 */
function DSIQ_META(seriesId: string, field: string) {
  const series = normalizeSeriesId(seriesId);
  const normalizedField = normalizeOptionalString(field);
  if (!normalizedField) {
    throw new Error('Field is required for DSIQ_META.');
  }
  const { response, errorMessage } = fetchSeries(series, { mode: 'meta' });
  if (errorMessage) {
    throw new Error(errorMessage);
  }
  const meta = response?.meta;
  if (!meta || !(normalizedField in meta)) {
    throw new Error(`Metadata "${normalizedField}" not found.`);
  }
  // @ts-expect-error dynamic access
  return meta[normalizedField];
}

/**
 * Sidebar helper: save API key.
 */
function saveApiKey(key: string) {
  const trimmed = normalizeOptionalString(key);
  if (!trimmed) {
    throw new Error('API key is required.');
  }
  PropertiesService.getUserProperties().setProperty(API_KEY_PROP, trimmed);
  return { ok: true };
}

/**
 * Sidebar helper: clear API key.
 */
function clearApiKey() {
  PropertiesService.getUserProperties().deleteProperty(API_KEY_PROP);
  return { ok: true };
}

/**
 * Sidebar helper: fetch profile/entitlements for status panel.
 */
function getSidebarStatus(): SidebarStatus {
  const key = getApiKey();
  if (!key) {
    return { connected: false };
  }
  try {
    const response = UrlFetchApp.fetch(`${BASE_URL}${ME_PATH}`, {
      method: 'get',
      headers: { Authorization: `Bearer ${key}` },
      muteHttpExceptions: true,
    });
    if (response.getResponseCode() === 401) {
      return { connected: false, error: 'Invalid API key. Please reconnect.' };
    }
    const body = parseJson(response.getContentText());
    const me = body as MeResponse;
    return {
      connected: true,
      email: me.email,
      plan: me.plan,
      quota: me.quota,
      status: me.status,
    };
  } catch (err) {
    return { connected: false, error: formatError(err) };
  }
}

/**
 * Sidebar helper: search endpoint.
 */
function searchSeries(query: string): SearchResult[] {
  const q = normalizeOptionalString(query);
  if (!q) {
    return [];
  }
  const key = getApiKey();
  const params: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: 'get',
    muteHttpExceptions: true,
  };
  if (key) {
    params.headers = { Authorization: `Bearer ${key}` };
  }
  const url = `${BASE_URL}${SEARCH_PATH}?q=${encodeURIComponent(q)}`;
  const response = UrlFetchApp.fetch(url, params);
  if (response.getResponseCode() >= 300) {
    return [];
  }
  const parsed = parseJson(response.getContentText());
  if (!Array.isArray(parsed)) {
    return [];
  }
  return parsed.map((item: any) => ({
    id: item.id,
    title: item.title,
    frequency: item.frequency,
    units: item.units,
    source: item.source,
  }));
}

/**
 * Sidebar helper: browse by source
 */
function browseBySource(source: string): SearchResult[] {
  if (!source) return [];
  const key = getApiKey();
  const params: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: 'get',
    muteHttpExceptions: true,
  };
  if (key) {
    params.headers = { Authorization: `Bearer ${key}` };
  }
  const url = `${BASE_URL}${SEARCH_PATH}?source=${encodeURIComponent(source)}&limit=50`;
  const response = UrlFetchApp.fetch(url, params);
  if (response.getResponseCode() >= 300) {
    return [];
  }
  const parsed = parseJson(response.getContentText());
  if (!Array.isArray(parsed)) {
    return [];
  }
  return parsed.map((item: any) => ({
    id: item.id,
    title: item.title,
    frequency: item.frequency,
    units: item.units,
    source: item.source,
  }));
}

/**
 * Sidebar helper: get sources list
 */
function getSources() {
  return SOURCES;
}

/**
 * Sidebar helper: get preview data for a series
 */
function getPreviewData(seriesId: string) {
  const { response: latestRes, errorMessage: latestErr } = fetchSeries(seriesId, { mode: 'latest' });
  const { response: metaRes, errorMessage: metaErr } = fetchSeries(seriesId, { mode: 'meta' });
  
  if (latestErr || metaErr) {
    return { error: latestErr || metaErr };
  }
  
  return {
    latest: latestRes?.scalar,
    meta: metaRes?.meta,
  };
}

/**
 * Sidebar helper: insert formula into active cell.
 */
function insertFormulaIntoActiveCell(seriesId: string, functionName: string) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cell = sheet.getActiveCell();
  let formula = '';
  
  switch (functionName) {
    case 'DSIQ':
      formula = `=DSIQ("${seriesId}")`;
      break;
    case 'DSIQ_LATEST':
      formula = `=DSIQ_LATEST("${seriesId}")`;
      break;
    case 'DSIQ_VALUE':
      formula = `=DSIQ_VALUE("${seriesId}", TODAY())`;
      break;
    case 'DSIQ_YOY':
      formula = `=DSIQ_YOY("${seriesId}")`;
      break;
    case 'DSIQ_META':
      formula = `=DSIQ_META("${seriesId}", "title")`;
      break;
    default:
      formula = `=DSIQ("${seriesId}")`;
  }
  
  cell.setFormula(formula);
  addToRecent(seriesId);
  return { ok: true };
}

/**
 * Favorites management
 */
function getFavorites(): string[] {
  const stored = PropertiesService.getUserProperties().getProperty(FAVORITES_PROP);
  return stored ? JSON.parse(stored) : [];
}

function addFavorite(seriesId: string) {
  const favorites = getFavorites();
  if (!favorites.includes(seriesId)) {
    favorites.unshift(seriesId);
    PropertiesService.getUserProperties().setProperty(
      FAVORITES_PROP,
      JSON.stringify(favorites.slice(0, 50))
    );
  }
  return { ok: true };
}

function removeFavorite(seriesId: string) {
  const favorites = getFavorites();
  const filtered = favorites.filter((id) => id !== seriesId);
  PropertiesService.getUserProperties().setProperty(FAVORITES_PROP, JSON.stringify(filtered));
  return { ok: true };
}

/**
 * Recent series management
 */
function getRecent(): string[] {
  const stored = PropertiesService.getUserProperties().getProperty(RECENT_PROP);
  return stored ? JSON.parse(stored) : [];
}

function addToRecent(seriesId: string) {
  const recent = getRecent();
  const filtered = recent.filter((id) => id !== seriesId);
  filtered.unshift(seriesId);
  PropertiesService.getUserProperties().setProperty(
    RECENT_PROP,
    JSON.stringify(filtered.slice(0, 20))
  );
}

/**
 * Shared scalar handler with retry/backoff.
 */
function handleScalar(seriesId: string, opts: { mode: string; date?: string }) {
  const series = normalizeSeriesId(seriesId);
  const { response, errorMessage } = fetchSeries(series, opts);
  if (errorMessage) {
    throw new Error(errorMessage);
  }
  if (typeof response?.scalar === 'undefined') {
    throw new Error('Value not available.');
  }
  return response.scalar;
}

/**
 * Central fetch with single retry on 429/5xx.
 */
function fetchSeries(
  seriesId: string,
  options: { mode: string; freq?: string; start?: string; date?: string }
): { response?: SeriesResponse; errorMessage?: string } {
  const key = getApiKey();
  const params: Record<string, string> = {};
  if (options.mode) params.mode = options.mode;
  if (options.freq) params.freq = options.freq;
  if (options.start) params.start = options.start;
  if (options.date) params.date = options.date;

  const query = Object.keys(params)
    .map((k) => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`)
    .join('&');
  const url = `${BASE_URL}${SERIES_PATH}${encodeURIComponent(seriesId)}${query ? `?${query}` : ''}`;

  const request: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: 'get',
    muteHttpExceptions: true,
    headers: key ? { Authorization: `Bearer ${key}` } : {},
  };

  let attempt = 0;
  while (attempt < 2) {
    const response = UrlFetchApp.fetch(url, request);
    const status = response.getResponseCode();
    const bodyText = response.getContentText();
    const body = parseJson(bodyText) as SeriesResponse;
    const headers = normalizeHeaders(response.getAllHeaders());

    if (status >= 200 && status < 300 && !body?.error) {
      return { response: body };
    }

    const retryable = status === 429 || status >= 500;
    if (attempt === 0 && retryable) {
      const retryAfterMs = computeRetryAfter(headers, attempt);
      Utilities.sleep(retryAfterMs);
      attempt += 1;
      continue;
    }

    const code = body?.error?.code as ErrorCode | undefined;
    const messageFromBody = body?.error?.message;
    const message = mapError(code, status, messageFromBody);
    return { errorMessage: message };
  }
  return { errorMessage: 'Unable to reach DataSetIQ. Please try again.' };
}

/**
 * Build spill array with header.
 */
function buildArrayResult(data: Array<[string, number]>): any[][] {
  if (!Array.isArray(data)) {
    return [HEADER_ROW];
  }
  const sorted = [...data].sort((a, b) => {
    const aDate = new Date(a[0]).getTime();
    const bDate = new Date(b[0]).getTime();
    return bDate - aDate;
  });
  return [HEADER_ROW, ...sorted];
}

/**
 * Normalize optional values that may come as null/undefined.
 */
function normalizeOptionalString(value: any): string | undefined {
  if (value === null || typeof value === 'undefined') return undefined;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    return trimmed.length ? trimmed : undefined;
  }
  return String(value);
}

function normalizeSeriesId(seriesId: any): string {
  const normalized = normalizeOptionalString(seriesId);
  if (!normalized) {
    throw new Error('series_id is required.');
  }
  return normalized;
}

function normalizeDateInput(value: any): string | undefined {
  if (value === null || typeof value === 'undefined' || value === '') {
    return undefined;
  }
  if (Object.prototype.toString.call(value) === '[object Date]') {
    const date = value as Date;
    if (isNaN(date.getTime())) {
      throw new Error('Invalid date.');
    }
    if (typeof Utilities !== 'undefined') {
      return Utilities.formatDate(date, 'GMT', 'yyyy-MM-dd');
    }
    // Fallback for tests outside Apps Script.
    return date.toISOString().slice(0, 10);
  }
  if (typeof value === 'string') {
    return value;
  }
  throw new Error('Invalid date input.');
}

/**
 * Map standardized error codes to user-friendly messages.
 */
function mapError(code: ErrorCode | undefined, status: number, fallback?: string): string {
  if (code === 'NO_KEY') return 'Please open DataSetIQ sidebar to connect.';
  if (code === 'INVALID_KEY') return 'Invalid API Key. Reconnect at datasetiq.com/dashboard/api-keys';
  if (code === 'REVOKED_KEY') return 'API Key revoked. Get a new key at datasetiq.com/dashboard/api-keys';
  if (code === 'FREE_LIMIT') return 'Free plan limit reached. Upgrade at datasetiq.com/pricing';
  if (code === 'QUOTA_EXCEEDED') return 'Daily Quota Exceeded. Upgrade at datasetiq.com/pricing';
  if (code === 'PLAN_REQUIRED') return 'Upgrade required. Visit datasetiq.com/pricing';
  if (status === 429) return 'Rate limited. Please retry shortly.';
  if (status >= 500) return 'Server unavailable. Please retry.';
  return fallback || 'Unable to fetch data.';
}

function parseJson(body: string): any {
  try {
    return JSON.parse(body);
  } catch (_err) {
    return {};
  }
}

function getApiKey(): string | null {
  return PropertiesService.getUserProperties().getProperty(API_KEY_PROP);
}

function normalizeHeaders(headers: any): Record<string, string> {
  const normalized: Record<string, string> = {};
  Object.keys(headers || {}).forEach((key) => {
    normalized[key.toLowerCase()] = String((headers as any)[key]);
  });
  return normalized;
}

function computeRetryAfter(headers: Record<string, string>, attempt: number): number {
  const retryAfter = headers['retry-after'];
  if (retryAfter) {
    const asNumber = Number(retryAfter);
    if (!isNaN(asNumber)) {
      return asNumber * 1000;
    }
    const parsed = Date.parse(retryAfter);
    if (!isNaN(parsed)) {
      const diff = parsed - Date.now();
      return diff > 0 ? diff : 500 * Math.pow(2, attempt);
    }
  }
  return 500 * Math.pow(2, attempt);
}

function formatError(err: any): string {
  if (err instanceof Error) return err.message;
  return 'Unexpected error';
}

// Expose functions to HTML sandbox.
function include(filename: string) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Ensure custom functions and menu handlers are available globally after bundling.
const g = globalThis as any;
g.onOpen = onOpen;
g.authorize = authorize;
g.showSidebar = showSidebar;
g.DSIQ = DSIQ;
g.DSIQ_LATEST = DSIQ_LATEST;
g.DSIQ_VALUE = DSIQ_VALUE;
g.DSIQ_YOY = DSIQ_YOY;
g.DSIQ_META = DSIQ_META;
g.saveApiKey = saveApiKey;
g.clearApiKey = clearApiKey;
g.getSidebarStatus = getSidebarStatus;
g.searchSeries = searchSeries;
g.include = include;

// Export helpers for tests.
export { normalizeDateInput, mapError, buildArrayResult, normalizeOptionalString };
