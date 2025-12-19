/* eslint-disable @typescript-eslint/no-explicit-any */
// Core Google Sheets add-on logic for DataSetIQ.
// Custom functions must remain synchronous in Apps Script, so all fetches are blocking.

const BASE_URL = 'https://datasetiq.com';
const SERIES_PATH = '/api/public/series/';
const SERIES_DATA_PATH = '/data';
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
  seriesId?: string;
  data?: Array<{ date: string; value: number }>;
  dataset?: any;
  scalar?: number;
  error?: { code: string; message: string };
  message?: string;
  status?: string;
}

// Note: User profile endpoint not available in public API

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
  { id: 'BEA', name: 'BEA (Bureau of Economic Analysis)' },
  { id: 'CENSUS', name: 'US Census Bureau' },
  { id: 'EIA', name: 'EIA (Energy Information)' },
  { id: 'IMF', name: 'IMF (International Monetary Fund)' },
  { id: 'OECD', name: 'OECD' },
  { id: 'WORLDBANK', name: 'World Bank' },
  { id: 'ECB', name: 'ECB (European Central Bank)' },
  { id: 'EUROSTAT', name: 'Eurostat' },
  { id: 'BOE', name: 'Bank of England' },
  { id: 'ONS', name: 'ONS (UK Office for National Statistics)' },
  { id: 'STATCAN', name: 'StatCan (Statistics Canada)' },
  { id: 'RBA', name: 'RBA (Reserve Bank of Australia)' },
  { id: 'BOJ', name: 'BOJ (Bank of Japan)' },
];

const PAID_PLANS = ['starter', 'premium', 'pro', 'team', 'enterprise'];

const PREMIUM_FEATURES = {
  FORMULA_BUILDER: 'Formula Builder Wizard',
  RICH_METADATA: 'Full Metadata Panel',
  MULTI_INSERT: 'Multi-Series Insert',
  TEMPLATES: 'Templates Import/Export',
};

const TEMPLATES_PROP = 'DATASETIQ_TEMPLATES';

interface SidebarStatus {
  connected: boolean;
  isPaid: boolean;
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
  const result = buildArrayResult(data);
  
  // Add upgrade message for free users if data is truncated at 100 observations
  const key = getApiKey();
  if (!key && data.length >= 100) {
    result.push(['', '']);
    result.push(['⚠️ Free tier limited to 100 most recent observations', '']);
    result.push(['Upgrade for full access: datasetiq.com/pricing', '']);
  }
  
  return result;
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
  const dataset = response?.dataset;
  if (!dataset || !(normalizedField in dataset)) {
    throw new Error(`Metadata "${normalizedField}" not found.`);
  }
  // @ts-expect-error dynamic access
  return dataset[normalizedField];
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
    return { connected: false, isPaid: false };
  }
  try {
    // Test API key with a minimal search request
    const response = UrlFetchApp.fetch(`${BASE_URL}${SEARCH_PATH}?q=test&limit=1`, {
      method: 'get',
      headers: { Authorization: `Bearer ${key}` },
      muteHttpExceptions: true,
    });
    const status = response.getResponseCode();
    if (status === 401 || status === 403) {
      return { connected: false, isPaid: false, error: 'Invalid API key. Please reconnect.' };
    }
    if (status >= 200 && status < 300) {
      // Valid API key = paid user with premium features
      return {
        connected: true,
        isPaid: true,
        status: '✅ Connected - Premium features unlocked',
      };
    }
    return { connected: false, isPaid: false, error: 'Unable to verify API key' };
  } catch (err) {
    return { connected: false, isPaid: false, error: formatError(err) };
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
  if (!parsed.results || !Array.isArray(parsed.results)) {
    return [];
  }
  return parsed.results.map((item: any) => ({
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
  if (!parsed.results || !Array.isArray(parsed.results)) {
    return [];
  }
  return parsed.results.map((item: any) => ({
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
  
  // Check if this is a metadata-only dataset
  const isMetadataOnly = latestRes?.status === 'metadata_only';
  const isPending = latestRes?.status === 'ingestion_pending';
  
  return {
    latest: latestRes?.scalar,
    meta: metaRes?.dataset,
    isMetadataOnly,
    isPending,
    statusMessage: latestRes?.message
  };
}

/**
 * Request full ingestion for a metadata-only dataset
 */
function requestFullIngestion(seriesId: string) {
  const key = getApiKey();
  const url = `${BASE_URL}/api/datasets/${encodeURIComponent(seriesId)}/fetch`;
  
  const request: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: 'post',
    muteHttpExceptions: true,
    headers: {
      'Content-Type': 'application/json',
      ...(key ? { 'Authorization': `Bearer ${key}` } : {})
    },
    payload: JSON.stringify({})
  };
  
  try {
    const response = UrlFetchApp.fetch(url, request);
    const status = response.getResponseCode();
    const bodyText = response.getContentText();
    const data = parseJson(bodyText);
    
    if (status === 401 || data.requiresAuth) {
      return { requiresAuth: true, error: 'Authentication required' };
    }
    
    if (status === 429 || data.upgradeToPro) {
      return { 
        upgradeToPro: true, 
        limit: data.limit || 100,
        remaining: data.remaining || 0,
        resetAt: data.resetAt
      };
    }
    
    if (status >= 200 && status < 300 && data.success) {
      return { success: true, message: 'Dataset ingestion started' };
    }
    
    return { error: data.error || data.message || 'Failed to request ingestion' };
  } catch (err: any) {
    return { error: err.message || 'Network error' };
  }
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
  const { mode, start, date } = options;
  
  // For metadata mode, use /api/public/series/[id]
  // For data modes, use /api/public/series/[id]/data
  const isMetaMode = mode === 'meta';
  let url = `${BASE_URL}${SERIES_PATH}${encodeURIComponent(seriesId)}`;
  
  if (!isMetaMode) {
    url += SERIES_DATA_PATH;
    const params: string[] = [];
    if (start) params.push(`start=${encodeURIComponent(start)}`);
    if (date) params.push(`end=${encodeURIComponent(date)}`);
    // Set higher limit for authenticated requests
    params.push(`limit=${key ? '1000' : '100'}`);
    if (params.length > 0) {
      url += '?' + params.join('&');
    }
  }

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
      // Transform new API response to expected format
      let transformedResponse: SeriesResponse;
      if (mode === 'meta' && body.dataset) {
        // Metadata response
        transformedResponse = { dataset: body.dataset };
      } else if (body.data) {
        // Data response - transform [{date, value}] to [[date, value]]
        const dataArray = body.data.map((obs: any) => [obs.date, obs.value]);
        transformedResponse = { 
          data: dataArray, 
          seriesId: body.seriesId,
          status: body.status,
          message: body.message
        };
        
        // Handle scalar modes (latest, value, yoy)
        if (mode === 'latest' && dataArray.length > 0) {
          const latest = dataArray[dataArray.length - 1];
          transformedResponse.scalar = latest[1];
        } else if (mode === 'value' && dataArray.length > 0) {
          transformedResponse.scalar = dataArray[0][1];
        }
      } else {
        transformedResponse = body;
      }
      return { response: transformedResponse };
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

/**
 * Check if user has premium access (valid API key)
 */
function checkPremiumAccess(): { allowed: boolean; message?: string } {
  const key = getApiKey();
  if (!key) {
    return {
      allowed: false,
      message: '🔒 Premium features require an API key. Visit datasetiq.com/dashboard/api-keys to get started.',
    };
  }
  // Valid API key = premium access
  return { allowed: true };
}

/**
 * Formula Builder: Generate formula with wizard
 */
function buildFormula(config: {
  functionName: string;
  seriesId: string;
  freq?: string;
  startDate?: string;
}): { formula: string } {
  const { functionName, seriesId, freq, startDate } = config;
  
  let formula = `=${functionName}("${seriesId}"`;
  
  if (functionName === 'DSIQ' || functionName === 'DSIQ_VALUE') {
    if (freq) formula += `, "${freq}"`;
    if (startDate) formula += `, "${startDate}"`;
  }
  
  formula += ')';
  
  return { formula };
}

/**
 * Insert formula into active cell
 */
function insertBuilderFormula(formula: string) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  cell.setFormula(formula);
  return { ok: true };
}

/**
 * Templates: Scan sheet for DSIQ formulas
 */
function scanFormulas(): { formulas: Array<{ cell: string; formula: string }> } {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getDataRange();
  const formulas = range.getFormulas();
  const result: Array<{ cell: string; formula: string }> = [];
  
  for (let row = 0; row < formulas.length; row++) {
    for (let col = 0; col < formulas[row].length; col++) {
      const formula = formulas[row][col];
      if (formula && typeof formula === 'string' && formula.includes('DSIQ')) {
        const cell = sheet.getRange(row + 1, col + 1).getA1Notation();
        result.push({ cell, formula });
      }
    }
  }
  
  return { formulas: result };
}

/**
 * Templates: Save template
 */
function saveTemplate(name: string, formulas: Array<{ cell: string; formula: string }>) {
  const templates = getTemplates();
  const newTemplate = {
    id: Date.now().toString(),
    name,
    formulas,
    createdAt: new Date().toISOString(),
  };
  
  templates.unshift(newTemplate);
  PropertiesService.getUserProperties().setProperty(
    TEMPLATES_PROP,
    JSON.stringify(templates.slice(0, 20))
  );
  
  return { ok: true, template: newTemplate };
}

/**
 * Templates: Get all templates
 */
function getTemplates(): any[] {
  const stored = PropertiesService.getUserProperties().getProperty(TEMPLATES_PROP);
  return stored ? JSON.parse(stored) : [];
}

/**
 * Templates: Load template into sheet
 */
function loadTemplate(templateId: string) {
  const templates = getTemplates();
  const template = templates.find((t) => t.id === templateId);
  
  if (!template) {
    return { ok: false, error: 'Template not found' };
  }
  
  const sheet = SpreadsheetApp.getActiveSheet();
  
  template.formulas.forEach((item: { cell: string; formula: string }) => {
    try {
      const cell = sheet.getRange(item.cell);
      cell.setFormula(item.formula);
    } catch (err) {
      // Continue on error
    }
  });
  
  return { ok: true };
}

/**
 * Templates: Delete template
 */
function deleteTemplate(templateId: string) {
  const templates = getTemplates();
  const filtered = templates.filter((t) => t.id !== templateId);
  PropertiesService.getUserProperties().setProperty(TEMPLATES_PROP, JSON.stringify(filtered));
  return { ok: true };
}

/**
 * Multi-insert: Insert multiple series at once
 */
function insertMultipleSeries(seriesIds: string[], functionName: string) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const startCell = sheet.getActiveCell();
  const startRow = startCell.getRow();
  
  seriesIds.forEach((seriesId, index) => {
    const cell = sheet.getRange(startRow + index, startCell.getColumn());
    const formula = `=${functionName}("${seriesId}")`;
    cell.setFormula(formula);
    addToRecent(seriesId);
  });
  
  return { ok: true, count: seriesIds.length };
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
g.checkPremiumAccess = checkPremiumAccess;
g.buildFormula = buildFormula;
g.insertBuilderFormula = insertBuilderFormula;
g.scanFormulas = scanFormulas;
g.saveTemplate = saveTemplate;
g.getTemplates = getTemplates;
g.loadTemplate = loadTemplate;
g.deleteTemplate = deleteTemplate;
g.insertMultipleSeries = insertMultipleSeries;
g.include = include;

// Export helpers for tests.
export { normalizeDateInput, mapError, buildArrayResult, normalizeOptionalString };
