/**
 * Server entrypoint and public RPC surface.
 *
 * Keep this file thin:
 * - doGet() handles access checks, action routing, and Index.html rendering.
 * - UI data APIs delegate to Config/Data/Overview/Navigation services.
 * - HtmlService viewport meta is set here because mobile browsers can ignore
 *   the template meta tag inside the Apps Script wrapper.
 */
const ACTION_HANDLERS = {
  normalize_overview_config: {
    admin: true,
    handler: () => normalizeOverviewConfigSetup_(),
  },
  sync_config_table_column_types: {
    admin: true,
    invalidateCache: true,
    wrapOk: true,
    handler: () => syncConfigNativeTableColumnTypes_(),
  },
  sync_config_table_schema: {
    admin: true,
    invalidateCache: true,
    wrapOk: true,
    handler: () => syncConfigNativeTableColumnSchema_(),
  },
  sync_view_config_schema: {
    admin: true,
    invalidateCache: true,
    wrapOk: true,
    handler: () => syncViewConfigSchema_(),
  },
  sync_form_config_schema: {
    admin: true,
    invalidateCache: true,
    wrapOk: true,
    handler: () => syncFormConfigSchema_(),
  },
  seed_config_app_tables: {
    admin: true,
    invalidateCache: true,
    wrapOk: true,
    handler: () => seedConfigAppTables_(),
  },
  ensure_seed_config_app_tables: {
    admin: true,
    invalidateCache: true,
    wrapOk: true,
    handler: () => ensureSeedConfigAppTables_(),
  },
  upsert_user_app_control_defaults: {
    admin: true,
    invalidateCache: true,
    wrapOk: true,
    handler: () => upsertUserAppControlDefaults_(),
  },
  sync_config_app_generated_rows: {
    admin: true,
    invalidateCache: true,
    wrapOk: true,
    handler: () => syncConfigAppGeneratedRows_(),
  },
  write_webapp_link: {
    admin: true,
    handler: () => {
      const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      const navSheet = getSheetByNameOrId_(spreadsheet, '≡ Nav', 1783949092);
      if (navSheet) writeWebAppLink_(navSheet);
      return {
        ok: !!navSheet,
        webAppUrl: ScriptApp.getService().getUrl(),
        updatedAt: new Date().toISOString(),
      };
    },
  },
  debug_overview: {
    admin: true,
    handler: (params) => {
      const overviewKey = params.key || 'process_overview';
      const data = getOverviewData(overviewKey, true);
      return {
        ok: true,
        key: overviewKey,
        title: data.title,
        cards: (data.cards || []).length,
        groups: (data.groups || []).length,
        tableRows: data.table && data.table.rows ? data.table.rows.length : 0,
      };
    },
  },
  debug_table_registry: {
    admin: true,
    handler: () => ({
      ok: true,
      tables: getTableRegistryDebug_(true),
      updatedAt: new Date().toISOString(),
    }),
  },
};

const CONFIG_CLIENT_BUNDLE_FILES = [
  'ClientConfigState',
  'ClientConfigData',
  'ClientConfigRender',
  'ClientConfigForm',
  'ClientConfigRelationships',
  'ClientConfigPreview',
  'ClientConfigActions',
];

const IMPORT_CLIENT_BUNDLE_FILES = [
  'ClientImportActions',
];

function doGet(e) {
  const access = getSpreadsheetAccess_();
  if (!access.ok) return renderAccessDenied_(access);

  const params = e && e.parameter ? e.parameter : {};
  const action = params.action || '';
  if (action) return runAction_(action, params);
  return renderApp_();
}

function runAction_(action, params) {
  const definition = ACTION_HANDLERS[action];
  if (!definition) {
    return jsonResponse_({
      ok: false,
      error: `Unknown action: ${action}`,
      updatedAt: new Date().toISOString(),
    });
  }
  if (definition.admin) assertConfigAppAccess_();
  const result = definition.handler(params || {});
  if (definition.invalidateCache) invalidateDataCache_();
  return jsonResponse_(definition.wrapOk ? {
    ok: true,
    result,
    updatedAt: new Date().toISOString(),
  } : result);
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function renderApp_() {
  const template = HtmlService.createTemplateFromFile('Index');
  template.webAppUrl = ScriptApp.getService().getUrl();
  return template
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setTitle('Process App')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function normalizeOverviewConfigSetup() {
  return normalizeOverviewConfigSetup_();
}

function upsertUserAppControlDefaults() {
  assertConfigAppAccess_();
  const result = upsertUserAppControlDefaults_();
  invalidateDataCache_();
  return sanitizeClientPayload_(result);
}

function syncConfigAppGeneratedRows() {
  assertConfigAppAccess_();
  const result = syncConfigAppGeneratedRows_();
  invalidateDataCache_();
  return sanitizeClientPayload_(result);
}

function hardRefreshAppSettings() {
  assertSpreadsheetAccess_();
  const invalidation = invalidateDataCache_({ strict: true });
  return sanitizeClientPayload_({
    ok: true,
    clearedServerCache: invalidation.ok === true,
    invalidation,
    refreshedAt: new Date().toISOString(),
  });
}

// Backward-compatible API for older clients.
function refreshAppDataCache() {
  return hardRefreshAppSettings();
}

function syncViewConfigSchema_() {
  const registry = getConfigAppRegistry_(true);
  const item = registry.items && registry.items.view;
  const result = ensureConfigViewColumnOrderSchema_(item);
  invalidateConfigTableCache_('view');
  return result;
}

function debugTableRegistry() {
  assertSpreadsheetAccess_();
  return sanitizeClientPayload_(getTableRegistryDebug_(true));
}

function getTableRegistryDebug_(forceRefresh) {
  const tables = getTables_(!!forceRefresh);
  return Object.keys(tables).map((key) => {
    const table = tables[key] || {};
    return {
      key,
      name: table.name || '',
      department: table.department || '',
      group: table.group || '',
      spreadsheetId: table.spreadsheetId || SPREADSHEET_ID,
      gid: table.gid || '',
      sheetName: table.sheetName || '',
      binding: table.binding || '',
      source: table.source || '',
      sourceSpreadsheetUrl: table.sourceSpreadsheetUrl || table.spreadsheetUrl || '',
      accessFallbackReason: table.accessFallbackReason || '',
    };
  });
}

function include(filename) {
  assertSpreadsheetAccess_();
  return HtmlService.createTemplateFromFile(filename).getRawContent();
}

function getAppBootstrap(forceRefresh) {
  const force = !!forceRefresh;
  const timings = [];
  const measure = (label, callback) => {
    const startedAt = Date.now();
    try {
      return callback();
    } finally {
      timings.push({ label, ms: Date.now() - startedAt });
    }
  };

  const access = measure('access', () => assertSpreadsheetAccess_());
  const autoSync = getStartupAutoSyncStatus_();
  const tables = measure('tables', () => getTables_(force));
  const uiConfig = measure('uiConfig', () => getUserAppConfig_(force, { deferDataConfig: true }));
  const projectInfo = measure('projectInfo', () => applyUserProjectConfig_(getProjectInfo_(), uiConfig));
  const navigation = measure('navigation', () => getNavigation(force));
  const initialItem = measure('initialItem', () => getBootstrapInitialItem_(navigation, uiConfig));
  return sanitizeClientPayload_({
    spreadsheetId: access.spreadsheetId,
    spreadsheetUrl: access.spreadsheetUrl,
    tables: compactBootstrapTables_(tables, { includeSchema: false }),
    projectInfo,
    uiConfig,
    navigation,
    initialItem: initialItem ? {
      id: initialItem.id || '',
      pageId: initialItem.pageId || '',
      tableKey: initialItem.tableKey || '',
      overviewKey: initialItem.overviewKey || '',
      linkUrl: initialItem.linkUrl || '',
    } : null,
    initialItemData: null,
    auth: Object.assign({}, access, { configAccess: { deferred: true, canAccessConfig: true } }),
    autoSync,
    performance: { bootstrapTimings: timings },
    generatedAt: new Date().toISOString(),
  });
}

function getAppCacheManifest() {
  return getAppChangeManifest();
}

function getBootstrapInitialItem_(navigation, uiConfig) {
  const items = flattenBootstrapNavItems_(navigation || []);
  const app = uiConfig && uiConfig.app ? uiConfig.app : {};
  const configured = String(app.defaultLandingPage || '').trim();
  const configuredItem = configured
    ? items.find((item) => isBootstrapLandingItemMatch_(item, configured) && isBootstrapLoadableItem_(item))
    : null;
  return configuredItem ||
    items.find((item) => item && item.tableKey && !item.disabled) ||
    items.find(isBootstrapLoadableItem_) ||
    null;
}

function getBootstrapInitialItemData_(item) {
  if (!item || !item.tableKey || item.disabled) return null;
  try {
    return {
      type: 'table',
      itemId: item.id || '',
      pageId: item.pageId || '',
      tableKey: item.tableKey,
      data: getTableDataInitial(item.tableKey, false),
      generatedAt: new Date().toISOString(),
    };
  } catch (error) {
    return {
      type: 'table',
      itemId: item.id || '',
      pageId: item.pageId || '',
      tableKey: item.tableKey,
      error: error && error.message ? error.message : String(error),
    };
  }
}

function flattenBootstrapNavItems_(items) {
  return (items || []).flatMap((item) => [item].concat(item.children ? flattenBootstrapNavItems_(item.children) : []));
}

function isBootstrapLoadableItem_(item) {
  return !!(item && !item.disabled && (item.tableKey || item.overviewKey || item.linkUrl));
}

function isBootstrapLandingItemMatch_(item, key) {
  const normalized = normalizeBootstrapLandingKey_(key);
  if (!normalized || !item) return false;
  return [item.id, item.pageId, item.tableKey, item.overviewKey, item.label]
    .filter(Boolean)
    .some((value) => normalizeBootstrapLandingKey_(value) === normalized);
}

function normalizeBootstrapLandingKey_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function getUserAppRuntimeConfig(forceRefresh) {
  const startedAt = Date.now();
  let stage = 'assert_spreadsheet_access';
  let access = null;
  try {
    access = assertSpreadsheetAccess_();
    stage = 'build_user_app_config';
    logAppDiagnostic_('info', 'config_hydration_started', {
      forceRefresh: !!forceRefresh,
      primarySpreadsheetId: access.spreadsheetId || SPREADSHEET_ID,
    });
    const config = getUserAppConfig_(!!forceRefresh, { deferDataConfig: false });
    logAppDiagnostic_('info', 'config_hydration_completed', {
      forceRefresh: !!forceRefresh,
      primarySpreadsheetId: access.spreadsheetId || SPREADSHEET_ID,
      durationMs: Date.now() - startedAt,
    });
    return sanitizeClientPayload_(config);
  } catch (error) {
    logAppDiagnostic_('error', 'config_hydration_failed', {
      stage,
      forceRefresh: !!forceRefresh,
      primarySpreadsheetId: access && access.spreadsheetId ? access.spreadsheetId : SPREADSHEET_ID,
      durationMs: Date.now() - startedAt,
    }, error);
    throw error;
  }
}

function getBootstrapTableSchemas(forceRefresh) {
  assertSpreadsheetAccess_();
  return sanitizeClientPayload_(compactBootstrapTables_(getTables_(!!forceRefresh), { includeSchema: true }));
}

function compactBootstrapTables_(tables, options) {
  const includeSchema = !!(options && options.includeSchema);
  return Object.keys(tables || {}).reduce((acc, key) => {
    const table = tables[key] || {};
    const compact = {
      key,
      name: table.name || '',
      sheetName: table.sheetName || '',
      gid: table.gid || '',
      type: table.type || '',
      spreadsheetId: table.spreadsheetId || SPREADSHEET_ID,
      spreadsheetUrl: table.spreadsheetUrl || table.sourceSpreadsheetUrl || '',
      sourceSpreadsheetUrl: table.sourceSpreadsheetUrl || table.spreadsheetUrl || '',
      source: table.source || '',
      department: table.department || '',
      group: table.group || '',
      binding: table.binding || '',
      inherited: !!table.inherited,
      readOnly: !!table.readOnly,
      apiTableName: table.apiTableName || '',
      accessFallbackReason: table.accessFallbackReason || '',
    };
    if (includeSchema) {
      compact.columns = table.columns || [];
      compact.columnProperties = compactBootstrapColumnProperties_(table.columnProperties || []);
    }
    acc[key] = compact;
    return acc;
  }, {});
}

function compactBootstrapColumnProperties_(columnProperties) {
  return (columnProperties || []).map((column) => ({
    columnIndex: column.columnIndex,
    columnName: column.columnName || '',
    columnType: column.columnType || '',
  }));
}

function getStartupAutoSyncStatus_() {
  const cached = getCached_(cacheKey_('registry', 'configAutoSync'));
  return cached
    ? Object.assign({}, cached, { cached: true, skippedStartupSync: true })
    : {
      addedCount: 0,
      deletedCount: 0,
      skipped: ['startup_auto_sync_disabled'],
      skippedStartupSync: true,
    };
}

function ensureGeneratedConfigForUserApp_() {
  try {
    return syncConfigAppGeneratedRows_();
  } catch (error) {
    console.warn('Config auto-sync skipped for user app:', error && error.message ? error.message : error);
    return {
      addedCount: 0,
      deletedCount: 0,
      skipped: [error && error.message ? error.message : String(error)],
      failed: true,
    };
  }
}

function getConfigAppData(options) {
  assertConfigAppAccess_();
  return sanitizeClientPayload_(getConfigAppBootstrap(options));
}

function getConfigClientBundle() {
  assertConfigAppAccess_();
  return CONFIG_CLIENT_BUNDLE_FILES
    .map((filename) => HtmlService.createTemplateFromFile(filename).getRawContent())
    .join('\n\n');
}

function getConfigStyleBundle() {
  assertConfigAppAccess_();
  return HtmlService.createTemplateFromFile('StylesConfigApp').getRawContent();
}

function getImportClientBundle() {
  assertSpreadsheetAccess_();
  return IMPORT_CLIENT_BUNDLE_FILES
    .map((filename) => HtmlService.createTemplateFromFile(filename).getRawContent())
    .join('\n\n');
}

function getProjectInfo_() {
  const cacheKey = cacheKey_('registry', 'projectInfo');
  const cached = getCached_(cacheKey);
  if (cached) return cached;

  const info = {
    processName: 'Process App',
  };

  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getProjectInfoSheet_(spreadsheet);
    if (!sheet) return info;

    const values = sheet.getDataRange().getDisplayValues();
    values.some((row) => {
      const itemIndex = row.findIndex((value, index) => {
        const normalized = normalizeConfigHeader_(value);
        return normalized === 'process_name' || (index === 7 && normalized === 'process_name');
      });
      if (itemIndex < 0) return false;

      const processName = row[itemIndex + 1] || row[8] || '';
      if (processName) info.processName = processName;
      return true;
    });
  } catch (error) {
    info.error = error && error.message ? error.message : String(error);
  }

  setCached_(cacheKey, info, PROJECT_INFO_CACHE_TTL_SECONDS);
  return info;
}

function getProjectInfoSheet_(spreadsheet) {
  const exact = spreadsheet.getSheetByName('Project info') || spreadsheet.getSheetByName('project info');
  if (exact) return exact;
  return spreadsheet.getSheets().find((sheet) => normalizeName(sheet.getName()) === 'project info') || null;
}

function isDefaultProcessName_(value) {
  return normalizeName(value) === 'process app';
}

function applyUserProjectConfig_(projectInfo, uiConfig) {
  const info = Object.assign({}, projectInfo || {});
  const app = uiConfig && uiConfig.app ? uiConfig.app : {};
  if (app.appName && (!info.processName || isDefaultProcessName_(info.processName))) info.processName = app.appName;
  if (app.logoUrl) info.logoUrl = app.logoUrl;
  if (app.accentColor) info.accentColor = app.accentColor;
  return info;
}
