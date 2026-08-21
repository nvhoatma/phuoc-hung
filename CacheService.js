/**
 * Thin script-cache helpers.
 *
 * Use these for server-side data/validation caching. Invalidate after
 * mutations or config maintenance changes that affect rendered data.
 */
function cacheKey_(type, key) {
  return [CACHE_VERSION, SPREADSHEET_ID, type, key].join(':');
}

function tableScopedCacheKey_(type, tableKey, table) {
  const scopedTable = table || {};
  const sourceId = scopedTable.spreadsheetId || SPREADSHEET_ID;
  const sourceLocation = scopedTable.gid || scopedTable.sheetName || scopedTable.apiTableId || '';
  return cacheKey_(type, [sourceId, sourceLocation, tableKey].join(':'));
}

function tableScopedCacheKeys_(type, tables) {
  return Object.keys(tables || {}).map((key) => tableScopedCacheKey_(type, key, tables[key]));
}

function getCached_(key) {
  try {
    const value = CacheService.getScriptCache().get(key);
    return value ? JSON.parse(value) : null;
  } catch (error) {
    return null;
  }
}

function setCached_(key, value, ttlSeconds) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(value), ttlSeconds || CACHE_TTL_SECONDS);
  } catch (error) {
    // CacheService has a size limit; large tables simply fall back to live reads.
  }
}

function invalidateDataCache_(options) {
  const strict = !!(options && options.strict);
  const changeVersion = typeof markAllAppCachesChanged_ === 'function'
    ? markAllAppCachesChanged_()
    : '';
  try {
    const tables = getCached_(cacheKey_('registry', 'tables')) || getTables_(false);
    const overviewKeys = typeof getOverviewKeys_ === 'function' ? getOverviewKeys_() : [];
    const configKeys = getConfigCacheKeys_();
    const keys = []
      .concat(tableScopedCacheKeys_('table', tables))
      .concat(tableScopedCacheKeys_('tableInitial', tables))
      .concat(tableScopedCacheKeys_('tableMeta', tables))
      .concat(tableScopedCacheKeys_('validation', tables))
      .concat(tableScopedCacheKeys_('formula', tables))
      .concat(configKeys.map((key) => cacheKey_('configTable', key)))
      .concat(overviewKeys.map((key) => cacheKey_('overview', key)))
      .concat([
        cacheKey_('registry', 'tables'),
        cacheKey_('registry', 'nativeTables'),
        cacheKey_('registry', `nativeTables:${SPREADSHEET_ID}`),
        cacheKey_('registry', 'dataConfigSources'),
        cacheKey_('registry', 'userAppConfig'),
        cacheKey_('registry', 'userAppConfig:lite'),
        cacheKey_('registry', 'navigation'),
        cacheKey_('registry', 'projectInfo'),
        cacheKey_('registry', 'configAutoSync'),
      ]);
    CacheService.getScriptCache().removeAll(keys);
    return {
      ok: true,
      removedKeyCount: keys.length,
      changeVersion,
    };
  } catch (error) {
    if (strict) throw error;
    return {
      ok: false,
      removedKeyCount: 0,
      changeVersion,
      message: error && error.message ? error.message : String(error),
    };
  }
}

function invalidateTableDataCache_(tableKey) {
  let tableVersion = '';
  try {
    if (typeof markTableCacheChanged_ === 'function') {
      tableVersion = markTableCacheChanged_(tableKey) || '';
    }
  } catch (error) {
    // Cache removal can still continue when the manifest version cannot be written.
  }
  try {
    const table = (getCached_(cacheKey_('registry', 'tables')) || getTables_(false))[tableKey] || null;
    const keys = [
      cacheKey_('table', tableKey),
      tableScopedCacheKey_('table', tableKey, table),
      tableScopedCacheKey_('tableInitial', tableKey, table),
      tableScopedCacheKey_('tableMeta', tableKey, table),
    ].filter(Boolean);
    CacheService.getScriptCache().removeAll(keys);
    return {
      ok: true,
      tableVersion: String(tableVersion || ''),
      removedKeyCount: keys.length,
    };
  } catch (error) {
    return {
      ok: false,
      tableVersion: String(tableVersion || ''),
      removedKeyCount: 0,
      message: error && error.message ? error.message : String(error),
    };
  }
}

function invalidateNativeTableRegistryCache_(spreadsheetId) {
  const targetSpreadsheetId = spreadsheetId || SPREADSHEET_ID;
  const keys = [
    cacheKey_('registry', 'tables'),
    cacheKey_('registry', 'nativeTables'),
    cacheKey_('registry', `nativeTables:${targetSpreadsheetId}`),
  ];
  CacheService.getScriptCache().removeAll(keys);
  return {
    ok: true,
    removedKeyCount: keys.length,
  };
}

function invalidateConfigTableCache_(configKey) {
  if (typeof markConfigCacheChanged_ === 'function') markConfigCacheChanged_(configKey);
  try {
    const keys = [
      cacheKey_('configTable', configKey),
      cacheKey_('registry', 'userAppConfig'),
      cacheKey_('registry', 'userAppConfig:lite'),
      cacheKey_('registry', 'navigation'),
      cacheKey_('registry', 'configAutoSync'),
    ];
    if (['tableIndex', 'page', 'navigation'].includes(configKey)) {
      keys.push(cacheKey_('registry', 'tables'));
      keys.push(cacheKey_('registry', 'nativeTables'));
      keys.push(cacheKey_('registry', `nativeTables:${SPREADSHEET_ID}`));
      keys.push(cacheKey_('registry', 'dataConfigSources'));
    }
    if (['appConfig', 'projectInfo'].includes(configKey)) {
      keys.push(cacheKey_('registry', 'projectInfo'));
    }
    CacheService.getScriptCache().removeAll(keys);
  } catch (error) {
    // Best-effort targeted invalidation.
  }
}

function getConfigCacheKeys_() {
  try {
    return (typeof CONFIG_APP_TYPES === 'undefined' ? [] : CONFIG_APP_TYPES)
      .flatMap((type) => type.groups || [])
      .flatMap((group) => group.items || [])
      .map((item) => item.key)
      .filter(Boolean);
  } catch (error) {
    return [];
  }
}

function invalidateValidationCache_(tables) {
  try {
    const tableMap = tables || getTables_(true);
    const keys = tableScopedCacheKeys_('validation', tableMap)
      .concat(tableScopedCacheKeys_('formula', tableMap))
      .concat(Object.keys(tableMap).map((key) => cacheKey_('validation', key)));
    if (keys.length) CacheService.getScriptCache().removeAll(keys);
  } catch (error) {
    // Best-effort cache invalidation.
  }
}
