/**
 * Lightweight cache change manifest.
 *
 * Normal app loads compare these versions with the last client manifest and
 * refresh only changed config/tables. Hard refresh remains the recovery path.
 */
const APP_CHANGE_VERSION_PREFIX = 'processAppChange:v1:';

function getAppChangeManifest() {
  assertSpreadsheetAccess_();
  const versions = getAppChangeVersions_();
  const globalVersion = readAppChangeVersion_(versions, 'global');
  const configSourceVersion = getConfigSourceChangeVersion_(versions);
  const tables = getTables_(false);
  const tableVersions = Object.keys(tables || {}).reduce((result, tableKey) => {
    const table = tables[tableKey] || {};
    result[tableKey] = maxAppChangeVersion_([
      globalVersion,
      readAppChangeVersion_(versions, `table:${tableKey}`),
      readAppChangeVersion_(versions, `source:${getAppChangeSourceKey_(table)}`),
    ]);
    return result;
  }, {});

  return sanitizeClientPayload_({
    schemaVersion: 1,
    globalVersion,
    bootstrapVersion: maxAppChangeVersion_([
      globalVersion,
      readAppChangeVersion_(versions, 'bootstrap'),
      configSourceVersion,
    ]),
    configVersion: maxAppChangeVersion_([
      globalVersion,
      readAppChangeVersion_(versions, 'config'),
      configSourceVersion,
    ]),
    tables: tableVersions,
    generatedAt: new Date().toISOString(),
  });
}

function markAllAppCachesChanged_() {
  return writeAppChangeVersion_('global');
}

function markTableCacheChanged_(tableKey) {
  if (!tableKey) return '';
  return writeAppChangeVersion_(`table:${tableKey}`);
}

function markConfigCacheChanged_(configKey) {
  const version = writeAppChangeVersion_('config');
  writeAppChangeVersion_('bootstrap', version);
  if (configKey) writeAppChangeVersion_(`config:${configKey}`, version);
  return version;
}

function markSourceCacheChanged_(spreadsheetId, sheetId) {
  const sourceKey = [spreadsheetId || SPREADSHEET_ID, sheetId == null ? '' : sheetId].join(':');
  if (!sourceKey || /:$/.test(sourceKey)) return '';
  return writeAppChangeVersion_(`source:${sourceKey}`);
}

function onEdit(e) {
  try {
    const range = e && e.range;
    const sheet = range && range.getSheet ? range.getSheet() : null;
    const spreadsheet = e && e.source;
    if (!sheet) return;
    const spreadsheetId = spreadsheet && spreadsheet.getId ? spreadsheet.getId() : SPREADSHEET_ID;
    const sheetId = sheet.getSheetId();
    markSourceCacheChanged_(spreadsheetId, sheetId);
    invalidateEditedSourceTableCaches_(spreadsheetId, sheetId);
  } catch (error) {
    console.warn('App cache edit version skipped:', error && error.message ? error.message : error);
  }
}

function getAppChangeVersions_() {
  try {
    const properties = PropertiesService.getScriptProperties().getProperties();
    return Object.keys(properties || {}).reduce((result, key) => {
      if (key.indexOf(APP_CHANGE_VERSION_PREFIX) !== 0) return result;
      result[key.slice(APP_CHANGE_VERSION_PREFIX.length)] = String(properties[key] || '0');
      return result;
    }, {});
  } catch (error) {
    return {};
  }
}

function readAppChangeVersion_(versions, key) {
  return String(versions && versions[key] || '0');
}

function writeAppChangeVersion_(key, requestedVersion) {
  if (!key) return '';
  const propertyKey = `${APP_CHANGE_VERSION_PREFIX}${key}`;
  try {
    const properties = PropertiesService.getScriptProperties();
    const current = Number(properties.getProperty(propertyKey) || 0);
    const version = String(Math.max(Date.now(), current + 1, Number(requestedVersion || 0)));
    properties.setProperty(propertyKey, version);
    return version;
  } catch (error) {
    return String(Date.now());
  }
}

function maxAppChangeVersion_(values) {
  return String(Math.max.apply(null, (values || []).map((value) => Number(value || 0)).concat([0])));
}

function getAppChangeSourceKey_(table) {
  const source = table || {};
  const sheetId = source.gid != null && source.gid !== ''
    ? source.gid
    : source.apiRange && source.apiRange.sheetId != null
    ? source.apiRange.sheetId
    : '';
  return [source.spreadsheetId || SPREADSHEET_ID, sheetId].join(':');
}

function getConfigSourceChangeVersion_(versions) {
  try {
    const registry = getConfigAppRegistry_(false);
    return maxAppChangeVersion_(Object.keys(registry.items || {}).map((key) => {
      const item = registry.items[key] || {};
      return maxAppChangeVersion_([
        readAppChangeVersion_(versions, `config:${key}`),
        readAppChangeVersion_(versions, `source:${getAppChangeSourceKey_(item.table || {})}`),
      ]);
    }));
  } catch (error) {
    return readAppChangeVersion_(versions, 'config');
  }
}

function invalidateEditedSourceTableCaches_(spreadsheetId, sheetId) {
  try {
    const tables = getCached_(cacheKey_('registry', 'tables')) || {};
    Object.keys(tables).forEach((tableKey) => {
      const table = tables[tableKey] || {};
      if (getAppChangeSourceKey_(table) !== [spreadsheetId || SPREADSHEET_ID, sheetId].join(':')) return;
      CacheService.getScriptCache().removeAll([
        tableScopedCacheKey_('table', tableKey, table),
        tableScopedCacheKey_('tableInitial', tableKey, table),
        tableScopedCacheKey_('tableMeta', tableKey, table),
        tableScopedCacheKey_('validation', tableKey, table),
        tableScopedCacheKey_('formula', tableKey, table),
      ]);
    });
  } catch (error) {
    // The version token is sufficient; clients will force-fetch changed data.
  }
}
