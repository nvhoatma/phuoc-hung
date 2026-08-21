/**
 * Configuration discovery layer.
 *
 * Reads project info, table definitions, native Sheets table metadata,
 * overview/report config, and shared constants. UI/server features should
 * consume this layer instead of hard-coding sheet names or table ranges.
 */
const DEFAULT_SPREADSHEET_ID = '1t8yqlgoucumfbN8x6W3FP4hBkJVDX1y0WLu9uSozErE';
const SPREADSHEET_ID = getBoundSpreadsheetId_();
const DATA_START_COLUMN = 8; // System-table fallback only.
const HEADER_ROW = 1; // System-table fallback only.
const DATA_START_ROW = 2; // System-table fallback only.
const CACHE_VERSION = 'v66';
const CACHE_TTL_SECONDS = 300;
const UI_CONFIG_CACHE_TTL_SECONDS = 3600;
const CONFIG_TABLE_CACHE_TTL_SECONDS = 3600;
const REGISTRY_CACHE_TTL_SECONDS = 3600;
const TABLE_REGISTRY_CACHE_TTL_SECONDS = 21600;
const PROJECT_INFO_CACHE_TTL_SECONDS = 3600;
const VALIDATION_SCAN_ROWS = 1000;
const TABLE_INDEX_SHEET_NAME = 'Table index';
const TABLE_INDEX_SHEET_ID = 944252806;
const TABLE_INDEX_ALL_SHEET_NAME = '↳ Table index all';
const TABLE_INDEX_ALL_SHEET_ID = 2047597401;
const TABLE_INDEX_CONTEXT = 'process';
const UI_PAGE_CONFIG_SHEET_NAME = 'UI Page config';
const NATIVE_PROCESS_TABLE_PREFIX = 'Process_';
const DATA_TABLE_START_COLUMN = 1;
const DATA_TABLE_HEADER_ROW = 1;
const DATA_TABLE_START_ROW = 2;

function getBoundSpreadsheetId_() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (spreadsheet) return spreadsheet.getId();
  } catch (error) {
    // Standalone deployments fall back to the original project data file.
  }
  return DEFAULT_SPREADSHEET_ID;
}

const SYSTEM_TABLES = {
  tableIndex: {
    key: 'tableIndex',
    name: 'Table index',
    sheetName: TABLE_INDEX_SHEET_NAME,
    gid: TABLE_INDEX_SHEET_ID,
    type: 'system',
  },
  parameterCatalog: {
    key: 'parameterCatalog',
    name: 'App element config',
    sheetName: 'App element config',
    gid: 1046065665,
    type: 'system',
    binding: 'nativeTableName',
    apiTableName: 'App element config',
    dataStartColumn: 6,
    aliases: ['Parameter Catalog'],
  },
  overviewConfig: {
    key: 'overviewConfig',
    name: 'Overview Config',
    sheetName: 'Report config',
    gid: 600009002,
    type: 'system',
  },
  metricConfig: {
    key: 'metricConfig',
    name: 'Metric Config',
    sheetName: 'Metric',
    gid: 600009003,
    type: 'system',
  },
};

function getTables_(forceRefresh) {
  const cacheKey = cacheKey_('registry', 'tables');
  if (!forceRefresh) {
    const cached = getCached_(cacheKey);
    if (cached) return cached;
  }

  const tables = Object.assign({}, SYSTEM_TABLES);
  resolveNativeSystemTables_(tables, forceRefresh);
  readProcessTablesFromIndex_(forceRefresh).forEach((table) => {
    tables[table.key] = table;
  });

  setCached_(cacheKey, tables, TABLE_REGISTRY_CACHE_TTL_SECONDS);
  return tables;
}

function getTableNameToKey_(forceRefresh) {
  const tables = getTables_(forceRefresh);
  return Object.keys(tables).reduce((acc, key) => {
    getTableLookupAliases_(tables[key]).forEach((alias) => {
      acc[normalizeName(alias)] = key;
      acc[normalizeConfigHeader_(alias)] = key;
    });
    (tables[key].aliases || []).forEach((alias) => {
      acc[normalizeName(alias)] = key;
      acc[normalizeName(stripTableLabelIcon_(alias))] = key;
      acc[normalizeConfigHeader_(alias)] = key;
    });
    return acc;
  }, {});
}

function getTableLookupAliases_(table) {
  const rawAliases = [
    table && table.name,
    table && stripTableLabelIcon_(table.name),
    table && table.sheetName,
    table && table.apiTableName,
    table && table.key,
    table && toTableSlug_(table.name),
    table && toTableKey_(table.name),
  ];
  return rawAliases
    .map((value) => String(value || '').trim())
    .filter(Boolean)
    .filter((value, index, values) => values.indexOf(value) === index);
}

function resolveNativeSystemTables_(tables, forceRefresh) {
  Object.keys(tables).forEach((key) => {
    const table = tables[key];
    if (table.binding !== 'nativeTableName') return;

    const nativeTable = getNativeTableByName_(table.apiTableName || table.name, forceRefresh);
    if (!nativeTable) {
      table.binding = '';
      return;
    }

    const sheetMeta = getNativeSheetsById_(forceRefresh)[String(nativeTable.range.sheetId)] || {};
    table.binding = 'nativeTable';
    table.sheetName = sheetMeta.title || table.sheetName || '';
    table.gid = nativeTable.range.sheetId;
    table.spreadsheetId = SPREADSHEET_ID;
    table.spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit#gid=${nativeTable.range.sheetId}`;
    table.apiTableId = nativeTable.tableId;
    table.apiTableName = nativeTable.name;
    table.apiRange = nativeTable.range;
    table.columnProperties = getNativeTableColumnProperties_(nativeTable);
    table.columns = getNativeTableColumns_(nativeTable);
  });
}

function readProcessTablesFromIndex_(forceRefresh) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const configuredNames = getProcessConfiguredTableNameSet_(spreadsheet, forceRefresh);
  const indexRows = [
    readTableIndexRows_(spreadsheet, TABLE_INDEX_ALL_SHEET_NAME, TABLE_INDEX_ALL_SHEET_ID, configuredNames),
    readTableIndexRows_(spreadsheet, TABLE_INDEX_SHEET_NAME, TABLE_INDEX_SHEET_ID, configuredNames),
  ].flat();
  if (!indexRows.length) return [];

  const sourceHints = buildTableIndexSourceHints_(indexRows);
  const scopedIndexRows = Object.keys(configuredNames).length
    ? indexRows.filter((indexRow) => isProcessConfiguredTableName_(indexRow.tableName, configuredNames))
    : indexRows;
  const nativeRegistryCache = {};
  const resolvedByKey = {};
  scopedIndexRows.forEach((indexRow) => {
    const table = resolveProcessTableFromIndexRow_(spreadsheet, indexRow, sourceHints, forceRefresh, nativeRegistryCache);
    if (!table) return;

    const existing = resolvedByKey[table.key];
    if (!existing || shouldReplaceResolvedTable_(existing, table)) {
      resolvedByKey[table.key] = table;
    }
  });

  return Object.keys(resolvedByKey).map((key) => resolvedByKey[key]);
}

function getProcessConfiguredTableNameSet_(spreadsheet, forceRefresh) {
  const sheet = spreadsheet.getSheetByName(UI_PAGE_CONFIG_SHEET_NAME);
  if (!sheet) return {};

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return {};

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  const headers = (values[0] || []).map(normalizeTableIndexHeader_);
  const tableIndex = findHeaderIndex_(headers, ['table_name', 'table', 'source_table']);
  if (tableIndex < 0) return {};
  const pageIndex = findHeaderIndex_(headers, ['page_id', 'page', 'id', 'key']);
  const statusIndex = findHeaderIndex_(headers, ['status', 'active', 'enabled']);
  const activePageIds = getProcessNavigationPageIdSet_(!!forceRefresh);
  const hasActivePageScope = Object.keys(activePageIds).length > 0;

  return values.slice(1).reduce((acc, row) => {
    const status = statusIndex >= 0 ? normalizeName(cellAt_(row, statusIndex)) : '';
    if (status && !['active', 'enabled', 'true', 'yes'].includes(status)) return acc;
    if (hasActivePageScope && pageIndex >= 0) {
      const pageId = normalizeTableLookupName_(cellAt_(row, pageIndex));
      if (pageId && !activePageIds[pageId]) return acc;
    }
    const tableName = cellAt_(row, tableIndex);
    if (!tableName || normalizeName(tableName) === 'n/a') return acc;
    getProcessConfiguredTableLookupNames_(tableName).forEach((name) => {
      if (name) acc[name] = true;
    });
    return acc;
  }, {});
}

function getProcessNavigationPageIdSet_(forceRefresh) {
  try {
    const nativeTable = getNativeTableByName_('Nav config', !!forceRefresh) ||
      getNativeTableByName_('Setting Nav', !!forceRefresh) ||
      getNativeTableByName_('Navigation config', !!forceRefresh);
    if (!nativeTable || !nativeTable.range) return {};
    const sheetMeta = getNativeSheetsById_(!!forceRefresh)[String(nativeTable.range.sheetId)] || {};
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getSheetByNameOrId_(spreadsheet, sheetMeta.title, nativeTable.range.sheetId);
    if (!sheet) return {};

    const startRow = Number(nativeTable.range.startRowIndex || 0) + 1;
    const startColumn = Number(nativeTable.range.startColumnIndex || 0) + 1;
    const rowCount = Math.max(0, Number(nativeTable.range.endRowIndex || 0) - Number(nativeTable.range.startRowIndex || 0));
    const columnCount = Math.max(0, Number(nativeTable.range.endColumnIndex || 0) - Number(nativeTable.range.startColumnIndex || 0));
    if (rowCount < 2 || columnCount < 1) return {};

    const values = sheet.getRange(startRow, startColumn, rowCount, columnCount).getDisplayValues();
    const headers = (values[0] || []).map(normalizeTableIndexHeader_);
    const pageIndex = findHeaderIndex_(headers, ['page', 'page_id', 'target_page']);
    const statusIndex = findHeaderIndex_(headers, ['status', 'active', 'enabled']);
    if (pageIndex < 0) return {};

    return values.slice(1).reduce((acc, row) => {
      const status = statusIndex >= 0 ? normalizeName(cellAt_(row, statusIndex)) : '';
      if (status && !['active', 'enabled', 'true', 'yes'].includes(status)) return acc;
      const pageId = normalizeTableLookupName_(cellAt_(row, pageIndex));
      if (pageId) acc[pageId] = true;
      return acc;
    }, {});
  } catch (error) {
    return {};
  }
}

function isProcessConfiguredTableName_(tableName, configuredNames) {
  return getProcessConfiguredTableLookupNames_(tableName).some((name) => !!configuredNames[name]);
}

function getProcessConfiguredTableLookupNames_(tableName) {
  const names = [
    tableName,
    stripTableLabelIcon_(tableName),
    toTableKey_(tableName),
    toTableSlug_(tableName),
  ];
  return names
    .flatMap((name) => [normalizeTableLookupName_(name), normalizeName(name), normalizeConfigHeader_(name)])
    .filter(Boolean);
}

function readTableIndexRows_(spreadsheet, sheetName, sheetId, configuredNames) {
  const indexSheet = getSheetByNameOrId_(spreadsheet, sheetName, sheetId);
  if (!indexSheet) return [];

  const lastRow = indexSheet.getLastRow();
  const lastColumn = indexSheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];

  const headerValues = indexSheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0] || [];
  const headers = headerValues.map(normalizeTableIndexHeader_);
  const tableIndex = findHeaderIndex_(headers, ['table', 'table_name', 'source_table']);
  if (tableIndex < 0) return [];

  const indexes = {
    context: findHeaderIndex_(headers, ['context', 'process_context', 'scope']),
    department: findHeaderIndex_(headers, ['department', 'dept']),
    group: findHeaderIndex_(headers, ['group', 'group_process_level', 'table_group']),
    table: tableIndex,
    dataType: findHeaderIndex_(headers, ['data_type', 'table_type', 'type']),
    sheetLink: findHeaderIndex_(headers, ['sheet_link', 'sheet_url', 'sheet', 'source_link', 'file_link', 'link']),
    sort: findHeaderIndex_(headers, ['sort', 'sort_order', 'order']),
  };
  const rowCount = lastRow - 1;
  if (rowCount < 1) return [];
  const columnValues = readTableIndexSparseColumns_(indexSheet, indexes, rowCount);

  const carry = { department: '', group: '', context: '' };
  return Array.from({ length: rowCount }, (_, rowOffset) => {
    const rowNumber = rowOffset + 2;
    const context = sparseCellAt_(columnValues, indexes.context, rowOffset);
    const department = sparseCellAt_(columnValues, indexes.department, rowOffset);
    const group = sparseCellAt_(columnValues, indexes.group, rowOffset);
    if (context) carry.context = context;
    if (department) carry.department = department;
    if (group) carry.group = group;

    const tableName = sparseCellAt_(columnValues, indexes.table, rowOffset);
    if (!tableName || normalizeName(tableName) === 'n/a') return null;
    if (configuredNames && Object.keys(configuredNames).length && !isProcessConfiguredTableName_(tableName, configuredNames)) return null;

    const effectiveContext = context || carry.context || TABLE_INDEX_CONTEXT;
    if (indexes.context >= 0 && effectiveContext && normalizeName(effectiveContext) !== TABLE_INDEX_CONTEXT) {
      return null;
    }

    const rawSheetLink = sparseCellAt_(columnValues, indexes.sheetLink, rowOffset);
    const metadata = rawSheetLink && !parseSpreadsheetId_(rawSheetLink)
      ? getTableIndexLinkMetadata_(indexSheet, indexes.sheetLink, rowNumber)
      : {};
    const sheetLink = indexes.sheetLink >= 0
      ? getSheetLinkValue_(rawSheetLink, metadata.richText, metadata.formula)
      : '';

    return {
      sourceSheetName: sheetName,
      sourceSheetId: indexSheet.getSheetId(),
      rowNumber,
      context: effectiveContext,
      department: department || carry.department || '',
      group: group || carry.group || '',
      tableName,
      dataType: sparseCellAt_(columnValues, indexes.dataType, rowOffset),
      sheetLink,
      parsedLink: parseSpreadsheetLink_(sheetLink),
      sortOrder: sparseCellAt_(columnValues, indexes.sort, rowOffset),
    };
  }).filter(Boolean);
}

function readTableIndexSparseColumns_(sheet, indexes, rowCount) {
  const uniqueIndexes = Object.keys(indexes || {})
    .map((key) => Number(indexes[key]))
    .filter((index) => index >= 0)
    .filter((index, position, values) => values.indexOf(index) === position);
  return uniqueIndexes.reduce((acc, index) => {
    acc[index] = sheet.getRange(2, index + 1, rowCount, 1).getDisplayValues().map((row) => row[0]);
    return acc;
  }, {});
}

function sparseCellAt_(columnValues, columnIndex, rowOffset) {
  if (columnIndex < 0 || !columnValues || !columnValues[columnIndex]) return '';
  return String(columnValues[columnIndex][rowOffset] || '').trim();
}

function getTableIndexLinkMetadata_(sheet, sheetLinkIndex, rowNumber) {
  if (sheetLinkIndex < 0 || !rowNumber) return {};
  try {
    const cell = sheet.getRange(rowNumber, sheetLinkIndex + 1);
    return {
      richText: cell.getRichTextValue(),
      formula: cell.getFormula(),
    };
  } catch (error) {
    return {};
  }
}

function resolveProcessTableFromIndexRow_(appSpreadsheet, indexRow, sourceHints, forceRefresh, nativeRegistryCache) {
  const dataType = indexRow.dataType || '';
  const isInherited = normalizeName(dataType).indexOf('inherited') >= 0;
  if (isInherited) {
    const inheritedTable = resolveInheritedProcessTable_(appSpreadsheet, [], {
      tableName: indexRow.tableName,
      dataType,
      sheetName: indexRow.sheetLink || indexRow.tableName,
      context: indexRow.context,
      forceRefresh,
    });
    if (inheritedTable) return enrichIndexTableConfig_(inheritedTable, indexRow);
  }

  const candidates = getCandidateSpreadsheetIdsForIndexRow_(indexRow, sourceHints);
  let accessFallbackReason = '';
  for (let index = 0; index < candidates.length; index += 1) {
    const spreadsheetId = candidates[index];
    let nativeRegistry;
    try {
      nativeRegistry = getNativeRegistryForIndex_(spreadsheetId, forceRefresh, nativeRegistryCache);
    } catch (error) {
      accessFallbackReason = error && error.message ? error.message : String(error);
      continue;
    }

    const nativeTable = resolveNativeTableForIndexRow_(nativeRegistry, indexRow);
    if (nativeTable) {
      const sheetMeta = nativeRegistry.sheetsById[String(nativeTable.range.sheetId)] || {};
      return createNativeProcessTableConfig_({
        tableName: indexRow.tableName,
        dataType,
        nativeTable,
        sheetMeta,
        spreadsheetId,
        source: indexRow.sourceSheetName,
        context: indexRow.context,
        inherited: spreadsheetId !== SPREADSHEET_ID,
        readOnly: false,
        department: indexRow.department,
        group: indexRow.group,
        sortOrder: indexRow.sortOrder,
      });
    }

    const sheetMeta = resolveDataSheetForIndexRow_(nativeRegistry, indexRow);
    if (sheetMeta) {
      return createDataSheetRangeProcessTableConfig_({
        tableName: indexRow.tableName,
        dataType,
        sheetMeta,
        spreadsheetId,
        source: indexRow.sourceSheetName,
        context: indexRow.context,
        department: indexRow.department,
        group: indexRow.group,
        sortOrder: indexRow.sortOrder,
      });
    }
  }

  if (accessFallbackReason && indexRow.parsedLink && indexRow.parsedLink.spreadsheetId) {
    return createUnavailableExternalProcessTableConfig_(indexRow, accessFallbackReason);
  }
  return null;
}

function getNativeRegistryForIndex_(spreadsheetId, forceRefresh, nativeRegistryCache) {
  const targetSpreadsheetId = spreadsheetId || SPREADSHEET_ID;
  const cache = nativeRegistryCache || {};
  if (cache[targetSpreadsheetId]) return cache[targetSpreadsheetId];
  const registry = getNativeSpreadsheetTablesById_(targetSpreadsheetId, forceRefresh);
  cache[targetSpreadsheetId] = registry;
  return registry;
}

function enrichIndexTableConfig_(table, indexRow) {
  table.department = indexRow.department || table.department || '';
  table.group = indexRow.group || table.group || '';
  table.sortOrder = indexRow.sortOrder || table.sortOrder || '';
  table.source = indexRow.sourceSheetName || table.source || '';
  return table;
}

function shouldReplaceResolvedTable_(existing, candidate) {
  if (!existing) return true;
  if (candidate.binding === 'nativeTable' && existing.binding !== 'nativeTable') return true;
  if (candidate.binding === 'inheritedNativeTable' && existing.binding === 'dataSheetRange') return true;
  if (candidate.department && !existing.department) return true;
  if (candidate.spreadsheetId !== SPREADSHEET_ID && existing.spreadsheetId === SPREADSHEET_ID) return true;
  return false;
}

function buildTableIndexSourceHints_(indexRows) {
  const hints = { all: [], byDepartment: {}, bySourceSheet: {} };
  indexRows.forEach((indexRow) => {
    const spreadsheetId = indexRow.parsedLink && indexRow.parsedLink.spreadsheetId;
    if (!spreadsheetId) return;

    addUniqueString_(hints.all, spreadsheetId);
    addSourceHint_(hints.bySourceSheet, indexRow.sourceSheetName, spreadsheetId);
    addSourceHint_(hints.byDepartment, indexRow.department, spreadsheetId);
  });
  return hints;
}

function addSourceHint_(target, key, spreadsheetId) {
  const normalizedKey = normalizeName(key || '');
  if (!normalizedKey) return;
  if (!target[normalizedKey]) target[normalizedKey] = [];
  addUniqueString_(target[normalizedKey], spreadsheetId);
}

function getCandidateSpreadsheetIdsForIndexRow_(indexRow, sourceHints) {
  const candidates = [];
  if (indexRow.parsedLink && indexRow.parsedLink.spreadsheetId) {
    addUniqueString_(candidates, indexRow.parsedLink.spreadsheetId);
  }

  const departmentKey = normalizeName(indexRow.department || '');
  const sourceSheetKey = normalizeName(indexRow.sourceSheetName || '');
  (sourceHints.byDepartment[departmentKey] || []).forEach((spreadsheetId) => addUniqueString_(candidates, spreadsheetId));
  (sourceHints.bySourceSheet[sourceSheetKey] || []).forEach((spreadsheetId) => addUniqueString_(candidates, spreadsheetId));
  if (!departmentKey) {
    sourceHints.all.forEach((spreadsheetId) => addUniqueString_(candidates, spreadsheetId));
  }
  addUniqueString_(candidates, SPREADSHEET_ID);
  return candidates;
}

function resolveNativeTableForIndexRow_(nativeRegistry, indexRow) {
  const nativeTablesByName = (nativeRegistry.tables || []).reduce((acc, table) => {
    getNativeTableLookupNames_(table.name).forEach((name) => {
      acc[name] = table;
    });
    return acc;
  }, {});
  const directGid = indexRow.parsedLink && indexRow.parsedLink.gid ? Number(indexRow.parsedLink.gid) : 0;
  const nativeTable = resolveNativeProcessTable_(nativeTablesByName, indexRow.tableName);
  if (!nativeTable) return null;
  if (directGid && Number(nativeTable.range && nativeTable.range.sheetId) !== directGid) {
    const sameSheetTables = (nativeRegistry.tables || []).filter((table) => (
      Number(table.range && table.range.sheetId) === directGid
    ));
    if (sameSheetTables.length === 1) return sameSheetTables[0];
    return null;
  }
  return nativeTable;
}

function resolveDataSheetForIndexRow_(nativeRegistry, indexRow) {
  const directGid = indexRow.parsedLink && indexRow.parsedLink.gid ? String(indexRow.parsedLink.gid) : '';
  if (directGid && nativeRegistry.sheetsById[directGid]) return nativeRegistry.sheetsById[directGid];

  const expectedNames = getDataSheetLookupNames_(indexRow.tableName);
  return Object.keys(nativeRegistry.sheetsById || {}).reduce((found, sheetId) => {
    if (found) return found;
    const sheetMeta = nativeRegistry.sheetsById[sheetId] || {};
    return expectedNames.indexOf(normalizeTableLookupName_(sheetMeta.title)) >= 0 ? sheetMeta : null;
  }, null);
}

function getDataSheetLookupNames_(tableName) {
  return getNativeTableLookupNames_(tableName)
    .concat(getNativeTableLookupNames_(toTableSlug_(tableName)))
    .concat(getNativeTableLookupNames_(toTableKey_(tableName)));
}

function createNativeProcessTableConfig_(options) {
  const nativeTable = options.nativeTable;
  const sheetMeta = options.sheetMeta || {};
  const spreadsheetId = options.spreadsheetId || SPREADSHEET_ID;
  const sourceSpreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${nativeTable.range.sheetId}`;
  return {
    key: toTableKey_(options.tableName),
    name: options.tableName,
    sheetName: sheetMeta.title || options.sheetName || '',
    gid: nativeTable.range.sheetId,
    type: options.dataType || '',
    spreadsheetId,
    spreadsheetUrl: sourceSpreadsheetUrl,
    sourceSpreadsheetUrl,
    source: options.source || TABLE_INDEX_SHEET_NAME,
    context: options.context || '',
    department: options.department || '',
    group: options.group || '',
    sortOrder: options.sortOrder || '',
    binding: options.inherited ? 'inheritedNativeTable' : 'nativeTable',
    inherited: !!options.inherited,
    readOnly: !!options.readOnly,
    apiTableId: nativeTable.tableId,
    apiTableName: nativeTable.name,
    apiRange: nativeTable.range,
    columnProperties: getNativeTableColumnProperties_(nativeTable),
    columns: getNativeTableColumns_(nativeTable),
    localSheetName: options.localSheetName || '',
    localGid: options.localGid || '',
    importRangeFormula: options.importRangeFormula || '',
    importRangeSourceRange: options.importRangeSourceRange || '',
  };
}

function createDataSheetRangeProcessTableConfig_(options) {
  const sheetMeta = options.sheetMeta || {};
  const spreadsheetId = options.spreadsheetId || SPREADSHEET_ID;
  const gid = Number(sheetMeta.sheetId || 0);
  const sourceSpreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${gid}`;
  return {
    key: toTableKey_(options.tableName),
    name: options.tableName,
    sheetName: sheetMeta.title || options.tableName || '',
    gid,
    type: options.dataType || '',
    spreadsheetId,
    spreadsheetUrl: sourceSpreadsheetUrl,
    sourceSpreadsheetUrl,
    source: options.source || TABLE_INDEX_SHEET_NAME,
    context: options.context || '',
    department: options.department || '',
    group: options.group || '',
    sortOrder: options.sortOrder || '',
    binding: 'dataSheetRange',
    inherited: spreadsheetId !== SPREADSHEET_ID,
    readOnly: false,
    dataStartColumn: DATA_TABLE_START_COLUMN,
    headerRow: DATA_TABLE_HEADER_ROW,
    dataStartRow: DATA_TABLE_START_ROW,
  };
}

function createUnavailableExternalProcessTableConfig_(indexRow, reason) {
  const parsedLink = indexRow.parsedLink || {};
  const spreadsheetId = parsedLink.spreadsheetId || SPREADSHEET_ID;
  const gid = Number(parsedLink.gid || 0);
  return {
    key: toTableKey_(indexRow.tableName),
    name: indexRow.tableName,
    sheetName: indexRow.tableName,
    gid,
    type: indexRow.dataType || '',
    spreadsheetId,
    spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit${gid ? `#gid=${gid}` : ''}`,
    sourceSpreadsheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit${gid ? `#gid=${gid}` : ''}`,
    source: indexRow.sourceSheetName || TABLE_INDEX_SHEET_NAME,
    context: indexRow.context || '',
    department: indexRow.department || '',
    group: indexRow.group || '',
    sortOrder: indexRow.sortOrder || '',
    binding: 'dataSheetRange',
    inherited: true,
    readOnly: true,
    dataStartColumn: DATA_TABLE_START_COLUMN,
    headerRow: DATA_TABLE_HEADER_ROW,
    dataStartRow: DATA_TABLE_START_ROW,
    accessFallbackReason: reason || 'External data spreadsheet is not available.',
  };
}

function resolveInheritedProcessTable_(spreadsheet, row, options) {
  const localSheet = findInheritedSheet_(spreadsheet, options.sheetName || options.tableName);
  if (!localSheet) return null;

  const importCell = localSheet.getRange(HEADER_ROW, DATA_START_COLUMN);
  const importRange = parseImportRangeFormula_(importCell.getFormula());
  if (!importRange || !importRange.spreadsheetId || !importRange.tableName) {
    return createInheritedLocalProcessTableConfig_(options, localSheet, importCell, importRange, 'No valid IMPORTRANGE source pointer found in H1.');
  }

  try {
    const nativeRegistry = getNativeSpreadsheetTablesById_(importRange.spreadsheetId, options.forceRefresh);
    const nativeTablesByName = nativeRegistry.tables.reduce((acc, table) => {
      getNativeTableLookupNames_(table.name).forEach((name) => {
        acc[name] = table;
      });
      return acc;
    }, {});
    const nativeTable = resolveNativeProcessTable_(
      nativeTablesByName,
      importRange.tableName || options.tableName
    ) || resolveNativeProcessTable_(nativeTablesByName, options.tableName);
    if (!nativeTable) {
      return createInheritedLocalProcessTableConfig_(options, localSheet, importCell, importRange, 'Source native table not found; using imported local snapshot.');
    }

    const sheetMeta = nativeRegistry.sheetsById[String(nativeTable.range.sheetId)] || {};
    return createNativeProcessTableConfig_({
      tableName: options.tableName,
      dataType: options.dataType,
      nativeTable,
      sheetMeta,
      spreadsheetId: importRange.spreadsheetId,
      source: 'IMPORTRANGE',
      context: options.context || '',
      inherited: true,
      readOnly: false,
      localSheetName: localSheet.getName(),
      localGid: localSheet.getSheetId(),
      importRangeFormula: importCell.getFormula(),
      importRangeSourceRange: importRange.rangeReference,
    });
  } catch (error) {
    return createInheritedLocalProcessTableConfig_(
      options,
      localSheet,
      importCell,
      importRange,
      error && error.message ? error.message : String(error)
    );
  }
}

function createInheritedLocalProcessTableConfig_(options, localSheet, importCell, importRange, reason) {
  return {
    key: toTableKey_(options.tableName),
    name: options.tableName,
    sheetName: localSheet.getName(),
    gid: localSheet.getSheetId(),
    type: options.dataType || '',
    spreadsheetId: SPREADSHEET_ID,
    spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit#gid=${localSheet.getSheetId()}`,
    sourceSpreadsheetUrl: importRange && importRange.spreadsheetId
      ? `https://docs.google.com/spreadsheets/d/${importRange.spreadsheetId}/edit`
      : '',
    source: 'IMPORTRANGE local snapshot',
    context: options.context || '',
    binding: 'inheritedRange',
    inherited: true,
    readOnly: true,
    dataStartColumn: DATA_START_COLUMN,
    headerRow: HEADER_ROW,
    dataStartRow: DATA_START_ROW,
    localSheetName: localSheet.getName(),
    localGid: localSheet.getSheetId(),
    importRangeFormula: importCell.getFormula(),
    importRangeSourceRange: importRange ? importRange.rangeReference || '' : '',
    accessFallbackReason: reason || '',
  };
}

function findInheritedSheet_(spreadsheet, sheetName) {
  const direct = sheetName ? spreadsheet.getSheetByName(sheetName) : null;
  if (direct) return direct;

  const normalized = normalizeSheetLabel_(sheetName);
  return spreadsheet.getSheets().find((sheet) => (
    normalizeSheetLabel_(sheet.getName()) === normalized
  )) || null;
}

function parseImportRangeFormula_(formula) {
  const text = String(formula || '').trim();
  const match = text.match(/IMPORTRANGE\s*\(\s*"([^"]+)"\s*[,;]\s*"([^"]+)"\s*\)/i);
  if (!match) return null;

  const spreadsheetId = parseSpreadsheetId_(match[1]);
  const rangeReference = match[2].trim();
  const tableReference = rangeReference.match(/^'?(.+?)'?\s*\[#ALL\]$/i);
  return {
    spreadsheetId,
    rangeReference,
    tableName: tableReference ? tableReference[1].trim() : '',
  };
}

function parseSpreadsheetId_(value) {
  const text = String(value || '').trim();
  const urlMatch = text.match(/\/spreadsheets\/d\/([A-Za-z0-9_-]+)/);
  if (urlMatch) return urlMatch[1];
  const idMatch = text.match(/[A-Za-z0-9_-]{20,}/);
  return idMatch ? idMatch[0] : '';
}

function parseSpreadsheetLink_(value) {
  const text = String(value || '').trim();
  return {
    spreadsheetId: parseSpreadsheetId_(text),
    gid: parseSheetGid_(text),
    url: text,
  };
}

function parseSheetGid_(value) {
  const text = String(value || '').trim();
  const gidMatch = text.match(/[?#&]gid=(\d+)/);
  return gidMatch ? gidMatch[1] : '';
}

function getSheetLinkValue_(displayValue, richText, formula) {
  if (displayValue && parseSpreadsheetId_(displayValue)) return displayValue;

  if (richText && typeof richText.getLinkUrl === 'function') {
    const url = richText.getLinkUrl();
    if (url) return url;
  }

  if (formula && parseSpreadsheetId_(formula)) return formula;
  return displayValue;
}

function cellAt_(row, index) {
  if (index == null || index < 0) return '';
  return String((row || [])[index] == null ? '' : (row || [])[index]).trim();
}

function normalizeTableIndexHeader_(value) {
  return stripTableLabelIcon_(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function findHeaderIndex_(headers, aliases) {
  return (aliases || []).reduce((found, alias) => {
    if (found >= 0) return found;
    return headers.indexOf(alias);
  }, -1);
}

function addUniqueString_(target, value) {
  const text = String(value || '').trim();
  if (!text || target.indexOf(text) >= 0) return;
  target.push(text);
}

function getNativeProcessTablesByName_(forceRefresh) {
  return getNativeSpreadsheetTables_(forceRefresh).tables.reduce((acc, table) => {
    getNativeTableLookupNames_(table.name).forEach((name) => {
      acc[name] = table;
    });
    return acc;
  }, {});
}

function getNativeSheetsById_(forceRefresh) {
  return getNativeSpreadsheetTables_(forceRefresh).sheetsById;
}

function getNativeTableByName_(tableName, forceRefresh) {
  const normalized = normalizeTableLookupName_(tableName);
  return getNativeSpreadsheetTables_(forceRefresh).tables.find((table) => (
    normalizeTableLookupName_(table.name) === normalized
  )) || null;
}

function getNativeSpreadsheetTables_(forceRefresh) {
  return getNativeSpreadsheetTablesById_(SPREADSHEET_ID, forceRefresh);
}

function getNativeSpreadsheetTablesById_(spreadsheetId, forceRefresh) {
  const targetSpreadsheetId = spreadsheetId || SPREADSHEET_ID;
  const cacheKey = cacheKey_('registry', `nativeTables:${targetSpreadsheetId}`);
  const cached = getCached_(cacheKey);
  if (!forceRefresh) {
    if (cached) return cached;
  }

  let spreadsheet;
  try {
    spreadsheet = sheetsApiGet_(
      targetSpreadsheetId,
      'sheets(properties(sheetId,title),tables(tableId,name,range,columnProperties(columnIndex,columnName,columnType,dataValidationRule)))'
    );
  } catch (error) {
    if (cached && isQuotaError_(error)) return Object.assign({}, cached, { stale: true, quotaLimited: true });
    throw toFriendlyQuotaError_(error);
  }
  const sheets = spreadsheet.sheets || [];
  const result = {
    sheetsById: sheets.reduce((acc, sheet) => {
      if (sheet.properties) acc[String(sheet.properties.sheetId)] = sheet.properties;
      return acc;
    }, {}),
    tables: sheets.flatMap((sheet) => (sheet.tables || []).map((table) => {
      table.range = table.range || {};
      if (table.range.sheetId == null && sheet.properties) {
        table.range.sheetId = sheet.properties.sheetId;
      }
      return table;
    })),
  };

  setCached_(cacheKey, result, TABLE_REGISTRY_CACHE_TTL_SECONDS);
  return result;
}

function resolveNativeProcessTable_(nativeTablesByName, tableName) {
  const names = getNativeTableLookupNames_(tableName)
    .concat(getNativeTableLookupNames_(`${NATIVE_PROCESS_TABLE_PREFIX}${toTableKey_(tableName)}`))
    .concat(getNativeTableLookupNames_(`${NATIVE_PROCESS_TABLE_PREFIX}${toTableSlug_(tableName)}`));
  return names.reduce((found, name) => found || nativeTablesByName[name] || null, null);
}

function getNativeTableLookupNames_(name) {
  const raw = String(name || '');
  const withoutPrefix = raw.replace(/^process[_\s-]+/i, '');
  return [raw, withoutPrefix].map(normalizeTableLookupName_);
}

function normalizeTableLookupName_(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function getNativeTableColumns_(table) {
  return getNativeTableColumnProperties_(table)
    .map((column) => column.columnName || '')
    .filter(Boolean);
}

function getNativeTableColumnProperties_(table) {
  return (table.columnProperties || [])
    .sort((a, b) => Number(a.columnIndex || 0) - Number(b.columnIndex || 0))
    .map((column) => ({
      columnIndex: column.columnIndex,
      columnName: column.columnName || '',
      columnType: column.columnType || '',
      dataValidationRule: column.dataValidationRule || null,
    }));
}

function getSheetByNameOrId_(spreadsheet, sheetName, sheetId) {
  let sheet = sheetName ? spreadsheet.getSheetByName(sheetName) : null;
  if (!sheet && sheetId != null) {
    sheet = spreadsheet.getSheets().find((candidate) => candidate.getSheetId() === Number(sheetId));
  }
  return sheet || null;
}

function normalizeConfigHeader_(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g, '_');
}

function stripTableLabelIcon_(value) {
  return String(value || '').replace(/^[^A-Za-z0-9\u00C0-\u1EF9]+/, '').trim();
}

function normalizeSheetLabel_(value) {
  return String(value || '').replace(/^[^A-Za-z0-9\u00C0-\u1EF9]+/, '').trim().toLowerCase();
}

function toTableKey_(tableName) {
  const words = String(tableName || '').trim().split(/[^A-Za-z0-9]+/).filter(Boolean);
  if (!words.length) return Utilities.getUuid().replace(/-/g, '');
  return words.map((word, index) => {
    const lower = word.toLowerCase();
    return index === 0 ? lower : lower.charAt(0).toUpperCase() + lower.slice(1);
  }).join('');
}

function toTableSlug_(tableName) {
  return String(tableName || '').trim().replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}
