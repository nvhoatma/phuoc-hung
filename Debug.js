/**
 * Debug-only pages and payload builders.
 *
 * Keep route handling in Code.js. This file owns heavy debug payload assembly
 * so the main entrypoint stays focused on access checks and dispatch.
 */
function renderDataLogPage_(params) {
  const payload = buildDataDebugLog_(params || {});
  return renderDebugJsonPage_('Process App Data Log', '?log=data', payload);
}

function renderColumnIndexLogPage_(params) {
  const payload = buildColumnIndexDebugLog_(params || {});
  return renderDebugJsonPage_('Column Index Data Log', '?log=columnIndex', payload);
}

function renderDebugJsonPage_(title, routeHint, payload) {
  const json = JSON.stringify(sanitizeClientPayload_(payload), null, 2);
  const html = `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeDebugHtml_(title)}</title>
  <style>
    body {
      margin: 0;
      padding: 24px;
      background: #f6f8fb;
      color: #111827;
      font: 14px/1.5 ui-monospace, SFMono-Regular, Menlo, Consolas, "Liberation Mono", monospace;
    }
    header {
      margin-bottom: 16px;
      font-family: Arial, sans-serif;
    }
    h1 {
      margin: 0 0 6px;
      font-size: 20px;
      line-height: 1.2;
    }
    p {
      margin: 0;
      color: #64748b;
    }
    pre {
      margin: 0;
      padding: 18px;
      overflow: auto;
      background: #0f172a;
      color: #dbeafe;
      border-radius: 8px;
      box-shadow: 0 14px 35px rgba(15, 23, 42, 0.12);
      white-space: pre-wrap;
      word-break: break-word;
    }
  </style>
</head>
<body>
  <header>
    <h1>${escapeDebugHtml_(title)}</h1>
    <p>Generated from <code>${escapeDebugHtml_(routeHint)}</code>. Add <code>&force=1</code> to bypass caches.</p>
  </header>
  <pre>${escapeDebugHtml_(json)}</pre>
</body>
</html>`;
  return HtmlService.createHtmlOutput(html)
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function buildColumnIndexDebugLog_(params) {
  const timings = [];
  const forceRefresh = isDebugLogTrue_(params.force || params.refresh || params.hardRefresh);
  const measure = (label, callback) => {
    const startedAt = Date.now();
    try {
      return {
        ok: true,
        value: callback(),
        ms: Date.now() - startedAt,
      };
    } catch (error) {
      return {
        ok: false,
        error: debugLogError_(error),
        ms: Date.now() - startedAt,
      };
    } finally {
      timings.push({ label, ms: Date.now() - startedAt });
    }
  };

  const access = measure('access', () => assertSpreadsheetAccess_());
  const columnIndex = measure('columnIndex', () => buildColumnIndexTableDebug_(forceRefresh));
  return {
    ok: access.ok && columnIndex.ok && columnIndex.value && columnIndex.value.ok,
    log: 'columnIndex',
    generatedAt: new Date().toISOString(),
    request: {
      forceRefresh,
    },
    access: access.ok ? {
      spreadsheetId: access.value.spreadsheetId,
      spreadsheetUrl: access.value.spreadsheetUrl,
      userEmail: access.value.userEmail || '',
      userKey: access.value.userKey || '',
    } : access.error,
    columnIndex: columnIndex.ok ? columnIndex.value : {
      ok: false,
      error: columnIndex.error,
    },
    errors: [access, columnIndex]
      .filter((result) => !result.ok)
      .map((result) => result.error),
    performance: {
      timings,
    },
  };
}

function buildColumnIndexTableDebug_(forceRefresh) {
  return buildRawConfigTableDebug_('dataColumn', !!forceRefresh, 0, {
    includeAllRows: true,
    includeAllSourceRows: true,
    displayName: 'Column index',
  });
}

function buildDataDebugLog_(params) {
  const timings = [];
  const forceRefresh = isDebugLogTrue_(params.force || params.refresh || params.hardRefresh);
  const requestedOverviewKey = String(params.key || params.overview || params.overviewKey || '').trim();
  const maxOverviewRows = clampDebugLogNumber_(params.maxRows, 10, 0, 100);
  const maxRawConfigRows = clampDebugLogNumber_(params.rawRows, 200, 0, 2000);
  const measure = (label, callback) => {
    const startedAt = Date.now();
    try {
      return {
        ok: true,
        value: callback(),
        ms: Date.now() - startedAt,
      };
    } catch (error) {
      return {
        ok: false,
        error: debugLogError_(error),
        ms: Date.now() - startedAt,
      };
    } finally {
      timings.push({ label, ms: Date.now() - startedAt });
    }
  };

  const access = measure('access', () => assertSpreadsheetAccess_());
  const tablesResult = measure('tables', () => getTables_(forceRefresh));
  const uiConfigResult = measure('uiConfig', () => getUserAppConfig_(forceRefresh, { deferDataConfig: false }));
  const navigationResult = measure('navigation', () => getNavigation(forceRefresh));
  const projectInfoResult = measure('projectInfo', () => applyUserProjectConfig_(getProjectInfo_(), uiConfigResult.value || {}));
  const overviewKeysResult = measure('overviewKeys', () => getOverviewKeys_());
  const rawConfigTablesResult = measure('rawConfigTables', () => buildRawConfigTablesDebug_(forceRefresh, maxRawConfigRows));

  const tables = tablesResult.value || {};
  const uiConfig = uiConfigResult.value || {};
  const navigation = navigationResult.value || [];
  const overviewKeys = requestedOverviewKey
    ? [requestedOverviewKey]
    : (overviewKeysResult.value || []);
  const overviews = overviewKeys.map((overviewKey) => {
    const result = measure(`overview:${overviewKey}`, () => getOverviewData(overviewKey, forceRefresh, {}));
    return compactDebugOverview_(overviewKey, result, maxOverviewRows);
  });

  return {
    ok: access.ok && tablesResult.ok && uiConfigResult.ok && navigationResult.ok,
    log: 'data',
    generatedAt: new Date().toISOString(),
    request: {
      forceRefresh,
      overviewKey: requestedOverviewKey || '',
      maxOverviewRows,
      maxRawConfigRows,
    },
    access: access.ok ? {
      spreadsheetId: access.value.spreadsheetId,
      spreadsheetUrl: access.value.spreadsheetUrl,
      userEmail: access.value.userEmail || '',
      userKey: access.value.userKey || '',
    } : access.error,
    projectInfo: projectInfoResult.ok ? projectInfoResult.value : projectInfoResult.error,
    summary: {
      tableCount: Object.keys(tables).length,
      navigationRootCount: navigation.length,
      navigationItemCount: flattenBootstrapNavItems_(navigation).length,
      overviewCount: overviewKeys.length,
      pageCount: debugConfigCollectionCount_(uiConfig.pages),
      viewCount: debugConfigCollectionCount_(uiConfig.views),
      formCount: debugConfigCollectionCount_(uiConfig.forms),
      componentCount: debugConfigCollectionCount_(uiConfig.components),
      dataColumnCount: debugConfigCollectionCount_(uiConfig.dataColumns),
      dataServiceCount: debugConfigCollectionCount_(uiConfig.dataServices),
      relationshipCount: debugConfigCollectionCount_(uiConfig.relationships),
    },
    tables: compactDebugTables_(tables),
    navigation: compactDebugNavigation_(navigation),
    uiConfig: compactDebugUiConfig_(uiConfig),
    rawConfigTables: rawConfigTablesResult.ok ? rawConfigTablesResult.value : {
      ok: false,
      error: rawConfigTablesResult.error,
    },
    overviews,
    errors: [access, tablesResult, uiConfigResult, navigationResult, projectInfoResult, overviewKeysResult, rawConfigTablesResult]
      .filter((result) => !result.ok)
      .map((result) => result.error),
    performance: {
      timings,
    },
  };
}

function buildRawConfigTablesDebug_(forceRefresh, maxRows) {
  const columnIndex = buildRawConfigTableDebug_('dataColumn', forceRefresh, maxRows, {
    includeAllRows: true,
    includeAllSourceRows: true,
    displayName: 'Column index',
  });
  return {
    columnIndex,
    dataColumn: {
      aliasOf: 'columnIndex',
      ok: columnIndex.ok,
      rowCount: columnIndex.rowCount || 0,
      returnedRows: columnIndex.returnedRows || 0,
      truncated: !!columnIndex.truncated,
    },
    relationship: buildRawConfigTableDebug_('relationship', forceRefresh, maxRows),
  };
}

function buildRawConfigTableDebug_(configKey, forceRefresh, maxRows, options) {
  try {
    const debugOptions = options || {};
    const item = getConfigItemOrThrow_(configKey, !!forceRefresh);
    if (debugOptions.includeAllSourceRows && item.dataConfig) {
      return buildUnscopedDataConfigTableDebug_(configKey, item, !!forceRefresh, debugOptions);
    }
    const data = getConfigTableData(configKey, !!forceRefresh);
    const rows = data.rows || [];
    const returnedRows = debugOptions.includeAllRows ? rows : rows.slice(0, maxRows);
    return {
      ok: true,
      configKey,
      label: debugOptions.displayName || item.label || '',
      dataConfig: !!item.dataConfig,
      connected: item.connected !== false,
      bindingLabel: item.bindingLabel || '',
      table: compactDebugConfigTableRef_(item.table || data.table || null),
      tableRefs: (item.tableRefs || data.tableRefs || []).map(compactDebugConfigTableRef_),
      headers: data.headers || [],
      rowCount: rows.length,
      returnedRows: returnedRows.length,
      truncated: !debugOptions.includeAllRows && rows.length > maxRows,
      rows: returnedRows.map(compactDebugConfigRow_),
    };
  } catch (error) {
    return {
      ok: false,
      configKey,
      error: debugLogError_(error),
    };
  }
}

function buildUnscopedDataConfigTableDebug_(configKey, item, forceRefresh, options) {
  const debugOptions = options || {};
  const tableRefs = item.tableRefs || [];
  const dataSets = tableRefs.map((tableRef, sourceIndex) => {
    try {
      const data = buildDebugSourceTableRowsData_(tableRef);
      return Object.assign({}, data, { tableRef, sourceIndex });
    } catch (error) {
      return {
        tableRef,
        sourceIndex,
        headers: [],
        rows: [],
        error: debugLogError_(error),
      };
    }
  });
  const headers = getDebugUnionHeaders_(dataSets);
  const rows = dataSets.flatMap((dataSet) => (
    (dataSet.rows || []).map((row) => createDebugConfigAggregateRow_(headers, row, dataSet.tableRef, dataSet.sourceIndex))
  ));
  rows.sort((a, b) => {
    const sourceDelta = Number(a.sourceIndex || 0) - Number(b.sourceIndex || 0);
    if (sourceDelta) return sourceDelta;
    return Number(a.actualRowNumber || a.rowNumber || 0) - Number(b.actualRowNumber || b.rowNumber || 0);
  });
  const enrichedRows = rows.map((row) => enrichDebugConfigRowChoices_(compactDebugConfigRow_(row)));
  return {
    ok: dataSets.every((dataSet) => !dataSet.error),
    configKey,
    label: debugOptions.displayName || item.label || '',
    dataConfig: true,
    unscopedSourceRows: true,
    connected: item.connected !== false,
    bindingLabel: item.bindingLabel || '',
    table: compactDebugConfigTableRef_(item.table || null),
    tableRefs: tableRefs.map(compactDebugConfigTableRef_),
    sourceErrors: dataSets.filter((dataSet) => dataSet.error).map((dataSet) => ({
      sourceIndex: dataSet.sourceIndex,
      table: compactDebugConfigTableRef_(dataSet.tableRef),
      error: dataSet.error,
    })),
    headers,
    rowCount: rows.length,
    returnedRows: rows.length,
    truncated: false,
    rows: enrichedRows,
    choiceRanges: buildDebugChoiceRanges_(enrichedRows),
  };
}

function getDebugUnionHeaders_(dataSets) {
  const seen = {};
  return (dataSets || []).flatMap((dataSet) => dataSet.headers || []).filter((header) => {
    const key = normalizeConfigHeader_(header);
    if (!key || seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function createDebugConfigAggregateRow_(headers, row, tableRef, sourceIndex) {
  const record = {};
  const formulas = {};
  headers.forEach((header) => {
    record[header] = row.record && row.record[header] != null ? row.record[header] : '';
    if (row.formulas && row.formulas[header]) formulas[header] = row.formulas[header];
  });
  return {
    rowNumber: row.rowNumber || '',
    actualRowNumber: row.rowNumber || '',
    sourceIndex,
    sourceSpreadsheetId: tableRef.spreadsheetId || '',
    sourceSheetName: tableRef.sheetName || '',
    id: row.id || `row-${row.rowNumber}`,
    record,
    formulas,
  };
}

function buildDebugSourceTableRowsData_(tableRef) {
  const sheet = getSheet_(tableRef);
  const dataStartColumn = getTableDataStartColumn_(tableRef);
  const headerRow = getTableHeaderRow_(tableRef);
  const dataStartRow = getTableDataStartRow_(tableRef);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < dataStartColumn || lastRow < dataStartRow) {
    return { headers: [], rows: [] };
  }

  const width = lastColumn - dataStartColumn + 1;
  const displayHeaders = sheet.getRange(headerRow, dataStartColumn, 1, width).getDisplayValues()[0] || [];
  const formulaHeaders = sheet.getRange(headerRow, dataStartColumn, 1, width).getFormulas()[0] || [];
  const rowCount = lastRow - dataStartRow + 1;
  const displayRows = sheet.getRange(dataStartRow, dataStartColumn, rowCount, width).getDisplayValues();
  const formulaRows = sheet.getRange(dataStartRow, dataStartColumn, rowCount, width).getFormulas();
  const inferredHeaders = inferDebugHeaders_(displayHeaders, formulaHeaders, displayRows, formulaRows);
  const lastHeaderIndex = inferredHeaders.reduce((last, value, index) => (value ? index : last), -1);
  const headers = lastHeaderIndex >= 0 ? inferredHeaders.slice(0, lastHeaderIndex + 1) : [];

  const rows = [];
  displayRows.forEach((values, rowIndex) => {
    const formulas = formulaRows[rowIndex] || [];
    if (!isNonBlankDisplayRow_(values) && !isNonBlankDisplayRow_(formulas)) return;
    const record = {};
    const formulaRecord = {};
    headers.forEach((header, index) => {
      record[header] = values[index] == null ? '' : values[index];
      if (formulas[index]) formulaRecord[header] = formulas[index];
    });
    rows.push({
      rowNumber: dataStartRow + rowIndex,
      values: headers.map((header, index) => values[index] == null ? '' : values[index]),
      record,
      formulas: formulaRecord,
    });
  });

  return { headers, rows };
}

function inferDebugHeaders_(displayHeaders, formulaHeaders, displayRows, formulaRows) {
  return (displayHeaders || []).map((header, index) => {
    const displayHeader = String(header || '').trim();
    if (displayHeader) return displayHeader;
    const formulaHeader = String((formulaHeaders || [])[index] || '').trim();
    if (formulaHeader) return `Formula ${index + 1}`;
    const values = (displayRows || []).map((row) => String((row || [])[index] || '').trim());
    const formulas = (formulaRows || []).map((row) => String((row || [])[index] || '').trim());
    const sample = values.concat(formulas).find((value) => value);
    if (!sample) return '';
    if (/choices?!/i.test(sample) || /choices?/i.test(sample)) return 'Choices';
    return `Column ${index + 1}`;
  });
}

function compactDebugConfigTableRef_(table) {
  if (!table) return null;
  const source = table.dataConfigSource || {};
  return {
    key: table.key || '',
    name: table.name || '',
    sheetName: table.sheetName || '',
    gid: table.gid || '',
    binding: table.binding || '',
    spreadsheetId: table.spreadsheetId || SPREADSHEET_ID,
    spreadsheetUrl: table.spreadsheetUrl || table.sourceSpreadsheetUrl || '',
    dataConfigKey: table.dataConfigKey || '',
    dataConfigSourceIndex: table.dataConfigSourceIndex,
    source: table.source || '',
    sourceDepartment: source.department || '',
    sourceGroup: source.group || '',
    sourceTables: source.tables || [],
  };
}

function compactDebugConfigRow_(row) {
  return {
    rowNumber: row.rowNumber || '',
    actualRowNumber: row.actualRowNumber || '',
    sourceIndex: row.sourceIndex,
    sourceSpreadsheetId: row.sourceSpreadsheetId || '',
    sourceSheetName: row.sourceSheetName || '',
    id: row.id || '',
    record: row.record || {},
    formulas: row.formulas || {},
  };
}

function enrichDebugConfigRowChoices_(row) {
  const record = row.record || {};
  const formulas = row.formulas || {};
  const choices = getDebugRecordValue_(record, ['choices', 'choice', 'options', 'option_range', 'dropdown_options']) ||
    getDebugRecordValue_(formulas, ['choices', 'choice', 'options', 'option_range', 'dropdown_options']);
  if (!choices) return row;
  row.choicesDebug = resolveDebugChoiceRange_(choices, row.sourceSpreadsheetId);
  return row;
}

function buildDebugChoiceRanges_(rows) {
  return (rows || [])
    .filter((row) => row.choicesDebug)
    .map((row) => ({
      rowNumber: row.rowNumber,
      actualRowNumber: row.actualRowNumber,
      id: row.id,
      table: getDebugRecordValue_(row.record, ['table', 'table_name', 'source_table']),
      column: getDebugRecordValue_(row.record, ['column_name', 'column', 'field', 'field_name']),
      inputType: getDebugRecordValue_(row.record, ['input_type', 'input control', 'control_type', 'input']),
      choices: row.choicesDebug,
    }));
}

function resolveDebugChoiceRange_(value, spreadsheetId) {
  const reference = normalizeValidationRangeReference_(value);
  if (!reference) {
    return {
      raw: value || '',
      reference: '',
      optionCount: 0,
      options: [],
    };
  }
  try {
    const options = getValidationOptionsFromA1_(reference, spreadsheetId || SPREADSHEET_ID);
    return {
      raw: value || '',
      reference,
      spreadsheetId: spreadsheetId || SPREADSHEET_ID,
      optionCount: options.length,
      options,
    };
  } catch (error) {
    return {
      raw: value || '',
      reference,
      spreadsheetId: spreadsheetId || SPREADSHEET_ID,
      optionCount: 0,
      options: [],
      error: debugLogError_(error),
    };
  }
}

function getDebugRecordValue_(record, names) {
  const normalizedNames = (names || []).map(normalizeConfigHeader_);
  const keys = Object.keys(record || {});
  for (let index = 0; index < keys.length; index += 1) {
    const key = keys[index];
    if (normalizedNames.includes(normalizeConfigHeader_(key)) && record[key] !== '') return record[key];
  }
  return '';
}

function compactDebugTables_(tables) {
  return Object.keys(tables || {}).sort().map((key) => {
    const table = tables[key] || {};
    return {
      key,
      name: table.name || '',
      type: table.type || '',
      binding: table.binding || '',
      spreadsheetId: table.spreadsheetId || SPREADSHEET_ID,
      sheetName: table.sheetName || '',
      gid: table.gid || '',
      department: table.department || '',
      group: table.group || '',
      inherited: !!table.inherited,
      readOnly: !!table.readOnly,
      apiTableName: table.apiTableName || '',
      columns: table.columns || [],
      accessFallbackReason: table.accessFallbackReason || '',
    };
  });
}

function compactDebugNavigation_(items) {
  return (items || []).map((item) => ({
    id: item.id || '',
    label: item.label || '',
    pageId: item.pageId || '',
    tableKey: item.tableKey || '',
    overviewKey: item.overviewKey || '',
    linkUrl: item.linkUrl || '',
    disabled: !!item.disabled,
    children: compactDebugNavigation_(item.children || []),
  }));
}

function compactDebugUiConfig_(uiConfig) {
  const config = uiConfig || {};
  return {
    app: config.app || {},
    pages: compactDebugConfigCollection_(config.pages),
    views: compactDebugConfigCollection_(config.views),
    components: compactDebugConfigCollection_(config.components),
    forms: compactDebugConfigCollection_(config.forms),
    dataColumns: compactDebugConfigCollection_(config.dataColumns),
    dataServices: compactDebugConfigCollection_(config.dataServices),
    relationships: compactDebugConfigCollection_(config.relationships),
  };
}

function compactDebugConfigCollection_(collection) {
  const source = collection || {};
  return {
    count: debugConfigCollectionCount_(source),
    keys: {
      byId: source.byId ? Object.keys(source.byId) : [],
      byTableKey: source.byTableKey ? Object.keys(source.byTableKey) : [],
      byPageId: source.byPageId ? Object.keys(source.byPageId) : [],
      byViewId: source.byViewId ? Object.keys(source.byViewId) : [],
      byType: source.byType ? Object.keys(source.byType) : [],
      byPair: source.byPair ? Object.keys(source.byPair) : [],
    },
    value: source,
  };
}

function debugConfigCollectionCount_(collection) {
  const source = collection || {};
  if (Array.isArray(source.items)) return source.items.length;
  if (Array.isArray(source.list)) return source.list.length;
  if (source.byId) return Object.keys(source.byId).length;
  if (source.byTableKey) {
    return Object.keys(source.byTableKey).reduce((count, key) => {
      const value = source.byTableKey[key];
      if (Array.isArray(value)) return count + value.length;
      if (isDebugConfigRecord_(value)) return count + 1;
      if (value && typeof value === 'object') return count + Math.max(1, Object.keys(value).length);
      return count + 1;
    }, 0);
  }
  return 0;
}

function isDebugConfigRecord_(value) {
  if (!value || Array.isArray(value) || typeof value !== 'object') return false;
  return ['id', 'type', 'tableKey', 'tableName', 'pageId', 'viewId', 'enableRead', 'enableCreate']
    .some((key) => Object.prototype.hasOwnProperty.call(value, key));
}

function compactDebugOverview_(overviewKey, result, maxRows) {
  if (!result.ok) {
    return {
      key: overviewKey,
      ok: false,
      error: result.error,
      ms: result.ms,
    };
  }
  const data = result.value || {};
  const table = data.table || null;
  return {
    key: overviewKey,
    ok: true,
    ms: result.ms,
    title: data.title || '',
    subtitle: data.subtitle || '',
    filters: data.filters || [],
    cards: data.cards || [],
    groups: data.groups || [],
    table: table ? {
      title: table.title || '',
      headers: table.headers || [],
      rowCount: table.rows ? table.rows.length : 0,
      rows: (table.rows || []).slice(0, maxRows),
    } : null,
  };
}

function isDebugLogTrue_(value) {
  return ['1', 'true', 'yes', 'y', 'on'].includes(String(value || '').trim().toLowerCase());
}

function clampDebugLogNumber_(value, defaultValue, min, max) {
  const number = Number(value);
  if (!Number.isFinite(number)) return defaultValue;
  return Math.min(max, Math.max(min, Math.floor(number)));
}

function debugLogError_(error) {
  return {
    message: error && error.message ? error.message : String(error),
    stack: error && error.stack ? String(error.stack) : '',
  };
}

function escapeDebugHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
