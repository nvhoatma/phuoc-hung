/**
 * Table data and CRUD service.
 *
 * Owns reading table rows, field validation schema, save/delete behavior,
 * native table range handling, and data cache invalidation after writes.
 * UI code should call getTableData/saveRecord/deleteRecord through Code.js.
 */
const TABLE_ROW_LIMIT = 2000;
const TABLE_INITIAL_ROW_LIMIT = 100;
const TABLE_READ_CHUNK_SIZE = 500;

function getTableData(tableKey, forceRefresh) {
  assertSpreadsheetAccess_();
  const table = getTableOrThrow_(tableKey, forceRefresh);
  const cacheKey = tableScopedCacheKey_('table', tableKey, table);
  const cached = getCached_(cacheKey);
  if (!forceRefresh) {
    if (cached) return sanitizeClientPayload_(attachFieldSchema_(cached, table, false));
  }

  let data;
  try {
    data = buildTableData_(table, forceRefresh);
  } catch (error) {
    if (isQuotaError_(error) && cached) {
      return sanitizeClientPayload_(Object.assign(
        attachFieldSchema_(cached, table, false),
        { stale: true, quotaLimited: true, message: quotaErrorMessage_() }
      ));
    }
    throw toFriendlyQuotaError_(error);
  }
  setCached_(cacheKey, data);
  return sanitizeClientPayload_(data);
}

function getTableDataInitial(tableKey, forceRefresh) {
  assertSpreadsheetAccess_();
  const table = getTableOrThrow_(tableKey, forceRefresh);
  const fullCacheKey = tableScopedCacheKey_('table', tableKey, table);
  const initialCacheKey = tableScopedCacheKey_('tableInitial', tableKey, table);
  if (!forceRefresh) {
    const cachedInitial = getCached_(initialCacheKey);
    if (cachedInitial) return sanitizeClientPayload_(cachedInitial);
    const cachedFull = getCached_(fullCacheKey);
    if (cachedFull) {
      const initialFromFull = createInitialTableData_(attachFieldSchema_(cachedFull, table, false));
      setCached_(initialCacheKey, initialFromFull);
      return sanitizeClientPayload_(initialFromFull);
    }
  }

  let data;
  try {
    const rowsData = buildTableRowsData_(table, {
      limit: TABLE_INITIAL_ROW_LIMIT,
    });
    data = createInitialTableData_(attachLightFieldSchema_(rowsData));
  } catch (error) {
    const cachedInitial = getCached_(initialCacheKey);
    if (isQuotaError_(error) && cachedInitial) {
      return sanitizeClientPayload_(Object.assign(
        cachedInitial,
        { stale: true, quotaLimited: true, message: quotaErrorMessage_() }
      ));
    }
    throw toFriendlyQuotaError_(error);
  }
  setCached_(initialCacheKey, data);
  return sanitizeClientPayload_(data);
}

function createInitialTableData_(data) {
  const headers = data && data.headers ? data.headers : [];
  const sourceRows = data && data.rows ? data.rows : [];
  const rows = sourceRows.length > TABLE_INITIAL_ROW_LIMIT
    ? sourceRows.slice(sourceRows.length - TABLE_INITIAL_ROW_LIMIT)
    : sourceRows;
  const sourceWindow = data && data.dataWindow ? data.dataWindow : {};
  return Object.assign({}, data || {}, {
    rows,
    summary: createSummary_(headers, rows),
    dataWindow: Object.assign({}, sourceWindow, {
      mode: sourceWindow.mode || 'latest',
      initial: true,
      rowLimit: TABLE_INITIAL_ROW_LIMIT,
      fullRowLimit: TABLE_ROW_LIMIT,
      returnedRows: rows.length,
      nextOffset: rows.length,
      truncated: !!sourceWindow.truncated || sourceRows.length > rows.length,
    }),
  });
}

function getTablePage(payload) {
  const startedAt = Date.now();
  assertSpreadsheetAccess_();
  const request = payload || {};
  const tableKey = request.tableKey || '';
  if (!tableKey) throw new Error('Missing tableKey.');
  const table = getTableOrThrow_(tableKey, false);
  const pageRequest = normalizeTablePageRequest_(request);
  const data = hasServerTableProcessing_(pageRequest)
    ? buildProcessedTablePage_(table, pageRequest)
    : buildCursorTablePage_(table, pageRequest);
  const result = attachLightFieldSchema_(data);
  result.performance = {
    operation: 'getTablePage',
    mode: data && data.serverPage ? data.serverPage.mode : 'paged',
    sourceRows: data && data.serverPage ? Number(data.serverPage.sourceRows || 0) : 0,
    returnedRows: data && data.rows ? data.rows.length : 0,
    totalMs: Date.now() - startedAt,
  };
  console.log(JSON.stringify(Object.assign({ tableKey }, result.performance)));
  return sanitizeClientPayload_(result);
}

function getTableDataPage(tableKey, request) {
  return getTablePage(Object.assign({}, request || {}, { tableKey }));
}

function buildProcessedTablePage_(table, payload) {
  const sourceData = buildTableRowsData_(table, { limit: TABLE_ROW_LIMIT });
  const processed = applyServerTableQuery_(sourceData.rows || [], sourceData.headers || [], payload);
  const pageCount = Math.max(1, Math.ceil(processed.rows.length / payload.pageSize));
  const page = Math.min(Math.max(1, payload.page), pageCount);
  const startIndex = (page - 1) * payload.pageSize;
  const sourceEndIndex = payload.sort && payload.sort.header
    ? Math.min(processed.rows.length, startIndex + payload.pageSize)
    : Math.max(0, processed.rows.length - startIndex);
  const sourceStartIndex = payload.sort && payload.sort.header
    ? startIndex
    : Math.max(0, sourceEndIndex - payload.pageSize);
  const pageRows = processed.rows.slice(sourceStartIndex, sourceEndIndex);
  return Object.assign({}, sourceData, {
    rows: pageRows,
    summary: createSummary_(sourceData.headers || [], processed.rows),
    filterFacets: createServerFilterFacets_(processed.rows, sourceData.headers || [], payload.filterHeaders),
    serverPage: {
      mode: 'processed',
      page,
      pageSize: payload.pageSize,
      pageCount,
      totalRows: processed.rows.length,
      totalEstimate: processed.rows.length,
      totalExact: true,
      sourceRows: (sourceData.rows || []).length,
      startIndex,
      endIndex: Math.min(processed.rows.length, startIndex + pageRows.length),
      cursor: startIndex,
      nextCursor: page < pageCount ? startIndex + pageRows.length : '',
      query: payload.query,
      filters: payload.filters,
      sort: payload.sort,
    },
  });
}

function buildCursorTablePage_(table, payload) {
  const offset = Math.max(0, Number(payload.cursor || ((payload.page - 1) * payload.pageSize)));
  const snapshotEndRow = resolveTableSnapshotEndRow_(table, payload.snapshotEndRow);
  const sourceData = buildTableRowsData_(table, {
    limit: payload.pageSize + 1,
    offset,
    endRow: snapshotEndRow,
  });
  const rows = sourceData.rows || [];
  // Tail reads return rows in ascending sheet order. The extra row is a
  // look-ahead before the requested window, so retain the newest pageSize rows.
  const pageRows = rows.slice(Math.max(0, rows.length - payload.pageSize));
  const hasMore = rows.length > payload.pageSize || !!(sourceData.dataWindow && sourceData.dataWindow.truncated);
  const totalEstimate = estimateTableRowCount_(table, snapshotEndRow);
  return Object.assign({}, sourceData, {
    rows: pageRows,
    summary: createSummary_(sourceData.headers || [], pageRows),
    filterFacets: createServerFilterFacets_(pageRows, sourceData.headers || [], payload.filterHeaders),
    serverPage: {
      mode: 'cursor',
      page: Math.max(1, Math.floor(offset / payload.pageSize) + 1),
      pageSize: payload.pageSize,
      pageCount: Math.max(1, Math.ceil(totalEstimate / payload.pageSize)),
      totalRows: totalEstimate,
      totalEstimate,
      totalExact: false,
      sourceRows: totalEstimate,
      snapshotEndRow,
      startIndex: offset,
      endIndex: offset + pageRows.length,
      cursor: offset,
      nextCursor: hasMore ? offset + pageRows.length : '',
      query: payload.query,
      filters: payload.filters,
      sort: payload.sort,
    },
    dataWindow: Object.assign({}, sourceData.dataWindow || {}, {
      returnedRows: pageRows.length,
      rowLimit: payload.pageSize,
      offset,
      nextOffset: hasMore ? offset + pageRows.length : '',
      truncated: hasMore,
    }),
  });
}

/**
 * Reads rows appended after a completed snapshot using physical row boundaries.
 * Unlike offset paging from the tail, this cursor cannot shift when new rows are
 * appended while the client is hydrating the table.
 */
function getTableDataDelta(payload) {
  const startedAt = Date.now();
  assertSpreadsheetAccess_();
  const request = payload || {};
  const tableKey = String(request.tableKey || '').trim();
  if (!tableKey) throw new Error('Missing tableKey.');
  const table = getTableOrThrow_(tableKey, false);
  const afterRowNumber = Math.max(0, Math.round(Number(request.afterRowNumber || 0)));
  const scanLimit = Math.max(50, Math.min(1000, Math.round(Number(request.scanLimit || 1000))));
  const bounds = getTablePhysicalBounds_(table, true);
  const snapshotEndRow = bounds.endRow;
  const scanStartRow = Math.max(bounds.dataStartRow, afterRowNumber + 1);
  const scanEndRow = Math.min(snapshotEndRow, scanStartRow + scanLimit - 1);
  const rows = scanStartRow <= scanEndRow && bounds.width
    ? bounds.sheet
      .getRange(scanStartRow, bounds.startColumn, scanEndRow - scanStartRow + 1, bounds.width)
      .getDisplayValues()
      .reduce((result, values, index) => {
        if (isNonBlankDisplayRow_(values)) {
          result.push(createDataRowPayload_(bounds.headers, scanStartRow + index, values));
        }
        return result;
      }, [])
    : [];
  const scannedThroughRow = scanStartRow <= scanEndRow ? scanEndRow : afterRowNumber;
  return sanitizeClientPayload_({
    tableKey,
    rows,
    deltaPage: {
      afterRowNumber,
      scanStartRow,
      scannedThroughRow,
      snapshotEndRow,
      hasMore: scannedThroughRow < snapshotEndRow,
    },
    performance: {
      operation: 'getTableDataDelta',
      scannedRows: scanStartRow <= scanEndRow ? scanEndRow - scanStartRow + 1 : 0,
      returnedRows: rows.length,
      totalMs: Date.now() - startedAt,
    },
  });
}

function getTableMeta(tableKey, forceRefresh) {
  assertSpreadsheetAccess_();
  const table = getTableOrThrow_(tableKey, forceRefresh);
  const cacheKey = tableScopedCacheKey_('tableMeta', tableKey, table);
  if (!forceRefresh) {
    const cached = getCached_(cacheKey);
    if (cached) return sanitizeClientPayload_(cached);
  }

  const meta = buildTableMeta_(tableKey, table, !!forceRefresh);
  setCached_(cacheKey, meta, CACHE_TTL_SECONDS);
  return sanitizeClientPayload_(meta);
}

function getTableSchema(tableKey, forceRefresh) {
  assertSpreadsheetAccess_();
  const table = getTableOrThrow_(tableKey, false);
  const headers = readTableMetaHeaders_(table);
  const schema = getTableFieldSchema_(table, headers, !!forceRefresh);
  return sanitizeClientPayload_({
    tableKey,
    headers,
    fields: schema.fields || [],
    validationByHeader: schema.validationByHeader || {},
    formulaAware: !!schema.formulaAware,
    formulaHeaders: schema.formulaHeaders || [],
    schemaGeneratedAt: schema.generatedAt,
    schemaMode: 'full',
  });
}

function buildTableMeta_(tableKey, table, forceRefresh) {
  const headers = readTableMetaHeaders_(table);
  const rowCount = estimateTableRowCount_(table);
  const ui = getTableMetaUiConfig_(tableKey, forceRefresh);
  const keyFields = getTableMetaKeyFields_(headers, ui);
  return {
    table: compactTableForMeta_(tableKey, table),
    headers,
    rowCount,
    keyFields,
    viewConfig: ui,
    summary: {
      rows: rowCount,
      columns: headers.length,
      truncated: rowCount > TABLE_ROW_LIMIT,
      rowLimit: TABLE_ROW_LIMIT,
    },
    latestUpdated: '',
    generatedAt: new Date().toISOString(),
  };
}

function readTableMetaHeaders_(table) {
  if (isNativeReadableTable_(table)) return getNativeHeaders_(table);
  const sheet = getSheet_(table);
  return getSystemTableLayout_(sheet, table).headers;
}

function estimateTableRowCount_(table, snapshotEndRow) {
  const requestedEndRow = Number(snapshotEndRow || 0);
  const bounds = getTablePhysicalBounds_(table);
  const endRow = requestedEndRow > 0 ? requestedEndRow : bounds.endRow;
  return Math.max(0, endRow - bounds.dataStartRow + 1);
}

function resolveTableSnapshotEndRow_(table, requestedEndRow) {
  const bounds = getTablePhysicalBounds_(table);
  const requested = Math.round(Number(requestedEndRow || 0));
  if (!requested) return bounds.endRow;
  return Math.max(bounds.dataStartRow - 1, Math.min(requested, bounds.maxRow));
}

function getTablePhysicalBounds_(table, includeColumns) {
  if (isNativeReadableTable_(table)) {
    const sheet = getNativeSheet_(table);
    const range = table.apiRange || {};
    const headers = includeColumns
      ? getNativeHeaderList_(table, readNativeTableHeaderRow_(table) || [])
      : [];
    const dataStartRow = Number(range.startRowIndex || 0) + 2;
    const tableEndRow = Number(range.endRowIndex || 0);
    return {
      sheet,
      headers,
      dataStartRow,
      startColumn: Number(range.startColumnIndex || 0) + 1,
      width: headers.length,
      endRow: Math.max(tableEndRow, Number(sheet.getLastRow() || 0), dataStartRow - 1),
      maxRow: Math.max(Number(sheet.getMaxRows() || 0), dataStartRow - 1),
    };
  }
  const sheet = getSheet_(table);
  const layout = includeColumns
    ? getSystemTableLayout_(sheet, table)
    : {
        headers: [],
        dataStartRow: getTableDataStartRow_(table),
        dataStartColumn: DATA_START_COLUMN,
      };
  return {
    sheet,
    headers: layout.headers || [],
    dataStartRow: Number(layout.dataStartRow || DATA_START_ROW),
    startColumn: Number(layout.dataStartColumn || DATA_START_COLUMN),
    width: (layout.headers || []).length,
    endRow: Math.max(Number(sheet.getLastRow() || 0), Number(layout.dataStartRow || DATA_START_ROW) - 1),
    maxRow: Math.max(Number(sheet.getMaxRows() || 0), Number(layout.dataStartRow || DATA_START_ROW) - 1),
  };
}

function getTableMetaUiConfig_(tableKey, forceRefresh) {
  try {
    const config = getUserAppConfig_(!!forceRefresh, { deferDataConfig: true });
    const pages = config.pages || {};
    const views = config.views || {};
    return {
      page: (pages.byTableKey || {})[tableKey] || null,
      views: (views.byTableKey || {})[tableKey] || [],
    };
  } catch (error) {
    return { page: null, views: [] };
  }
}

function getTableMetaKeyFields_(headers, ui) {
  const seen = {};
  const add = (value, acc) => {
    const field = String(value || '').trim();
    const normalized = normalizeConfigHeader_(field);
    if (!field || !normalized || seen[normalized]) return;
    if (headers && headers.length && !headers.some((header) => normalizeConfigHeader_(header) === normalized)) return;
    seen[normalized] = true;
    acc.push(field);
  };
  const result = [];
  const page = ui && ui.page ? ui.page : {};
  add(page.titleColumn || page.displayNameColumn || page.displayName || page.displayField || page.titleField, result);
  (ui && ui.views || []).forEach((view) => {
    add(view.displayNameColumn || view.displayName || view.displayField || view.titleField, result);
    (view.defaultGroupBy || []).forEach((field) => add(field, result));
    (view.filterFields || []).forEach((field) => add(field, result));
  });
  if (!result.length && headers && headers.length) result.push(headers[0]);
  return result;
}

function compactTableForMeta_(tableKey, table) {
  return {
    key: tableKey,
    name: table.name || '',
    sheetName: table.sheetName || '',
    gid: table.gid || '',
    type: table.type || '',
    binding: table.binding || '',
    spreadsheetId: table.spreadsheetId || SPREADSHEET_ID,
    spreadsheetUrl: table.spreadsheetUrl || table.sourceSpreadsheetUrl || '',
    inherited: !!table.inherited,
    readOnly: !!table.readOnly,
  };
}

function normalizeTablePageRequest_(request) {
  const pageSize = Math.max(50, Math.min(1000, Math.round(Number(request.pageSize || 100))));
  const filters = request.filters && typeof request.filters === 'object' ? request.filters : {};
  return {
    page: Math.max(1, Math.round(Number(request.page || 1))),
    pageSize,
    cursor: Math.max(0, Math.round(Number(request.cursor || 0))),
    snapshotEndRow: Math.max(0, Math.round(Number(request.snapshotEndRow || 0))),
    query: String(request.query || '').trim(),
    filters: Object.keys(filters).reduce((acc, header) => {
      const values = Array.isArray(filters[header]) ? filters[header] : [];
      const cleaned = values.map((value) => String(value == null ? '' : value)).filter((value) => value !== '');
      if (cleaned.length) acc[header] = cleaned;
      return acc;
    }, {}),
    sort: {
      header: String(request.sort && request.sort.header || '').trim(),
      direction: Number(request.sort && request.sort.direction) < 0 ? -1 : 1,
    },
    status: String(request.status || '').trim(),
    filterHeaders: Array.isArray(request.filterHeaders) ? request.filterHeaders.map((header) => String(header || '').trim()).filter(Boolean) : [],
  };
}

function hasServerTableProcessing_(payload) {
  if (payload.query || payload.status || (payload.sort && payload.sort.header) || (payload.filterHeaders || []).length) return true;
  return Object.keys(payload.filters || {}).some((header) => (payload.filters[header] || []).length);
}

function applyServerTableQuery_(rows, headers, payload) {
  const query = normalizeServerSearchText_(payload.query);
  let result = (rows || []).filter((row) => {
    const record = row.record || {};
    const statusHeader = payload.status ? (headers || []).find((header) => /status|stage/i.test(header)) : '';
    const matchesQuery = !query || normalizeServerSearchText_((row.values || []).join(' ')).indexOf(query) >= 0;
    const matchesStatus = !payload.status || (statusHeader && String(record[statusHeader] || '') === payload.status);
    const matchesFilters = Object.keys(payload.filters || {}).every((header) => {
      const selected = payload.filters[header] || [];
      if (!selected.length) return true;
      return selected.includes(String(record[header] == null ? '' : record[header]));
    });
    return matchesQuery && matchesStatus && matchesFilters;
  });

  if (payload.sort && payload.sort.header) {
    const header = resolveServerHeader_(payload.sort.header, headers);
    if (header) {
      result = result.slice().sort((left, right) => String((left.record || {})[header] || '')
        .localeCompare(String((right.record || {})[header] || ''), 'vi', { numeric: true }) * payload.sort.direction);
    }
  }
  return { rows: result };
}

function normalizeServerSearchText_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function resolveServerHeader_(header, headers) {
  const normalized = normalizeConfigHeader_(header);
  return (headers || []).find((candidate) => normalizeConfigHeader_(candidate) === normalized) || '';
}

function createServerFilterFacets_(rows, headers, requestedHeaders) {
  const selectedHeaders = (requestedHeaders && requestedHeaders.length ? requestedHeaders : headers || [])
    .map((header) => resolveServerHeader_(header, headers) || header)
    .filter(Boolean)
    .slice(0, 12);
  return selectedHeaders.reduce((acc, header) => {
    const counts = {};
    (rows || []).forEach((row) => {
      const value = String((row.record || {})[header] || '').trim();
      if (!value) return;
      counts[value] = (counts[value] || 0) + 1;
    });
    const options = Object.keys(counts)
      .sort((left, right) => counts[right] - counts[left] || left.localeCompare(right, 'vi', { numeric: true }))
      .slice(0, 40)
      .map((value) => ({ value, count: counts[value] }));
    if (options.length) acc[header] = options;
    return acc;
  }, {});
}

function buildTableData_(table, forceRefresh, options) {
  const data = buildTableRowsData_(table, options);
  return attachFieldSchema_(data, table, forceRefresh);
}

function buildTableRowsData_(table, options) {
  if (isNativeReadableTable_(table)) {
    return buildNativeTableRowsData_(table, options);
  }
  return buildSystemTableRowsData_(table, options);
}

function buildNativeTableData_(table, forceRefresh) {
  return attachFieldSchema_(buildNativeTableRowsData_(table), table, forceRefresh);
}

function buildNativeTableRowsData_(table, options) {
  const headerRow = readNativeTableHeaderRow_(table);
  const headers = getNativeHeaderList_(table, headerRow || []);
  if (!headers.length) {
    return { table, headers: [], rows: [], summary: createSummary_([], []), dataWindow: createDataWindowMeta_(0, false, 0) };
  }

  const latest = readLatestNativeDataRows_(table, headers, Number(options && options.limit || TABLE_ROW_LIMIT), options);
  const dataRows = latest.rows;

  return {
    table,
    headers,
    rows: dataRows,
    summary: createSummary_(headers, dataRows),
    dataWindow: createDataWindowMeta_(dataRows.length, latest.truncated, latest.offset, latest.mode),
  };
}

function buildSystemTableData_(table, forceRefresh) {
  return attachFieldSchema_(buildSystemTableRowsData_(table), table, forceRefresh);
}

function buildSystemTableRowsData_(table, options) {
  const sheet = getSheet_(table);
  const layout = getSystemTableLayout_(sheet, table);
  const headers = layout.headers;
  const width = headers.length;
  if (!width) {
    return { table, headers: [], rows: [], summary: createSummary_([], []), dataWindow: createDataWindowMeta_(0, false, 0) };
  }

  const latest = readLatestSystemDataRows_(sheet, layout, headers, Number(options && options.limit || TABLE_ROW_LIMIT), options);
  const rows = latest.rows;

  return {
    table,
    headers,
    rows,
    summary: createSummary_(headers, rows),
    dataWindow: createDataWindowMeta_(rows.length, latest.truncated, latest.offset, latest.mode),
  };
}

function saveRecord(payload) {
  assertSpreadsheetAccess_();
  validatePayload_(payload);
  const table = getWritableTableForPayload_(payload.tableKey);
  assertTableWritable_(table);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const rowNumber = isNativeReadableTable_(table)
      ? saveNativeRecord_(table, payload)
      : saveSystemRecord_(table, payload);
    SpreadsheetApp.flush();
    let savedRow = null;
    try {
      savedRow = readSavedRecordRow_(table, rowNumber);
    } catch (error) {
      console.warn('Saved row read-back skipped:', error && error.message ? error.message : error);
    }
    const invalidation = invalidateTableDataCache_(payload.tableKey) || {};

    return {
      ok: true,
      rowNumber,
      id: savedRow ? savedRow.id : '',
      recordId: savedRow ? savedRow.recordId : '',
      record: savedRow ? savedRow.record : payload.record,
      row: savedRow,
      tableVersion: invalidation.tableVersion || '',
    };
  } finally {
    lock.releaseLock();
  }
}

function importRecords(payload) {
  assertSpreadsheetAccess_();
  validateImportPayload_(payload);
  const table = getWritableTableForPayload_(payload.tableKey);
  assertTableWritable_(table);
  const records = payload.records.slice(0, 500);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const imported = isNativeReadableTable_(table)
      ? importNativeRecords_(table, records)
      : importSystemRecords_(table, records);
    invalidateTableDataCache_(payload.tableKey);

    return {
      ok: true,
      imported,
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteRecord(payload) {
  assertSpreadsheetAccess_();
  if (!payload || !payload.tableKey || !payload.rowNumber) {
    throw new Error('Missing tableKey or rowNumber.');
  }

  const table = getWritableTableForRow_(payload.tableKey, payload.rowNumber);
  assertTableWritable_(table);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const rowNumber = isNativeReadableTable_(table)
      ? clearNativeRecord_(table, payload)
      : clearSystemRecord_(table, payload);
    invalidateTableDataCache_(payload.tableKey);

    return {
      ok: true,
      rowNumber,
    };
  } finally {
    lock.releaseLock();
  }
}

function saveNativeRecord_(table, payload) {
  if (!payload.rowNumber) {
    table = getTableOrThrow_(table.key, true);
    assertTableWritable_(table);
  }
  const headers = getNativeHeaders_(table);
  const computedByHeader = getFormulaColumnByHeader_(table, headers);
  const editableHeaders = headers.filter((header) => !computedByHeader[header]);
  const rowNumber = Number(payload.rowNumber) || null;
  if (!editableHeaders.length) return rowNumber || getNextNativeAppendRowNumber_(table);
  if (rowNumber) {
    updateNativeEditableCells_(table, headers, editableHeaders, rowNumber, payload);
    return rowNumber;
  }

  const emptyRowNumber = findFirstEmptyNativeEditableRow_(table, headers, editableHeaders);
  if (emptyRowNumber) {
    updateNativeEditableCells_(table, headers, editableHeaders, emptyRowNumber, payload);
    return emptyRowNumber;
  }

  const nextRowNumber = getNextNativeAppendRowNumber_(table);
  const appendValues = headers.map((header) => computedByHeader[header] ? {} : toCellData_(payloadValue_(payload, header)));
  sheetsApiBatchUpdate_(getTableSpreadsheetId_(table), [{
    appendCells: {
      sheetId: getNativeSheetId_(table),
      tableId: table.apiTableId,
      rows: [{ values: appendValues }],
      fields: 'userEnteredValue',
    },
  }]);
  invalidateNativeTableRegistryCache_(getTableSpreadsheetId_(table));
  return nextRowNumber;
}

function updateNativeEditableCells_(table, headers, editableHeaders, rowNumber, payload) {
  const editableIndexes = editableHeaders
    .map((header) => headers.indexOf(header))
    .filter((index) => index >= 0)
    .sort((a, b) => a - b);
  const requests = groupContiguousIndexes_(editableIndexes).map((group) => {
    const startIndex = group[0];
    return {
      updateCells: {
        range: nativeColumnBlockRange_(table, rowNumber, startIndex, group.length),
        rows: [{ values: group.map((columnIndex) => toCellData_(payloadValue_(payload, headers[columnIndex]))) }],
        fields: 'userEnteredValue',
      },
    };
  });
  if (requests.length) sheetsApiBatchUpdate_(getTableSpreadsheetId_(table), requests);
}

function updateNativeEditableRows_(table, editableIndexes, rowNumber, rows) {
  if (!rows || !rows.length) return;
  const requests = groupContiguousIndexes_(editableIndexes).map((group) => {
    const startIndex = group[0];
    return {
      updateCells: {
        range: nativeColumnBlockRowsRange_(table, rowNumber, startIndex, group.length, rows.length),
        rows: rows.map((row) => ({
          values: group.map((columnIndex) => row[columnIndex] || {}),
        })),
        fields: 'userEnteredValue',
      },
    };
  });
  if (requests.length) sheetsApiBatchUpdate_(getTableSpreadsheetId_(table), requests);
}

function findFirstEmptyNativeEditableRow_(table, headers, editableHeaders) {
  const range = table.apiRange || {};
  const startRow = Number(range.startRowIndex || 0) + 2;
  const endRow = Number(range.endRowIndex || 0);
  if (!startRow || !endRow || endRow < startRow) return 0;
  const rows = readNativeTableDisplayRows_(table);
  const editableIndexes = editableHeaders.map((header) => headers.indexOf(header)).filter((index) => index >= 0);
  const bodyRows = rows.slice(1, 1 + endRow - startRow + 1);
  const emptyIndex = bodyRows.findIndex((row) => {
    const values = normalizePlainRowValues_(row, headers.length);
    return editableIndexes.every((index) => isDisplayBlank_(values[index]));
  });
  return emptyIndex >= 0 ? startRow + emptyIndex : 0;
}

function findFirstEmptyNativeEditableBlockRow_(table, headers, editableHeaders, blockSize) {
  const range = table.apiRange || {};
  const startRow = Number(range.startRowIndex || 0) + 2;
  const endRow = Number(range.endRowIndex || 0);
  const size = Math.max(1, Number(blockSize || 1));
  if (!startRow || !endRow || endRow < startRow || endRow - startRow + 1 < size) return 0;

  const rows = readNativeTableDisplayRows_(table);
  const editableIndexes = editableHeaders.map((header) => headers.indexOf(header)).filter((index) => index >= 0);
  const bodyRows = rows.slice(1, 1 + endRow - startRow + 1);
  for (let rowIndex = 0; rowIndex <= bodyRows.length - size; rowIndex += 1) {
    const hasEmptyBlock = bodyRows
      .slice(rowIndex, rowIndex + size)
      .every((row) => {
        const values = normalizePlainRowValues_(row, headers.length);
        return editableIndexes.every((index) => isDisplayBlank_(values[index]));
      });
    if (hasEmptyBlock) return startRow + rowIndex;
  }
  return 0;
}

function getNextNativeAppendRowNumber_(table) {
  const range = table.apiRange || {};
  if (range.endRowIndex != null) return Number(range.endRowIndex) + 1;
  return getNativeSheet_(table).getLastRow() + 1;
}

function importNativeRecords_(table, records) {
  table = getTableOrThrow_(table.key, true);
  assertTableWritable_(table);
  const headers = getNativeHeaders_(table);
  const computedByHeader = getFormulaColumnByHeader_(table, headers);
  const rows = records
    .map((record) => headers.map((header) => computedByHeader[header] ? {} : toCellData_(payloadValue_({ record }, header))))
    .filter((values) => values.some((value) => value.userEnteredValue && Object.keys(value.userEnteredValue).length));
  if (!rows.length) return 0;

  const editableHeaders = headers.filter((header) => !computedByHeader[header]);
  const editableIndexes = editableHeaders.map((header) => headers.indexOf(header)).filter((index) => index >= 0);
  const emptyRowNumber = findFirstEmptyNativeEditableBlockRow_(table, headers, editableHeaders, rows.length);
  if (emptyRowNumber) {
    updateNativeEditableRows_(table, editableIndexes, emptyRowNumber, rows);
    return rows.length;
  }

  sheetsApiBatchUpdate_(getTableSpreadsheetId_(table), [{
    appendCells: {
      sheetId: getNativeSheetId_(table),
      tableId: table.apiTableId,
      rows: rows.map((values) => ({ values })),
      fields: 'userEnteredValue',
    },
  }]);
  invalidateNativeTableRegistryCache_(getTableSpreadsheetId_(table));
  return rows.length;
}

function saveSystemRecord_(table, payload) {
  const sheet = getSheet_(table);
  const layout = getSystemTableLayout_(sheet, table);
  const headers = layout.headers;
  const computedByHeader = getFormulaColumnByHeader_(table, headers);
  const editableIndexes = getEditableHeaderIndexes_(headers, computedByHeader);
  const rowNumber = Number(payload.rowNumber) || getNextSystemAppendRowNumber_(sheet, layout, headers.length, editableIndexes);
  setSystemEditableRowValues_(sheet, layout, headers, editableIndexes, rowNumber, payload);
  SpreadsheetApp.flush();
  return rowNumber;
}

function importSystemRecords_(table, records) {
  const sheet = getSheet_(table);
  const layout = getSystemTableLayout_(sheet, table);
  const headers = layout.headers;
  const computedByHeader = getFormulaColumnByHeader_(table, headers);
  const editableHeaders = headers.filter((header) => !computedByHeader[header]);
  const rows = records
    .map((record) => editableHeaders.map((header) => payloadValue_({ record }, header)))
    .filter((values) => values.some((value) => value !== ''));
  if (!rows.length) return 0;

  const editableIndexes = editableHeaders.map((header) => headers.indexOf(header)).filter((index) => index >= 0);
  const rowNumber = getNextSystemAppendRowNumber_(sheet, layout, headers.length, editableIndexes, rows.length);
  setSystemEditableRowsValues_(sheet, layout, editableIndexes, rows, rowNumber);
  SpreadsheetApp.flush();
  return rows.length;
}

function clearNativeRecord_(table, payload) {
  const headers = getNativeHeaders_(table);
  const rowNumber = resolveNativeDeleteRowNumber_(table, payload, headers);
  sheetsApiBatchUpdate_(getTableSpreadsheetId_(table), [{
    repeatCell: {
      range: nativeRowRange_(table, rowNumber, headers.length),
      cell: {},
      fields: 'userEnteredValue',
    },
  }]);
  return rowNumber;
}

function resolveNativeDeleteRowNumber_(table, payload, headers) {
  const requestedRowNumber = Number(payload.rowNumber);
  if (isNativeRowNumberInRange_(table, requestedRowNumber)) return requestedRowNumber;

  const resolved = findNativeRowNumberByRecord_(table, payload.record || {}, headers);
  if (resolved) return resolved;
  throw new Error('Record row is no longer available. Reload and try again.');
}

function findNativeRowNumberByRecord_(table, record, headers) {
  const matchHeaders = headers.filter((header) => {
    const value = record && Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '';
    return String(value == null ? '' : value).trim() !== '';
  });
  if (!matchHeaders.length) return 0;

  const allRows = readNativeTableDisplayRows_(table);
  const rowNumberBase = Number(table.apiRange.startRowIndex || 0) + 2;
  const matches = allRows.slice(1).map((row, index) => {
    const values = normalizePlainRowValues_(row, headers.length);
    const candidate = toRecord_(headers, values);
    const matched = matchHeaders.every((header) => String(candidate[header] || '') === String(record[header] || ''));
    return matched ? rowNumberBase + index : 0;
  }).filter(Boolean);
  return matches.length === 1 ? matches[0] : 0;
}

function clearSystemRecord_(table, payload) {
  const sheet = getSheet_(table);
  const layout = getSystemTableLayout_(sheet, table);
  const rowNumber = Number(payload.rowNumber);
  if (rowNumber < layout.dataStartRow) {
    throw new Error('Invalid rowNumber.');
  }
  sheet.getRange(rowNumber, layout.dataStartColumn, 1, layout.headers.length).clearContent();
  SpreadsheetApp.flush();
  return rowNumber;
}

function getTableOrThrow_(tableKey, forceRefresh) {
  const table = getTables_(forceRefresh)[tableKey];
  if (!table) {
    throw new Error(`Unknown table key: ${tableKey}`);
  }
  return table;
}

function getWritableTableForPayload_(tableKey) {
  try {
    return getTableOrThrow_(tableKey, false);
  } catch (error) {
    return getTableOrThrow_(tableKey, true);
  }
}

function getWritableTableForRow_(tableKey, rowNumber) {
  let table = getWritableTableForPayload_(tableKey);
  if (isNativeReadableTable_(table) && !isNativeRowNumberInRange_(table, Number(rowNumber))) {
    table = getTableOrThrow_(tableKey, true);
  }
  return table;
}

function getSheet_(table) {
  const spreadsheetId = getTableSpreadsheetId_(table);
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    logAppDiagnostic_('error', 'spreadsheet_open_failed', {
      operation: 'SpreadsheetApp.openById',
      spreadsheetId,
      tableKey: table.key || '',
      tableName: table.name || table.sheetName || '',
      sheetName: table.sheetName || '',
      gid: table.gid || '',
      binding: table.binding || '',
    }, error);
    throw new Error(`Cannot open data spreadsheet for ${table.name || table.sheetName || 'table'} (${spreadsheetId}). ${error && error.message ? error.message : String(error)}`);
  }
  let sheet = table.sheetName ? spreadsheet.getSheetByName(table.sheetName) : null;
  if (!sheet && table.gid) {
    sheet = spreadsheet.getSheets().find((candidate) => candidate.getSheetId() === Number(table.gid));
  }
  if (!sheet) {
    throw new Error(`Sheet not found for ${table.name || 'table'} in spreadsheet ${spreadsheetId}: ${table.sheetName || table.gid}`);
  }
  return sheet;
}

function getTableSpreadsheetId_(table) {
  return table.spreadsheetId || SPREADSHEET_ID;
}

function isNativeReadableTable_(table) {
  return table && (table.binding === 'nativeTable' || table.binding === 'inheritedNativeTable');
}

function assertTableWritable_(table) {
  if (table && table.readOnly) {
    throw new Error(`${table.name || 'This table'} is inherited/read-only and cannot be modified from this app.`);
  }
}

function getTableDataStartColumn_(table) {
  return table.dataStartColumn || DATA_START_COLUMN;
}

function getTableHeaderRow_(table) {
  return table.headerRow || HEADER_ROW;
}

function getTableDataStartRow_(table) {
  return table.dataStartRow || DATA_START_ROW;
}

function getHeaders_(sheet, table) {
  return getSystemTableLayout_(sheet, table).headers;
}

function getSystemTableLayout_(sheet, table) {
  const dataStartColumn = getTableDataStartColumn_(table);
  const headerRow = getTableHeaderRow_(table);
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < dataStartColumn) {
    return { headerRow, dataStartRow: getTableDataStartRow_(table), dataStartColumn, headers: [] };
  }

  const values = sheet
    .getRange(headerRow, dataStartColumn, 1, lastColumn - dataStartColumn + 1)
    .getDisplayValues()[0];
  const lastHeaderIndex = values.reduce((last, value, index) => (value ? index : last), -1);
  const headers = lastHeaderIndex >= 0 ? values.slice(0, lastHeaderIndex + 1) : [];
  return { headerRow, dataStartRow: getTableDataStartRow_(table), dataStartColumn, headers };
}

function findFirstEmptyDataRow_(sheet, layout, width) {
  return findFirstEmptyDataRowByColumns_(sheet, layout, width, Array.from({ length: width }, (_, index) => index));
}

function findFirstEmptyDataRowByColumns_(sheet, layout, width, columnIndexes) {
  return findFirstEmptyDataBlockRowByColumns_(sheet, layout, width, columnIndexes, 1);
}

function findFirstEmptyDataBlockRowByColumns_(sheet, layout, width, columnIndexes, blockSize) {
  const dataStartRow = layout.dataStartRow;
  const dataStartColumn = layout.dataStartColumn;
  const size = Math.max(1, Number(blockSize || 1));
  const lastRow = Math.max(sheet.getLastRow(), dataStartRow + size - 1);
  const rows = sheet.getRange(dataStartRow, dataStartColumn, lastRow - dataStartRow + 1, width).getDisplayValues();
  const indexes = columnIndexes && columnIndexes.length ? columnIndexes : Array.from({ length: width }, (_, index) => index);
  for (let rowIndex = 0; rowIndex <= rows.length - size; rowIndex += 1) {
    const hasEmptyBlock = rows
      .slice(rowIndex, rowIndex + size)
      .every((row) => indexes.every((index) => isDisplayBlank_(row[index])));
    if (hasEmptyBlock) return dataStartRow + rowIndex;
  }
  return lastRow + 1;
}

function getNextSystemAppendRowNumber_(sheet, layout, width, editableIndexes, blockSize) {
  if (width && editableIndexes && editableIndexes.length) {
    return findFirstEmptyDataBlockRowByColumns_(sheet, layout, width, editableIndexes, blockSize || 1);
  }
  return Math.max(Number(sheet.getLastRow() || 0) + 1, Number(layout.dataStartRow || DATA_START_ROW));
}

function getEditableHeaderIndexes_(headers, computedByHeader) {
  return (headers || [])
    .map((header, index) => computedByHeader && computedByHeader[header] ? -1 : index)
    .filter((index) => index >= 0);
}

function setSystemEditableRowValues_(sheet, layout, headers, editableIndexes, rowNumber, payload) {
  groupContiguousIndexes_(editableIndexes).forEach((group) => {
    const startIndex = group[0];
    const values = group.map((columnIndex) => payloadValue_(payload, headers[columnIndex]));
    sheet
      .getRange(rowNumber, layout.dataStartColumn + startIndex, 1, group.length)
      .setValues([values]);
  });
}

function setSystemEditableRowsValues_(sheet, layout, editableIndexes, rows, rowNumber) {
  groupContiguousIndexes_(editableIndexes).forEach((group) => {
    const startIndex = group[0];
    const values = rows.map((row) => group.map((columnIndex) => row[editableIndexes.indexOf(columnIndex)]));
    sheet
      .getRange(rowNumber, layout.dataStartColumn + startIndex, rows.length, group.length)
      .setValues(values);
  });
}

function groupContiguousIndexes_(indexes) {
  const groups = [];
  (indexes || []).forEach((index) => {
    const value = Number(index);
    if (!Number.isFinite(value) || value < 0) return;
    const last = groups[groups.length - 1];
    if (last && last[last.length - 1] === value - 1) last.push(value);
    else groups.push([value]);
  });
  return groups;
}

function getNativeHeaders_(table) {
  return getNativeHeaderList_(table, []);
}

function createDataWindowMeta_(returnedRows, truncated, offset, mode) {
  const startOffset = Math.max(0, Number(offset || 0));
  const count = Number(returnedRows || 0);
  return {
    mode: mode || 'latest',
    rowLimit: TABLE_ROW_LIMIT,
    returnedRows: count,
    offset: startOffset,
    nextOffset: startOffset + count,
    truncated: !!truncated,
  };
}

function createDataRowPayload_(headers, rowNumber, values) {
  const normalized = normalizePlainRowValues_(values, headers.length);
  return {
    rowNumber,
    id: normalized[0] || `row-${rowNumber}`,
    recordId: findRecordIdValue_(headers, normalized),
    values: normalized,
    record: toRecord_(headers, normalized),
  };
}

function findRecordIdValue_(headers, values) {
  const candidates = (headers || []).map((header, index) => ({
    index,
    name: String(header || '').trim().toLowerCase().replace(/[\s-]+/g, '_'),
  }));
  const idColumn = candidates.find((candidate) => candidate.name === 'id') ||
    candidates.find((candidate) => /_id$/.test(candidate.name));
  return idColumn && values[idColumn.index] != null ? String(values[idColumn.index]) : '';
}

function readSavedRecordRow_(table, rowNumber) {
  const targetRow = Number(rowNumber || 0);
  if (!targetRow) return null;

  let sheet;
  let headers;
  let startColumn;
  if (isNativeReadableTable_(table)) {
    sheet = getNativeSheet_(table);
    headers = getNativeHeaderList_(table, readNativeTableHeaderRow_(table));
    startColumn = Number(table.apiRange && table.apiRange.startColumnIndex || 0) + 1;
  } else {
    sheet = getSheet_(table);
    const layout = getSystemTableLayout_(sheet, table);
    headers = layout.headers;
    startColumn = layout.dataStartColumn;
  }

  if (!headers || !headers.length) return null;
  const values = sheet
    .getRange(targetRow, startColumn, 1, headers.length)
    .getDisplayValues()[0] || [];
  return createDataRowPayload_(headers, targetRow, values);
}

function isNonBlankDisplayRow_(values) {
  return (values || []).some((value) => !isDisplayBlank_(value));
}

function readLatestSystemDataRows_(sheet, layout, headers, limit, options) {
  const dataStartRow = Number(layout.dataStartRow || DATA_START_ROW);
  const dataStartColumn = Number(layout.dataStartColumn || DATA_START_COLUMN);
  const width = headers.length;
  const requestedEndRow = Math.round(Number(options && options.endRow || 0));
  const lastRow = requestedEndRow > 0
    ? Math.min(requestedEndRow, Number(sheet.getMaxRows() || requestedEndRow))
    : Number(sheet.getLastRow() || 0);
  const maxRows = Math.max(0, lastRow - dataStartRow + 1);
  if (!maxRows) return { rows: [], truncated: false };

  const rangeOptions = {
    startRow: dataStartRow,
    startColumn: dataStartColumn,
    width,
    rowCount: maxRows,
    limit,
    offset: options && options.offset,
  };
  const capped = options && options.fromStart
    ? readFirstRowsFromRange_(sheet, rangeOptions)
    : readLatestRowsFromRange_(sheet, rangeOptions);
  return {
    rows: capped.rows.map((entry) => createDataRowPayload_(headers, entry.rowNumber, entry.values)),
    truncated: capped.truncated,
    offset: capped.offset,
  };
}

function readLatestNativeDataRows_(table, headers, limit, options) {
  const sheet = getNativeSheet_(table);
  const range = table.apiRange || {};
  const width = headers.length;
  const dataStartRow = Number(range.startRowIndex || 0) + 2;
  const startColumn = Number(range.startColumnIndex || 0) + 1;
  const tableEndRow = Number(range.endRowIndex || 0);
  const sheetLastRow = Number(sheet.getLastRow() || 0);
  const requestedEndRow = Math.round(Number(options && options.endRow || 0));
  const currentEndRow = Math.max(tableEndRow, sheetLastRow, dataStartRow - 1);
  const lastRow = requestedEndRow > 0
    ? Math.max(dataStartRow - 1, Math.min(requestedEndRow, Number(sheet.getMaxRows() || requestedEndRow)))
    : currentEndRow;
  const rowCount = Math.max(0, lastRow - dataStartRow + 1);
  if (!rowCount || !width) return { rows: [], truncated: false };

  const rangeOptions = {
    startRow: dataStartRow,
    startColumn,
    width,
    rowCount,
    limit,
    offset: options && options.offset,
  };
  const capped = options && options.fromStart
    ? readFirstRowsFromRange_(sheet, rangeOptions)
    : readLatestRowsFromRange_(sheet, rangeOptions);
  return {
    rows: capped.rows.map((entry) => createDataRowPayload_(headers, entry.rowNumber, entry.values)),
    truncated: capped.truncated,
    offset: capped.offset,
  };
}

function readLatestRowsFromRange_(sheet, options) {
  const startRow = Number(options.startRow || 1);
  const startColumn = Number(options.startColumn || 1);
  const width = Number(options.width || 0);
  const rowCount = Number(options.rowCount || 0);
  const limit = Math.max(1, Number(options.limit || TABLE_ROW_LIMIT));
  const offset = Math.max(0, Number(options.offset || 0));
  if (!width || !rowCount) return { rows: [], truncated: false };

  const latest = [];
  let skipped = 0;
  let cursor = startRow + rowCount - 1;
  while (cursor >= startRow && latest.length <= limit) {
    const remainingNeeded = Math.max(1, limit + offset + 1 - latest.length - skipped);
    const chunkHeight = Math.min(TABLE_READ_CHUNK_SIZE, remainingNeeded, cursor - startRow + 1);
    const chunkStart = cursor - chunkHeight + 1;
    const values = sheet.getRange(chunkStart, startColumn, chunkHeight, width).getDisplayValues();
    for (let index = values.length - 1; index >= 0 && latest.length <= limit; index -= 1) {
      const row = values[index] || [];
      if (!isNonBlankDisplayRow_(row)) continue;
      if (skipped < offset) {
        skipped += 1;
        continue;
      }
      latest.push({ rowNumber: chunkStart + index, values: row });
    }
    cursor = chunkStart - 1;
  }

  const truncated = latest.length > limit;
  return {
    rows: latest.slice(0, limit).reverse(),
    truncated,
    offset,
  };
}

function readFirstRowsFromRange_(sheet, options) {
  const startRow = Number(options.startRow || 1);
  const startColumn = Number(options.startColumn || 1);
  const width = Number(options.width || 0);
  const rowCount = Number(options.rowCount || 0);
  const limit = Math.max(1, Number(options.limit || TABLE_INITIAL_ROW_LIMIT));
  const offset = Math.max(0, Number(options.offset || 0));
  if (!width || !rowCount) return { rows: [], truncated: false, offset, mode: 'first' };

  const rows = [];
  let skipped = 0;
  const endRow = startRow + rowCount - 1;
  let cursor = startRow;
  while (cursor <= endRow && rows.length <= limit) {
    const chunkHeight = Math.min(TABLE_READ_CHUNK_SIZE, endRow - cursor + 1);
    const values = sheet.getRange(cursor, startColumn, chunkHeight, width).getDisplayValues();
    for (let index = 0; index < values.length && rows.length <= limit; index += 1) {
      const row = values[index] || [];
      if (!isNonBlankDisplayRow_(row)) continue;
      if (skipped < offset) {
        skipped += 1;
        continue;
      }
      rows.push({ rowNumber: cursor + index, values: row });
    }
    cursor += chunkHeight;
  }

  const truncated = rows.length > limit;
  return {
    rows: rows.slice(0, limit),
    truncated,
    offset,
    mode: 'first',
  };
}

function getNativeHeadersFromRowData_(table, headerRowData) {
  const width = getNativeTableWidth_(table);
  const physicalHeaders = width && headerRowData
    ? normalizeRowValues_(headerRowData.values || [], width)
    : [];
  return getNativeHeaderList_(table, physicalHeaders);
}

function getNativeHeaderList_(table, physicalHeaders) {
  const width = getNativeTableWidth_(table);
  const tableColumns = table.columns || [];
  const headers = Array.from({ length: width || tableColumns.length }, (_, index) => (
    (physicalHeaders || [])[index] || tableColumns[index] || ''
  ));
  const lastHeaderIndex = headers.reduce((last, value, index) => (value ? index : last), -1);
  if (lastHeaderIndex >= 0) return headers.slice(0, lastHeaderIndex + 1);
  return tableColumns;
}

function readNativeTableHeaderRow_(table) {
  return (readNativeTableDisplayRows_(table, 1)[0] || []);
}

function readNativeTableDisplayRows_(table, maxRows) {
  const sheet = getNativeSheet_(table);
  const range = table.apiRange || {};
  const startRowIndex = Number(range.startRowIndex || 0);
  const startColumnIndex = Number(range.startColumnIndex || 0);
  const width = getNativeTableWidth_(table);
  if (!width) return [];

  const endRowIndex = maxRows
    ? startRowIndex + Number(maxRows)
    : Math.max(Number(range.endRowIndex || 0), Number(sheet.getLastRow() || 0));
  const height = Math.max(0, endRowIndex - startRowIndex);
  if (!height) return [];

  return sheet
    .getRange(startRowIndex + 1, startColumnIndex + 1, height, width)
    .getDisplayValues();
}

function getNativeSheet_(table) {
  const spreadsheetId = getTableSpreadsheetId_(table);
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    throw new Error(`Cannot open data spreadsheet for ${table.name || table.apiTableName || 'native table'} (${spreadsheetId}). ${error && error.message ? error.message : String(error)}`);
  }
  const sheetId = getNativeSheetId_(table);
  const sheet = spreadsheet.getSheets().find((candidate) => candidate.getSheetId() === sheetId);
  if (!sheet) throw new Error(`Sheet not found for native table ${table.name || table.apiTableName || sheetId} in spreadsheet ${spreadsheetId}`);
  return sheet;
}

function getNativeTableWidth_(table) {
  const range = table.apiRange || {};
  return Number(range.endColumnIndex || 0) - Number(range.startColumnIndex || 0);
}

function attachFieldSchema_(data, table, forceRefresh) {
  const headers = data.headers || [];
  const schema = getTableFieldSchema_(table, headers, forceRefresh);
  const rows = table && table.type === 'config'
    ? attachConfigRowValidation_(data.rows || [], table, headers)
    : data.rows;
  return Object.assign({}, data, {
    fields: schema.fields,
    rows,
    validationByHeader: schema.validationByHeader,
    formulaAware: !!schema.formulaAware,
    formulaHeaders: schema.formulaHeaders || [],
    schemaGeneratedAt: schema.generatedAt,
  });
}

function attachLightFieldSchema_(data) {
  const headers = data && data.headers ? data.headers : [];
  const fields = headers.map((header, index) => ({
    name: header,
    label: header,
    index,
    inputType: inferFieldInputType_(header, null),
    validation: null,
    computed: false,
    editable: true,
    formulaDriven: false,
  }));
  return Object.assign({}, data || {}, {
    fields,
    validationByHeader: {},
    formulaAware: false,
    formulaHeaders: [],
    schemaGeneratedAt: new Date().toISOString(),
    schemaMode: 'light',
  });
}

function getTableFieldSchema_(table, headers, forceRefresh) {
  const tableKey = table.key || table.apiTableId || table.name || table.sheetName || 'unknown';
  const cacheKey = tableScopedCacheKey_('validation', tableKey, table);
  const headerSignature = `${getHeaderSignature_(headers)}::formula-aware-v3`;
  if (!forceRefresh) {
    const cached = getCached_(cacheKey);
    if (cached && cached.headerSignature === headerSignature) return cached;
  }

  const deepScan = shouldScanDeepFieldSchema_(table, forceRefresh);
  const cellValidationByHeader = deepScan && table && table.type !== 'config' ? readCellValidationByHeader_(table, headers) : {};
  const nativeValidationByHeader = readNativeTableColumnValidationByHeader_(table, headers);
  const formulaByHeader = deepScan ? getFormulaColumnByHeader_(table, headers, forceRefresh) : {};
  const formulaHeaders = Object.keys(formulaByHeader);
  const validationByHeader = Object.assign({}, cellValidationByHeader, nativeValidationByHeader);
  let fields = headers.map((header, index) => {
    const validation = validationByHeader[header] || null;
    const computed = !!formulaByHeader[header];
    return {
      name: header,
      label: header,
      index,
      inputType: inferFieldInputType_(header, validation),
      validation,
      computed,
      editable: !computed,
      formulaDriven: computed,
    };
  });
  fields = applyUserFieldConfig_(table, fields);
  const schema = {
    fields,
    validationByHeader,
    formulaAware: !!deepScan,
    formulaHeaders,
    headerSignature,
    generatedAt: new Date().toISOString(),
  };
  setCached_(cacheKey, schema);
  return schema;
}

function shouldScanDeepFieldSchema_(table, forceRefresh) {
  if (forceRefresh) return true;
  return table && table.type === 'config';
}

function readFormulaColumnByHeader_(table, headers) {
  if (!headers || !headers.length) return {};
  try {
    const sheet = getSheet_(table);
    const layout = isNativeReadableTable_(table)
      ? getNativeValidationLayout_(table, headers)
      : getSystemTableLayout_(sheet, table);
    const scanRows = getValidationScanRowCount_(sheet, layout.dataStartRow);
    if (!scanRows) return {};
    const formulas = sheet
      .getRange(layout.dataStartRow, layout.dataStartColumn, scanRows, headers.length)
      .getFormulas();
    return headers.reduce((acc, header, columnIndex) => {
      const hasFormula = formulas.some((row) => String((row || [])[columnIndex] || '').trim() !== '');
      if (hasFormula) acc[header] = true;
      return acc;
    }, {});
  } catch (error) {
    return {};
  }
}

function getFormulaColumnByHeader_(table, headers, forceRefresh) {
  if (!headers || !headers.length) return {};
  const tableKey = table.key || table.apiTableId || table.name || table.sheetName || 'unknown';
  const cacheKey = tableScopedCacheKey_('formula', tableKey, table);
  const headerSignature = `${getHeaderSignature_(headers)}::formula-mask-v1`;
  if (!forceRefresh) {
    const cached = getCached_(cacheKey);
    if (cached && cached.headerSignature === headerSignature) return cached.formulaByHeader || {};
  }
  const formulaByHeader = readFormulaColumnByHeader_(table, headers);
  setCached_(cacheKey, { headerSignature, formulaByHeader });
  return formulaByHeader;
}

function attachConfigRowValidation_(rows, table, headers) {
  if (!rows || !rows.length || !headers || !headers.length) return rows;
  try {
    const sheet = getSheet_(table);
    const layout = isNativeReadableTable_(table)
      ? getNativeValidationLayout_(table, headers)
      : getSystemTableLayout_(sheet, table);
    const scanRows = getValidationScanRowCount_(sheet, layout.dataStartRow);
    if (!scanRows) return rows;

    const rules = sheet
      .getRange(layout.dataStartRow, layout.dataStartColumn, scanRows, headers.length)
      .getDataValidations();
    return rows.map((row) => {
      const rowIndex = Number(row.rowNumber || 0) - layout.dataStartRow;
      const rowRules = rowIndex >= 0 ? rules[rowIndex] || [] : [];
      const validations = headers.reduce((acc, header, columnIndex) => {
        const rule = rowRules[columnIndex] || null;
        const validation = rule ? spreadsheetDataValidationToFieldValidation_(rule) : null;
        if (validation) acc[header] = validation;
        return acc;
      }, {});
      return Object.keys(validations).length ? Object.assign({}, row, { validations }) : row;
    });
  } catch (error) {
    return rows;
  }
}

function getHeaderSignature_(headers) {
  return (headers || []).map((header) => String(header || '').trim()).join('\u001f');
}

function readNativeTableColumnValidationByHeader_(table, headers) {
  if (!isNativeReadableTable_(table) || !table.columnProperties) return {};
  const byHeader = {};
  table.columnProperties.forEach((column, index) => {
    const relativeIndex = getNativeTableColumnRelativeIndex_(column, index);
    const header = headers[relativeIndex] || column.columnName || '';
    if (!header || !column.dataValidationRule) return;
    const validation = apiDataValidationRuleToFieldValidation_(column.dataValidationRule, table);
    if (!validation) return;

    validation.source = 'native_table_column';
    validation.columnType = column.columnType || '';
    validation.columnName = column.columnName || header;
    validation.columnIndex = relativeIndex;
    byHeader[header] = validation;

    if (column.columnName && column.columnName !== header) {
      byHeader[column.columnName] = validation;
    }
  });
  return byHeader;
}

function getNativeTableColumnRelativeIndex_(column, fallbackIndex) {
  const index = Number(column && column.columnIndex);
  return Number.isFinite(index) ? index : fallbackIndex;
}

function readCellValidationByHeader_(table, headers) {
  if (!headers || !headers.length) return {};
  try {
    const sheet = getSheet_(table);
    const layout = isNativeReadableTable_(table)
      ? getNativeValidationLayout_(table, headers)
      : getSystemTableLayout_(sheet, table);
    const scanRows = getValidationScanRowCount_(sheet, layout.dataStartRow);
    if (!scanRows) return {};

    const rules = sheet
      .getRange(layout.dataStartRow, layout.dataStartColumn, scanRows, headers.length)
      .getDataValidations();
    return headers.reduce((acc, header, columnIndex) => {
      const rule = findColumnValidationRule_(rules, columnIndex);
      const validation = rule ? spreadsheetDataValidationToFieldValidation_(rule) : null;
      if (validation) acc[header] = validation;
      return acc;
    }, {});
  } catch (error) {
    return {};
  }
}

function getNativeValidationLayout_(table, headers) {
  const range = table.apiRange || {};
  return {
    headerRow: Number(range.startRowIndex || 0) + 1,
    dataStartRow: Number(range.startRowIndex || 0) + 2,
    dataStartColumn: Number(range.startColumnIndex || 0) + 1,
    headers,
  };
}

function getValidationScanRowCount_(sheet, dataStartRow) {
  const maxRows = sheet.getMaxRows();
  if (dataStartRow > maxRows) return 0;
  return Math.min(VALIDATION_SCAN_ROWS, maxRows - dataStartRow + 1);
}

function findColumnValidationRule_(rules, columnIndex) {
  for (let rowIndex = 0; rowIndex < rules.length; rowIndex += 1) {
    const rule = (rules[rowIndex] || [])[columnIndex] || null;
    if (rule) return rule;
  }
  return null;
}

function spreadsheetDataValidationToFieldValidation_(rule) {
  const criteria = rule.getCriteriaType();
  const values = rule.getCriteriaValues() || [];
  const validation = {
    criteria: String(criteria || ''),
    strict: !rule.getAllowInvalid(),
    helpText: rule.getHelpText() || '',
    options: [],
    rangeA1: '',
    args: [],
    source: 'cell_validation',
  };

  if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    validation.options = normalizeValidationOptions_(values[0] || []);
    validation.showDropdown = values.length > 1 ? !!values[1] : true;
  } else if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    const range = values[0];
    validation.rangeA1 = range && typeof range.getA1Notation === 'function' ? range.getA1Notation() : '';
    validation.options = range && typeof range.getDisplayValues === 'function'
      ? normalizeValidationOptions_(range.getDisplayValues().flat())
      : [];
    validation.showDropdown = values.length > 1 ? !!values[1] : true;
  } else {
    validation.args = values.map((value) => validationArgToString_(value)).filter((value) => value !== '');
  }

  return validation;
}

function apiDataValidationRuleToFieldValidation_(rule, table) {
  const condition = rule.condition || {};
  const type = condition.type || '';
  const values = condition.values || [];
  const validation = {
    criteria: type,
    strict: true,
    helpText: rule.inputMessage || '',
    options: [],
    rangeA1: '',
    args: [],
    showDropdown: true,
    source: 'native_table_column',
  };

  if (type === 'ONE_OF_LIST') {
    validation.options = normalizeValidationOptions_(values.map((value) => value.userEnteredValue || ''));
  } else if (type === 'ONE_OF_RANGE') {
    validation.rangeA1 = values[0] && values[0].userEnteredValue ? values[0].userEnteredValue : '';
    validation.options = getValidationOptionsFromA1_(validation.rangeA1, getTableSpreadsheetId_(table || {}));
  } else {
    validation.args = values.map((value) => value.userEnteredValue || '').filter((value) => value !== '');
  }

  return validation;
}

function getValidationOptionsFromA1_(a1Notation, spreadsheetId) {
  const reference = normalizeValidationRangeReference_(a1Notation);
  if (!reference) return [];
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId || SPREADSHEET_ID);
    const structuredOptions = getValidationOptionsFromStructuredTableReference_(spreadsheet, reference);
    if (structuredOptions.length) return structuredOptions;

    const namedRange = spreadsheet.getRangeByName(reference);
    const range = namedRange || spreadsheet.getRange(reference);
    return normalizeValidationOptions_(range.getDisplayValues().flat());
  } catch (error) {
    return [];
  }
}

function normalizeValidationRangeReference_(reference) {
  return String(reference || '').trim().replace(/^=/, '').trim();
}

function getValidationOptionsFromStructuredTableReference_(spreadsheet, reference) {
  const match = String(reference || '').match(/^'?(.+?)'?\[([^\]]+)\]$/);
  if (!match) return [];

  const tableName = match[1].trim();
  const columnName = match[2].trim();
  const table = findNativeTableInSpreadsheet_(spreadsheet.getId(), tableName);
  if (!table || !table.range) return [];

  const columns = table.columnProperties || [];
  const column = columns.find((candidate, index) => {
    const candidateName = candidate.columnName || '';
    const candidateIndex = getNativeTableColumnRelativeIndex_(candidate, index);
    return normalizeName(candidateName) === normalizeName(columnName) ||
      normalizeName(candidateName) === normalizeName(columnName.replace(/_/g, ' ')) ||
      candidateIndex === Number(columnName);
  });
  if (!column) return [];

  const relativeIndex = getNativeTableColumnRelativeIndex_(column, columns.indexOf(column));
  const sheet = spreadsheet.getSheets().find((candidate) => (
    candidate.getSheetId() === Number(table.range.sheetId)
  ));
  if (!sheet) return [];

  const startRow = Number(table.range.startRowIndex || 0) + 2;
  const endRow = Number(table.range.endRowIndex || sheet.getMaxRows());
  const rowCount = Math.max(0, endRow - startRow + 1);
  if (!rowCount) return [];

  const columnNumber = Number(table.range.startColumnIndex || 0) + relativeIndex + 1;
  return normalizeValidationOptions_(sheet.getRange(startRow, columnNumber, rowCount, 1).getDisplayValues().flat());
}

function findNativeTableInSpreadsheet_(spreadsheetId, tableName) {
  const normalized = normalizeTableLookupName_(tableName);
  try {
    const tables = getNativeSpreadsheetTablesById_(spreadsheetId || SPREADSHEET_ID, false).tables || [];
    return tables.find((table) => normalizeTableLookupName_(table.name) === normalized) || null;
  } catch (error) {
    return null;
  }
}

function normalizeValidationOptions_(values) {
  const seen = {};
  return (Array.isArray(values) ? values : [values])
    .map((value) => String(value == null ? '' : value).trim())
    .filter((value) => value !== '')
    .filter((value) => {
      const key = normalizeName(value);
      if (seen[key]) return false;
      seen[key] = true;
      return true;
    });
}

function validationArgToString_(value) {
  if (value == null) return '';
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (typeof value.getA1Notation === 'function') return value.getA1Notation();
  return String(value);
}

function inferFieldInputType_(header, validation) {
  if (validation && validation.options && validation.options.length) {
    return validation.strict ? 'select' : 'datalist';
  }
  const criteria = validation ? String(validation.criteria || '').toLowerCase() : '';
  const normalized = String(header || '').toLowerCase();
  if (criteria.indexOf('date') >= 0 || /date|deadline|target/.test(normalized)) return 'date';
  if (criteria.indexOf('number') >= 0 || /amount|budget|price|cost|number|limit|value|quantity|qty/.test(normalized)) return 'number';
  return 'text';
}

function normalizeRowValues_(cells, width) {
  const values = [];
  for (let index = 0; index < width; index += 1) {
    values.push(cellDisplayValue_(cells[index] || {}));
  }
  return values;
}

function normalizePlainRowValues_(values, width) {
  const row = Array.isArray(values) ? values : [];
  return Array.from({ length: width }, (_, index) => String(row[index] == null ? '' : row[index]));
}

function isDisplayBlank_(value) {
  return String(value == null ? '' : value).trim() === '';
}

function cellDisplayValue_(cell) {
  if (Object.prototype.hasOwnProperty.call(cell, 'formattedValue')) return cell.formattedValue || '';
  const value = cell.effectiveValue || cell.userEnteredValue || {};
  if (Object.prototype.hasOwnProperty.call(value, 'stringValue')) return value.stringValue || '';
  if (Object.prototype.hasOwnProperty.call(value, 'numberValue')) return String(value.numberValue);
  if (Object.prototype.hasOwnProperty.call(value, 'boolValue')) return value.boolValue ? 'TRUE' : 'FALSE';
  return '';
}

function nativeRowRange_(table, rowNumber, width) {
  validateNativeRowNumber_(table, rowNumber);
  return {
    sheetId: getNativeSheetId_(table),
    startRowIndex: rowNumber - 1,
    endRowIndex: rowNumber,
    startColumnIndex: table.apiRange.startColumnIndex || 0,
    endColumnIndex: Number(table.apiRange.startColumnIndex || 0) + width,
  };
}

function nativeSingleCellRange_(table, rowNumber, relativeColumnIndex) {
  validateNativeRowNumber_(table, rowNumber);
  const startColumnIndex = Number(table.apiRange.startColumnIndex || 0) + Number(relativeColumnIndex || 0);
  return {
    sheetId: getNativeSheetId_(table),
    startRowIndex: rowNumber - 1,
    endRowIndex: rowNumber,
    startColumnIndex,
    endColumnIndex: startColumnIndex + 1,
  };
}

function nativeColumnBlockRange_(table, rowNumber, relativeColumnIndex, width) {
  validateNativeRowNumber_(table, rowNumber);
  const startColumnIndex = Number(table.apiRange.startColumnIndex || 0) + Number(relativeColumnIndex || 0);
  return {
    sheetId: getNativeSheetId_(table),
    startRowIndex: rowNumber - 1,
    endRowIndex: rowNumber,
    startColumnIndex,
    endColumnIndex: startColumnIndex + Number(width || 1),
  };
}

function nativeColumnBlockRowsRange_(table, rowNumber, relativeColumnIndex, width, height) {
  const rowCount = Math.max(1, Number(height || 1));
  validateNativeRowNumber_(table, rowNumber);
  validateNativeRowNumber_(table, rowNumber + rowCount - 1);
  const startColumnIndex = Number(table.apiRange.startColumnIndex || 0) + Number(relativeColumnIndex || 0);
  return {
    sheetId: getNativeSheetId_(table),
    startRowIndex: rowNumber - 1,
    endRowIndex: rowNumber - 1 + rowCount,
    startColumnIndex,
    endColumnIndex: startColumnIndex + Number(width || 1),
  };
}

function getNativeSheetId_(table) {
  const sheetId = table.apiRange && table.apiRange.sheetId != null
    ? Number(table.apiRange.sheetId)
    : Number(table.gid);
  if (!Number.isFinite(sheetId)) {
    throw new Error(`Missing sheet id for native table: ${table.name || table.apiTableName}`);
  }
  return sheetId;
}

function validateNativeRowNumber_(table, rowNumber) {
  if (!isNativeRowNumberInRange_(table, rowNumber)) {
    throw new Error('Invalid rowNumber.');
  }
}

function isNativeRowNumberInRange_(table, rowNumber) {
  const firstDataRow = Number(table.apiRange.startRowIndex || 0) + 2;
  const endRow = Number(table.apiRange.endRowIndex || firstDataRow);
  return !!(rowNumber && rowNumber >= firstDataRow && rowNumber <= endRow);
}

function toCellData_(value) {
  return { userEnteredValue: toUserEnteredValue_(value) };
}

function toUserEnteredValue_(value) {
  if (value == null || value === '') return {};
  if (typeof value === 'boolean') return { boolValue: value };
  if (typeof value === 'number') return { numberValue: value };
  const text = String(value);
  const number = Number(text);
  if (/^-?\d+(\.\d+)?$/.test(text.trim()) && Number.isFinite(number)) {
    return { numberValue: number };
  }
  return { stringValue: text };
}

function payloadValue_(payload, header) {
  const value = payload.record && Object.prototype.hasOwnProperty.call(payload.record, header)
    ? payload.record[header]
    : '';
  return value == null ? '' : value;
}

function toRecord_(headers, values) {
  return headers.reduce((record, header, index) => {
    record[header] = values[index] || '';
    return record;
  }, {});
}

function createSummary_(headers, rows) {
  const statusHeader = headers.find((header) => /status|stage/i.test(header));
  const ownerHeader = headers.find((header) => /owner|team|department|position|channel/i.test(header));
  const statusCounts = {};
  const ownerCounts = {};

  rows.forEach((row) => {
    if (statusHeader) {
      const value = row.record[statusHeader] || 'Blank';
      statusCounts[value] = (statusCounts[value] || 0) + 1;
    }
    if (ownerHeader) {
      const value = row.record[ownerHeader] || 'Blank';
      ownerCounts[value] = (ownerCounts[value] || 0) + 1;
    }
  });

  return {
    rows: rows.length,
    columns: headers.length,
    statusHeader: statusHeader || '',
    ownerHeader: ownerHeader || '',
    statusCounts,
    ownerCounts,
  };
}

function validatePayload_(payload) {
  if (!payload || !payload.tableKey || !payload.record) {
    throw new Error('Missing save payload.');
  }
}

function validateImportPayload_(payload) {
  if (!payload || !payload.tableKey || !Array.isArray(payload.records)) {
    throw new Error('Missing import payload.');
  }
  if (!payload.records.length) {
    throw new Error('Import file has no data rows.');
  }
  if (payload.records.length > 500) {
    throw new Error('Import supports up to 500 rows at a time.');
  }
}
