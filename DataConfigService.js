/**
 * Source-aware data config adapter.
 *
 * Table/column/relationship config now lives in department data files. The
 * process app reads only the data files needed by tables declared in UI Page
 * config, then routes config writes back to the owning data file.
 */
const DATA_CONFIG_ROW_OFFSET = 100000;
const DATA_CONFIG_SHEETS = {
  tableIndex: { sheetName: 'Table index', label: 'Table index' },
  dataColumn: { sheetName: 'Column index', label: 'Data column' },
  relationship: { sheetName: 'Relationship', label: 'Relationship' },
};

function isDataConfigKey_(configKey) {
  return !!DATA_CONFIG_SHEETS[configKey];
}

function resolveDataConfigItem_(definition, type, group, forceRefresh) {
  const sources = getProcessDataConfigSources_(forceRefresh);
  const sheetDef = DATA_CONFIG_SHEETS[definition.key] || {};
  const tableRefs = sources.map((source, index) => createDataConfigTableRef_(definition.key, sheetDef, source, index)).filter(Boolean);
  const connected = tableRefs.length > 0;
  return {
    key: definition.key,
    label: definition.label,
    typeKey: type.key,
    typeLabel: type.label,
    groupKey: group.key,
    groupLabel: group.label,
    aliases: definition.aliases || [definition.label],
    table: tableRefs[0] || null,
    tableRefs,
    dataConfig: true,
    connected,
    previewMode: inferConfigPreviewMode_(definition.key, group.key),
    bindingLabel: connected
      ? `Data files: ${tableRefs.length}`
      : 'No department data file found from UI Page config tables',
  };
}

function getProcessDataConfigSources_(forceRefresh) {
  const cacheKey = cacheKey_('registry', 'dataConfigSources');
  const cached = getCached_(cacheKey);
  if (!forceRefresh && cached) return cached;

  const tables = getTables_(forceRefresh);
  const bySpreadsheetId = {};
  Object.keys(tables || {}).forEach((key) => {
    const table = tables[key];
    if (!table || table.type === 'config' || table.type === 'system') return;
    const spreadsheetId = table.spreadsheetId || SPREADSHEET_ID;
    if (!spreadsheetId) return;

    bySpreadsheetId[spreadsheetId] = bySpreadsheetId[spreadsheetId] || {
      spreadsheetId,
      spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
      department: table.department || '',
      group: table.group || '',
      tables: [],
      tableNameSet: {},
    };
    const source = bySpreadsheetId[spreadsheetId];
    addDataConfigSourceTable_(source, table.name);
    addDataConfigSourceTable_(source, table.apiTableName);
    addDataConfigSourceTable_(source, table.sheetName);
    addDataConfigSourceTable_(source, key);
  });

  const result = Object.keys(bySpreadsheetId).map((spreadsheetId) => bySpreadsheetId[spreadsheetId]);
  setCached_(cacheKey, result, TABLE_REGISTRY_CACHE_TTL_SECONDS);
  return result;
}

function addDataConfigSourceTable_(source, tableName) {
  const text = String(tableName || '').trim();
  if (!text) return;
  const lookupNames = getProcessConfiguredTableLookupNames_(text);
  const exists = lookupNames.some((name) => source.tableNameSet[name]);
  lookupNames.forEach((name) => {
    if (name) source.tableNameSet[name] = true;
  });
  if (!exists) source.tables.push(text);
}

function createDataConfigTableRef_(configKey, sheetDef, source, sourceIndex) {
  return {
    key: `data_config_${configKey}_${sourceIndex}`,
    name: sheetDef.label,
    sheetName: sheetDef.sheetName,
    gid: 0,
    type: 'config',
    spreadsheetId: source.spreadsheetId,
    spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${source.spreadsheetId}/edit`,
    binding: 'dataConfigSheet',
    dataConfigKey: configKey,
    dataConfigSourceIndex: sourceIndex,
    dataConfigSource: source,
    dataStartColumn: 1,
    headerRow: 1,
    dataStartRow: 2,
  };
}

function buildDataConfigTableData_(item, forceRefresh) {
  const tableRefs = item.tableRefs || [];
  const dataSets = tableRefs.map((tableRef, sourceIndex) => {
    try {
      const data = buildSystemTableRowsData_(tableRef);
      return Object.assign({}, data, { tableRef, sourceIndex });
    } catch (error) {
      logAppDiagnostic_('warn', 'data_config_source_skipped', {
        configKey: item.key || '',
        sourceIndex,
        spreadsheetId: tableRef.spreadsheetId || '',
        sheetName: tableRef.sheetName || '',
        sourceTables: tableRef.dataConfigSource && tableRef.dataConfigSource.tables
          ? tableRef.dataConfigSource.tables
          : [],
      }, error);
      return {
        tableRef,
        sourceIndex,
        headers: [],
        rows: [],
        skipped: error && error.message ? error.message : String(error),
      };
    }
  });
  const headers = getDataConfigUnionHeaders_(dataSets);
  const rows = dataSets.flatMap((dataSet) => {
    const source = dataSet.tableRef.dataConfigSource || {};
    return (dataSet.rows || [])
      .filter((row) => isDataConfigRowInScope_(item.key, row.record || {}, source))
      .map((row) => createDataConfigAggregateRow_(item.key, headers, row, dataSet.tableRef, dataSet.sourceIndex));
  });

  rows.sort((a, b) => {
    const sourceDelta = Number(a.sourceIndex || 0) - Number(b.sourceIndex || 0);
    if (sourceDelta) return sourceDelta;
    return Number(a.rowNumber || 0) - Number(b.rowNumber || 0);
  });

  return {
    table: tableRefs[0] || null,
    tableRefs,
    headers,
    fields: headers.map((header, index) => ({ name: header, label: header, index })),
    rows,
    summary: createSummary_(headers, rows),
    dataWindow: createDataWindowMeta_(rows.length, false),
  };
}

function getDataConfigUnionHeaders_(dataSets) {
  const seen = {};
  return (dataSets || []).flatMap((dataSet) => dataSet.headers || []).filter((header) => {
    const key = normalizeConfigHeader_(header);
    if (!key || seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function createDataConfigAggregateRow_(configKey, headers, row, tableRef, sourceIndex) {
  const record = {};
  headers.forEach((header) => {
    record[header] = row.record && row.record[header] != null ? row.record[header] : '';
  });
  enrichDataConfigRecord_(configKey, record);
  return {
    rowNumber: encodeDataConfigRowNumber_(sourceIndex, row.rowNumber),
    sourceIndex,
    sourceSpreadsheetId: tableRef.spreadsheetId,
    sourceSheetName: tableRef.sheetName,
    actualRowNumber: row.rowNumber,
    id: row.id || `row-${row.rowNumber}`,
    values: headers.map((header) => record[header] || ''),
    record,
  };
}

function enrichDataConfigRecord_(configKey, record) {
  if (configKey !== 'relationship') return record;
  const fromColumn = parseDataConfigColumnId_(recordValue_(record, ['from_column_id', 'from_column', 'source_column_id']));
  const toColumn = parseDataConfigColumnId_(recordValue_(record, ['to_column_id', 'to_column', 'target_column_id']));
  if (fromColumn.table) {
    record['Source table'] = record['Source table'] || fromColumn.table;
    record['Source field'] = record['Source field'] || fromColumn.column;
  }
  if (toColumn.table) {
    record['Target table'] = record['Target table'] || toColumn.table;
    record['Target field'] = record['Target field'] || toColumn.column;
  }
  record.Status = record.Status || 'Active';
  return record;
}

function isDataConfigRowInScope_(configKey, record, source) {
  const tableNames = getDataConfigRecordTableNames_(configKey, record);
  if (!tableNames.length) return false;
  return tableNames.some((tableName) => isDataConfigSourceTable_(source, tableName));
}

function getDataConfigRecordTableNames_(configKey, record) {
  if (configKey === 'tableIndex') {
    return [recordValue_(record, ['table', 'table_name', 'source_table'])].filter(Boolean);
  }
  if (configKey === 'dataColumn') {
    return [recordValue_(record, ['table', 'table_name', 'source_table'])].filter(Boolean);
  }
  if (configKey === 'relationship') {
    const fromColumn = parseDataConfigColumnId_(recordValue_(record, ['from_column_id', 'from_column', 'source_column_id']));
    const toColumn = parseDataConfigColumnId_(recordValue_(record, ['to_column_id', 'to_column', 'target_column_id']));
    return [
      recordValue_(record, ['source_table', 'parent_table', 'from_table', 'table_name', 'table']) || fromColumn.table,
      recordValue_(record, ['target_table', 'child_table', 'to_table', 'related_table']) || toColumn.table,
    ].filter(Boolean);
  }
  return [];
}

function isDataConfigSourceTable_(source, tableName) {
  const tableNameSet = (source && source.tableNameSet) || {};
  return getProcessConfiguredTableLookupNames_(tableName).some((name) => tableNameSet[name]);
}

function encodeDataConfigRowNumber_(sourceIndex, rowNumber) {
  return (Number(sourceIndex || 0) + 1) * DATA_CONFIG_ROW_OFFSET + Number(rowNumber || 0);
}

function decodeDataConfigRowNumber_(rowNumber) {
  const encoded = Number(rowNumber || 0);
  if (!encoded || encoded < DATA_CONFIG_ROW_OFFSET) return { sourceIndex: 0, rowNumber: encoded };
  return {
    sourceIndex: Math.floor(encoded / DATA_CONFIG_ROW_OFFSET) - 1,
    rowNumber: encoded % DATA_CONFIG_ROW_OFFSET,
  };
}

function saveDataConfigRecord_(item, payload) {
  const target = resolveDataConfigWriteTarget_(item, payload);
  const record = normalizeDataConfigRecordForWrite_(item.key, target.tableRef, payload.record || {});
  const actualPayload = Object.assign({}, payload, {
    rowNumber: target.rowNumber,
    record,
  });
  const rowNumber = saveSystemRecord_(target.tableRef, actualPayload);
  return encodeDataConfigRowNumber_(target.sourceIndex, rowNumber);
}

function deleteDataConfigRecord_(item, payload) {
  const decoded = decodeDataConfigRowNumber_(payload.rowNumber);
  const tableRef = (item.tableRefs || [])[decoded.sourceIndex];
  if (!tableRef || !decoded.rowNumber) throw new Error('Cannot resolve source data config row.');
  clearSystemRecord_(tableRef, { rowNumber: decoded.rowNumber });
  return payload.rowNumber;
}

function resolveDataConfigWriteTarget_(item, payload) {
  const decoded = decodeDataConfigRowNumber_(payload && payload.rowNumber);
  if (payload && payload.rowNumber) {
    const tableRef = (item.tableRefs || [])[decoded.sourceIndex];
    if (!tableRef) throw new Error('Cannot resolve source data config file.');
    return { tableRef, sourceIndex: decoded.sourceIndex, rowNumber: decoded.rowNumber };
  }

  const record = (payload && payload.record) || {};
  const tableName = inferDataConfigRecordOwnerTable_(item.key, record);
  const sourceIndex = Math.max(0, (item.tableRefs || []).findIndex((tableRef) => (
    isDataConfigSourceTable_(tableRef.dataConfigSource || {}, tableName)
  )));
  const tableRef = (item.tableRefs || [])[sourceIndex];
  if (!tableRef) throw new Error('No source data config file is available for this record.');
  return { tableRef, sourceIndex, rowNumber: null };
}

function inferDataConfigRecordOwnerTable_(configKey, record) {
  if (configKey === 'tableIndex' || configKey === 'dataColumn') {
    return recordValue_(record, ['table', 'table_name', 'source_table']) || '';
  }
  if (configKey === 'relationship') {
    const fromColumn = parseDataConfigColumnId_(recordValue_(record, ['from_column_id', 'from_column', 'source_column_id']));
    return recordValue_(record, ['source_table', 'parent_table', 'from_table', 'table_name', 'table']) || fromColumn.table;
  }
  return '';
}

function normalizeDataConfigRecordForWrite_(configKey, tableRef, record) {
  const headers = getSystemTableLayout_(getSheet_(tableRef), tableRef).headers || [];
  const next = {};
  headers.forEach((header) => {
    next[header] = record[header] != null ? record[header] : '';
  });

  if (configKey === 'relationship') {
    const sourceTable = recordValue_(record, ['source_table', 'parent_table', 'from_table', 'table_name', 'table']) ||
      parseDataConfigColumnId_(recordValue_(record, ['from_column_id', 'from_column', 'source_column_id'])).table;
    const sourceField = recordValue_(record, ['source_field', 'parent_field', 'from_field', 'key_field', 'field']) ||
      parseDataConfigColumnId_(recordValue_(record, ['from_column_id', 'from_column', 'source_column_id'])).column;
    const targetTable = recordValue_(record, ['target_table', 'child_table', 'to_table', 'related_table']) ||
      parseDataConfigColumnId_(recordValue_(record, ['to_column_id', 'to_column', 'target_column_id'])).table;
    const targetField = recordValue_(record, ['target_field', 'child_field', 'to_field', 'lookup_field', 'related_field']) ||
      parseDataConfigColumnId_(recordValue_(record, ['to_column_id', 'to_column', 'target_column_id'])).column;
    const fromColumnId = makeDataConfigColumnId_(sourceTable, sourceField);
    const toColumnId = makeDataConfigColumnId_(targetTable, targetField);
    setDataConfigValue_(next, headers, ['relationship_id', 'id', 'key', 'config_key'], `${fromColumnId} -> ${toColumnId}`);
    setDataConfigValue_(next, headers, ['from_column_id', 'from_column', 'source_column_id'], fromColumnId);
    setDataConfigValue_(next, headers, ['relationship_type', 'type'], recordValue_(record, ['relationship_type', 'type']) || 'n-1');
    setDataConfigValue_(next, headers, ['to_column_id', 'to_column', 'target_column_id'], toColumnId);
  }
  return next;
}

function setDataConfigValue_(record, headers, names, value) {
  const normalized = (names || []).map(normalizeConfigHeader_);
  const header = (headers || []).find((candidate) => normalized.indexOf(normalizeConfigHeader_(candidate)) >= 0);
  if (header) record[header] = value;
}

function parseDataConfigColumnId_(value) {
  const text = String(value || '').trim();
  if (!text) return { table: '', column: '' };
  const parts = text.split(/\s*>\s*/);
  if (parts.length >= 2) {
    return {
      table: parts[0].trim(),
      column: parts.slice(1).join(' > ').trim(),
    };
  }
  return { table: '', column: text };
}

function makeDataConfigColumnId_(tableName, columnName) {
  const table = String(tableName || '').trim();
  const column = String(columnName || '').trim();
  return table && column ? `${table} > ${column}` : table || column;
}
