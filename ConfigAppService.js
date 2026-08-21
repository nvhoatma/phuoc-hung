/**
 * Config App service.
 *
 * Builds a config-driven editor over native Google Sheets Tables.
 * Config runtime intentionally does not fall back to sheet ranges.
 */
const CONFIG_APP_TYPES = [
  {
    key: 'app',
    label: 'App setting',
    accent: 'blue',
    groups: [
      {
        key: 'ui',
        label: 'UI',
        items: [
          configItem_('page', 'Page', ['Page', 'Page config', 'UI Page config', 'UI_Page_Config']),
          configItem_('view', 'View', ['View', 'View config', 'UI View config', 'UI_View_Config']),
          configItem_('form', 'Form', ['Form', 'Form config', 'UI Form config', 'UI_Form_Config']),
          configItem_('component', 'Component', ['Component', 'Component config', 'UI Component config', 'UI_Component_Config']),
          configItem_('navigation', 'Navigation', ['Nav config', 'Nav_Config', 'Setting Nav', 'Setting_Nav', 'UI Navigation config', 'UI_Navigation_Config', 'Navigation config', 'Navigation']),
          configItem_('appConfig', 'App config', ['App config', 'App_Config', 'Application config', 'UI App config', 'UI_App_Config']),
        ],
      },
      {
        key: 'data',
        label: 'Data',
        items: [
          configItem_('tableIndex', 'Table index', ['Table index', 'Data Table index', 'Table_Index', 'Data_Table_Index']),
          configItem_('dataColumn', 'Data column', ['Column index', 'Data column', 'Data Data column', 'Data_Column', 'Data_Data_Column', 'Column config', 'Table column']),
          configItem_('relationship', 'Relationship', ['Relationship', 'Data Relationship', 'Relationship config', 'Data relationship']),
        ],
      },
      {
        key: 'service',
        label: 'Service',
        items: [
          configItem_('dataServices', 'Data services', ['Data services', 'Data service', 'Data_Service', 'Service config', 'Service Data services', 'Service_Data_Services']),
          configItem_('appPermission', 'App setting permission', ['App setting permission', 'App Setting Permission', 'Setting permission', 'App permission']),
        ],
      },
    ],
  },
  {
    key: 'business',
    label: 'Business setting',
    accent: 'green',
    groups: [
      {
        key: 'general',
        label: 'General',
        items: [
          configItem_('projectInfo', 'Project info', ['Project info', 'Business Project info', 'Business_Project_Info', 'General Project info']),
          configItem_('step', 'Step', ['Step', 'Business Step', 'Business_Step', 'General Step']),
        ],
      },
      {
        key: 'report',
        label: 'Report',
        items: [
          configItem_('metric', 'Metric', ['Metric', 'Report Metric', 'Metric_Config', 'Metric Config', 'Report Metric config']),
          configItem_('reportPlanning', 'Report planning', ['Report planning', 'Report config', 'Report_Config', 'Overview_Config', 'Report Planning config']),
        ],
      },
    ],
  },
];

function configItem_(key, label, aliases) {
  return {
    key,
    label,
    aliases: aliases && aliases.length ? aliases : [label],
  };
}

function getConfigAppBootstrap(options) {
  assertConfigAppAccess_();
  const skipAutoSync = options && options.skipAutoSync;
  const autoSync = skipAutoSync ? { addedCount: 0, skipped: ['skip_auto_sync'] } : syncConfigAppGeneratedRows_();
  const appDefaults = skipAutoSync ? { addedCount: 0, skipped: ['skip_auto_sync'] } : upsertUserAppControlDefaults_(['appConfig']);
  const registry = getConfigAppRegistry_(!(skipAutoSync || autoSync.cached));
  return sanitizeClientPayload_({
    types: registry.types,
    items: registry.items,
    summary: registry.summary,
    autoSync,
    appDefaults,
    generatedAt: new Date().toISOString(),
  });
}

function getConfigTableData(configKey, forceRefresh) {
  assertConfigAppAccess_();
  let item = getConfigItemOrThrow_(configKey, forceRefresh);
  if (configKey === 'view' && item && item.table) {
    const schema = ensureConfigViewColumnOrderSchema_(item);
    if (schema && Number(schema.addedCount || 0) > 0) {
      item = getConfigItemOrThrow_(configKey, true);
      forceRefresh = true;
    }
  }
  const cacheKey = cacheKey_('configTable', configKey);
  const cached = getCached_(cacheKey);
  if (!forceRefresh && cached) {
    return sanitizeClientPayload_(Object.assign({}, cached, { configItem: item }));
  }
  if (!item.table) {
    return sanitizeClientPayload_({
      configItem: item,
      table: null,
      headers: [],
      fields: [],
      rows: [],
      summary: createSummary_([], []),
    });
  }

  if (item.dataConfig) {
    const data = buildDataConfigTableData_(item, forceRefresh);
    setCached_(cacheKey, data, CONFIG_TABLE_CACHE_TTL_SECONDS);
    return sanitizeClientPayload_(Object.assign({}, data, { configItem: item }));
  }

  let data;
  try {
    data = buildConfigTableData_(item.table, forceRefresh);
  } catch (error) {
    if (isQuotaError_(error) && cached) {
      return sanitizeClientPayload_(Object.assign({}, cached, {
        configItem: item,
        stale: true,
        quotaLimited: true,
        message: quotaErrorMessage_(),
      }));
    }
    throw toFriendlyQuotaError_(error);
  }
  setCached_(cacheKey, data, CONFIG_TABLE_CACHE_TTL_SECONDS);
  return sanitizeClientPayload_(Object.assign({}, data, {
    configItem: item,
  }));
}

function buildConfigTableData_(table, forceRefresh) {
  const data = buildTableData_(table, forceRefresh);
  if (!isNativeReadableTable_(table) || (data.rows || []).length || !(data.headers || []).length) return data;

  const fallback = buildConfigSheetRangeFallbackData_(table, forceRefresh);
  return (fallback.rows || []).length ? fallback : data;
}

function buildConfigSheetRangeFallbackData_(table, forceRefresh) {
  const range = table.apiRange || {};
  const fallbackTable = Object.assign({}, table, {
    binding: 'configSheetRange',
    dataStartColumn: Number(range.startColumnIndex || 0) + 1,
    headerRow: Number(range.startRowIndex || 0) + 1,
    dataStartRow: Number(range.startRowIndex || 0) + 2,
  });
  const data = buildSystemTableData_(fallbackTable, forceRefresh);
  return Object.assign({}, data, { table });
}

function seedConfigItemIfEmpty_(item, headers) {
  if (!item || !item.table || !headers || !headers.length) return 0;
  const records = getInitialConfigRecords_(item, headers);
  if (!records.length) return 0;
  return isNativeReadableTable_(item.table)
    ? importNativeRecords_(item.table, records)
    : importSystemRecords_(item.table, records);
}

function saveConfigRecord(payload) {
  assertConfigAppAccess_();
  if (!payload || !payload.configKey || !payload.record) {
    throw new Error('Missing config save payload.');
  }

  const item = getConfigItemOrThrow_(payload.configKey, false);
  if (!item.table) throw new Error(`Config table not found: ${item.label}`);
  if (!item.dataConfig) assertTableWritable_(item.table);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const rowNumber = item.dataConfig
      ? saveDataConfigRecord_(item, payload)
      : isNativeReadableTable_(item.table)
      ? saveNativeRecord_(item.table, payload)
      : saveSystemRecord_(item.table, payload);
    invalidateConfigTableCache_(payload.configKey);
    return {
      ok: true,
      rowNumber,
      record: payload.record,
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteConfigRecord(payload) {
  assertConfigAppAccess_();
  if (!payload || !payload.configKey || !payload.rowNumber) {
    throw new Error('Missing config delete payload.');
  }

  const item = getConfigItemOrThrow_(payload.configKey, false);
  if (!item.table) throw new Error(`Config table not found: ${item.label}`);
  if (!item.dataConfig) assertTableWritable_(item.table);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (item.dataConfig) deleteDataConfigRecord_(item, payload);
    else if (isNativeReadableTable_(item.table)) clearNativeRecord_(item.table, payload);
    else clearSystemRecord_(item.table, payload);
    invalidateConfigTableCache_(payload.configKey);
    return {
      ok: true,
      rowNumber: payload.rowNumber,
    };
  } finally {
    lock.releaseLock();
  }
}

function saveFormRelationships(payload) {
  assertConfigAppAccess_();
  const relationships = payload && Array.isArray(payload.relationships) ? payload.relationships : [];
  const hasNestedScope = payload && Object.prototype.hasOwnProperty.call(payload, 'nestedTables');
  if (!relationships.length && !hasNestedScope) return { ok: true };

  const item = getConfigItemOrThrow_('relationship', false);
  if (!item.table) throw new Error('Relationship config table not found.');
  if (!item.dataConfig) assertTableWritable_(item.table);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const data = getConfigTableData('relationship', false);
    const headers = data.headers || [];
    const scopedSource = normalizeName(payload && payload.sourceTable);
    const scopedTargets = hasNestedScope
      ? (payload.nestedTables || []).map(normalizeName).filter(Boolean)
      : [];
    const keepIds = relationships.map((relationship) => {
      const record = createRelationshipConfigRecord_(headers, relationship);
      return normalizeName(recordValue_(record, ['relationship_id', 'id', 'key', 'config_key']));
    }).filter(Boolean);
    if (scopedSource && hasNestedScope) {
      (data.rows || []).forEach((row) => {
        const record = row.record || {};
        const rowSource = normalizeName(recordValue_(record, ['source_table', 'parent_table', 'from_table', 'table_name', 'table']));
        if (rowSource !== scopedSource) return;
        const rowTarget = normalizeName(recordValue_(record, ['target_table', 'child_table', 'to_table', 'related_table']));
        const rowId = normalizeName(recordValue_(record, ['relationship_id', 'id', 'key', 'config_key']));
        const targetStillNested = rowTarget && scopedTargets.indexOf(rowTarget) >= 0;
        const relationshipStillCurrent = rowId && keepIds.indexOf(rowId) >= 0;
        if (targetStillNested && relationshipStillCurrent) return;
        if (item.dataConfig) deleteDataConfigRecord_(item, { rowNumber: row.rowNumber });
        else if (isNativeReadableTable_(item.table)) clearNativeRecord_(item.table, { rowNumber: row.rowNumber });
        else clearSystemRecord_(item.table, { rowNumber: row.rowNumber });
      });
    }
    relationships.forEach((relationship) => {
      const record = createRelationshipConfigRecord_(headers, relationship);
      const id = normalizeName(recordValue_(record, ['relationship_id', 'id', 'key', 'config_key']));
      const existing = (data.rows || []).find((row) => (
        normalizeName(recordValue_(row.record || {}, ['relationship_id', 'id', 'key', 'config_key'])) === id
      ));
      if (existing) {
        if (item.dataConfig) saveDataConfigRecord_(item, { rowNumber: existing.rowNumber, record });
        else if (isNativeReadableTable_(item.table)) saveNativeRecord_(item.table, { rowNumber: existing.rowNumber, record });
        else saveSystemRecord_(item.table, { rowNumber: existing.rowNumber, record });
      } else if (item.dataConfig) {
        saveDataConfigRecord_(item, { record });
      } else if (isNativeReadableTable_(item.table)) {
        importNativeRecords_(item.table, [record]);
      } else {
        importSystemRecords_(item.table, [record]);
      }
    });
    invalidateConfigTableCache_('relationship');
    return {
      ok: true,
    };
  } finally {
    lock.releaseLock();
  }
}

function createRelationshipConfigRecord_(headers, relationship) {
  const sourceTable = relationship.sourceTable || '';
  const targetTable = relationship.targetTable || '';
  const sourceField = relationship.sourceField || '';
  const targetField = relationship.targetField || '';
  const id = [
    sourceTable,
    targetTable,
    sourceField,
    targetField,
  ].map(normalizeName).filter(Boolean).join('__');
  const record = (headers || []).reduce((acc, header) => {
    acc[header] = '';
    return acc;
  }, {});
  if ((headers || []).some((header) => normalizeConfigHeader_(header) === 'from_column_id')) {
    const fromColumnId = makeDataConfigColumnId_(sourceTable, sourceField);
    const toColumnId = makeDataConfigColumnId_(targetTable, targetField);
    setRelationshipValue_(record, headers, ['relationship_id', 'id', 'key', 'config_key'], `${fromColumnId} -> ${toColumnId}`);
    setRelationshipValue_(record, headers, ['from_column_id', 'from_column', 'source_column_id'], fromColumnId);
    setRelationshipValue_(record, headers, ['relationship_type', 'type'], 'n-1');
    setRelationshipValue_(record, headers, ['to_column_id', 'to_column', 'target_column_id'], toColumnId);
    record['Source table'] = sourceTable;
    record['Source field'] = sourceField;
    record['Target table'] = targetTable;
    record['Target field'] = targetField;
    return record;
  }
  setRelationshipValue_(record, headers, ['relationship_id', 'id', 'key', 'config_key'], id);
  setRelationshipValue_(record, headers, ['item', 'label', 'title', 'name'], `${sourceTable} to ${targetTable}`);
  setRelationshipValue_(record, headers, ['source_table', 'parent_table', 'from_table', 'table_name', 'table'], sourceTable);
  setRelationshipValue_(record, headers, ['source_field', 'parent_field', 'from_field', 'key_field', 'field'], sourceField);
  setRelationshipValue_(record, headers, ['target_table', 'child_table', 'to_table', 'related_table'], targetTable);
  setRelationshipValue_(record, headers, ['target_field', 'child_field', 'to_field', 'lookup_field', 'related_field'], targetField);
  setRelationshipValue_(record, headers, ['relationship_type', 'type'], 'Lookup');
  setRelationshipValue_(record, headers, ['status'], 'Active');
  return record;
}

function setRelationshipValue_(record, headers, names, value) {
  const normalized = names.map(normalizeConfigHeader_);
  const header = (headers || []).find((candidate) => normalized.includes(normalizeConfigHeader_(candidate)));
  if (header) record[header] = value;
}

function getConfigAppRegistry_(forceRefresh) {
  const nativeRegistry = getNativeSpreadsheetTables_(forceRefresh);
  const nativeByName = nativeRegistry.tables.reduce((acc, table) => {
    getNativeTableLookupNames_(table.name).forEach((name) => {
      acc[name] = table;
    });
    return acc;
  }, {});
  const items = {};
  const types = CONFIG_APP_TYPES.map((type) => Object.assign({}, type, {
    groups: type.groups.map((group) => Object.assign({}, group, {
      items: group.items.map((definition) => {
        const resolved = resolveConfigItem_(definition, type, group, nativeByName, nativeRegistry, forceRefresh);
        items[resolved.key] = resolved;
        return resolved;
      }),
    })),
  }));

  const allItems = Object.keys(items).map((key) => items[key]);
  return {
    types,
    items,
    summary: {
      total: allItems.length,
      connected: allItems.filter((item) => !!item.table).length,
      native: allItems.filter((item) => item.table && isNativeReadableTable_(item.table)).length,
      missing: allItems.filter((item) => !item.table).length,
    },
  };
}

function resolveConfigItem_(definition, type, group, nativeByName, nativeRegistry, forceRefresh) {
  if (group.key === 'data' && isDataConfigKey_(definition.key)) {
    return resolveDataConfigItem_(definition, type, group, forceRefresh);
  }

  const aliases = definition.aliases || [definition.label];
  const nativeTable = aliases.reduce((found, alias) => (
    found || nativeByName[normalizeTableLookupName_(alias)] || null
  ), null);

  const base = {
    key: definition.key,
    label: definition.label,
    typeKey: type.key,
    typeLabel: type.label,
    groupKey: group.key,
    groupLabel: group.label,
    aliases,
    table: null,
    connected: false,
    previewMode: inferConfigPreviewMode_(definition.key, group.key),
  };

  if (nativeTable) {
    const sheetMeta = nativeRegistry.sheetsById[String(nativeTable.range.sheetId)] || {};
    const table = {
      key: `config_${definition.key}`,
      name: nativeTable.name,
      sheetName: sheetMeta.title || definition.label,
      gid: nativeTable.range.sheetId,
      type: 'config',
      spreadsheetId: SPREADSHEET_ID,
      spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit#gid=${nativeTable.range.sheetId}`,
      binding: 'nativeTable',
      apiTableId: nativeTable.tableId,
      apiTableName: nativeTable.name,
      apiRange: nativeTable.range,
      columnProperties: getNativeTableColumnProperties_(nativeTable),
      columns: getNativeTableColumns_(nativeTable),
    };
    return Object.assign(base, {
      table,
      connected: true,
      bindingLabel: `Native table: ${nativeTable.name}`,
    });
  }

  return Object.assign(base, {
    bindingLabel: 'No matching native Google Sheets Table found',
  });
}

function inferConfigPreviewMode_(itemKey, groupKey) {
  if (groupKey === 'data' || groupKey === 'service' || ['dataServices', 'relationship', 'dataColumn', 'appPermission'].includes(itemKey)) return 'none';
  if (itemKey === 'navigation') return 'navigation';
  if (['appConfig', 'page', 'view', 'component'].includes(itemKey)) return 'ui';
  if (['metric', 'reportPlanning'].includes(itemKey)) return 'report';
  return 'record';
}

function getConfigItemOrThrow_(configKey, forceRefresh) {
  const registry = getConfigAppRegistry_(forceRefresh);
  const item = registry.items[configKey];
  if (!item) throw new Error(`Unknown config key: ${configKey}`);
  return item;
}
