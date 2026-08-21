/**
 * Maintenance utilities for config normalization/sync/debug actions.
 *
 * These functions are invoked by explicit doGet action parameters or manual
 * script runs. They are not part of the normal page render or CRUD flow.
 */
function normalizeOverviewConfigSetup_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const navSheet = getSheetByNameOrId_(spreadsheet, '≡ Nav', 1783949092);
  const reportSheet = getSheetByNameOrId_(spreadsheet, 'Report config', 600009002);
  const metricSheet = getSheetByNameOrId_(spreadsheet, 'Metric', 600009003);

  if (!navSheet || !reportSheet || !metricSheet) {
    throw new Error('Missing required config sheet.');
  }

  clearOverviewNavSeed_(navSheet);
  prepareReportConfigSheet_(reportSheet);
  prepareMetricConfigSheet_(metricSheet);
  const nativeColumnSync = syncConfigNativeTableColumnTypes_();
  invalidateDataCache_();

  return {
    ok: true,
    reportRows: 0,
    metricRows: 0,
    nativeColumnSync,
    webAppUrl: ScriptApp.getService().getUrl(),
    updatedAt: new Date().toISOString(),
  };
}

function clearOverviewNavSeed_(sheet) {
  sheet.getRange('V6:V8').clearContent();
}

function writeWebAppLink_(sheet) {
  const url = ScriptApp.getService().getUrl();
  if (!url) return;
  const cell = sheet.getRange('I4');
  const text = cell.getDisplayValue() || 'Open app';
  cell.setRichTextValue(SpreadsheetApp.newRichTextValue()
    .setText(text)
    .setLinkUrl(url)
    .build());
}

function prepareReportConfigSheet_(sheet) {
  const headers = [
    'Report', 'Item', 'item_type', 'title', 'subtitle', 'label', 'table_name',
    'metric_type', 'field', 'match_pattern', 'limit', 'sort_order', 'status',
    'source', 'note', 'metric_key',
  ];
  sheet.getRange(1, 8, 80, headers.length).clearContent();
  sheet.getRange(1, 8, 1, headers.length).setValues([headers]);

  applyListValidation_(sheet.getRange(2, 10, 199, 1), getSupportedReportItemTypes_(), true);
  applyListValidation_(sheet.getRange(2, 15, 199, 1), getSupportedMetricTypes_(), false);
  applyListValidation_(sheet.getRange(2, 20, 199, 1), ['Active', 'Draft', 'Inactive'], true);
}

function prepareMetricConfigSheet_(sheet) {
  const headers = [
    'metric_key', 'label', 'source_table', 'metric_type', 'field_expression',
    'filter_expression', 'format', 'status', 'source', 'note',
  ];
  sheet.getRange(1, 8, 80, headers.length).clearContent();
  sheet.getRange(1, 8, 1, headers.length).setValues([headers]);

  applyListValidation_(sheet.getRange(2, 11, 199, 1), getSupportedMetricTypes_(), false);
  applyListValidation_(sheet.getRange(2, 15, 199, 1), ['Active', 'Draft', 'Inactive'], true);
}

function applyListValidation_(range, values, strict) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(!strict)
    .build();
  range.setDataValidation(rule);
}

function syncConfigNativeTableColumnTypes_() {
  const definitions = [
    {
      tableNames: ['Report_Config', 'Overview_Config', 'Report config'],
      columns: {
        item_type: getSupportedReportItemTypes_(),
        metric_type: getSupportedMetricTypes_(),
        status: ['Active', 'Draft', 'Inactive'],
      },
    },
    {
      tableNames: ['Metric_Config', 'Metric Config', 'Metric'],
      columns: {
        metric_type: getSupportedMetricTypes_(),
        status: ['Active', 'Draft', 'Inactive'],
      },
    },
  ];

  const requests = [];
  const synced = [];
  definitions.forEach((definition) => {
    const table = definition.tableNames.reduce((found, tableName) => (
      found || getNativeTableByName_(tableName, true)
    ), null);
    if (!table) return;

    const columnProperties = (table.columnProperties || []).map((column) => {
      const normalizedName = normalizeOverviewHeader_(column.columnName);
      const allowedValues = definition.columns[normalizedName];
      const nextColumn = {
        columnIndex: column.columnIndex,
        columnName: column.columnName,
        columnType: column.columnType || 'TEXT',
      };

      if (allowedValues) {
        nextColumn.columnType = 'DROPDOWN';
        nextColumn.dataValidationRule = {
          condition: {
            type: 'ONE_OF_LIST',
            values: allowedValues.map((value) => ({ userEnteredValue: value })),
          },
        };
      } else if (column.dataValidationRule) {
        nextColumn.dataValidationRule = column.dataValidationRule;
      }

      return nextColumn;
    });

    const syncedColumns = columnProperties
      .filter((column) => definition.columns[normalizeOverviewHeader_(column.columnName)])
      .map((column) => ({
        columnName: column.columnName,
        columnType: column.columnType,
        optionCount: column.dataValidationRule &&
          column.dataValidationRule.condition &&
          column.dataValidationRule.condition.values
          ? column.dataValidationRule.condition.values.length
          : 0,
      }));

    requests.push({
      updateTable: {
        table: {
          tableId: table.tableId,
          columnProperties,
        },
        fields: 'columnProperties',
      },
    });
    synced.push({
      tableName: table.name,
      columns: syncedColumns,
    });
  });

  if (requests.length) {
    sheetsApiBatchUpdate_(SPREADSHEET_ID, requests);
  }

  return {
    tableCount: requests.length,
    tables: synced,
  };
}

function syncConfigNativeTableColumnSchema_() {
  const registry = getConfigAppRegistry_(true);
  const requests = [];
  const synced = [];
  const skipped = [];

  Object.keys(registry.items || {}).forEach((configKey) => {
    const item = registry.items[configKey];
    const table = item && item.table;
    if (!table || !isNativeReadableTable_(table)) {
      skipped.push({ key: configKey, label: item ? item.label : configKey, reason: 'missing_native_table' });
      return;
    }

    const width = getNativeTableWidth_(table);
    const headers = readNativeTableHeaderRow_(table);
    if (!width || !headers.some(Boolean)) {
      skipped.push({ key: configKey, label: item.label, reason: 'missing_header_row' });
      return;
    }

    const existingByIndex = (table.columnProperties || []).reduce((acc, column, fallbackIndex) => {
      const index = getNativeTableColumnRelativeIndex_(column, fallbackIndex);
      acc[index] = column;
      return acc;
    }, {});

    const columnProperties = Array.from({ length: width }, (_, index) => {
      const existing = existingByIndex[index] || {};
      const property = {
        columnIndex: index,
        columnName: headers[index] || existing.columnName || `Column ${index + 1}`,
        columnType: existing.columnType || getInitialConfigColumnProperty_(headers[index] || '', index).columnType || 'TEXT',
      };
      if (existing.dataValidationRule) property.dataValidationRule = existing.dataValidationRule;
      return property;
    });

    const changed = columnProperties.some((column, index) => {
      const existing = existingByIndex[index] || {};
      return column.columnName !== (existing.columnName || '');
    });
    if (!changed) {
      skipped.push({ key: configKey, label: item.label, reason: 'already_synced' });
      return;
    }

    requests.push({
      updateTable: {
        table: {
          tableId: table.apiTableId,
          columnProperties,
        },
        fields: 'columnProperties',
      },
    });
    synced.push({
      key: configKey,
      label: item.label,
      tableName: table.apiTableName || table.name,
      columns: columnProperties.map((column) => column.columnName),
    });
  });

  if (requests.length) {
    sheetsApiBatchUpdate_(SPREADSHEET_ID, requests);
  }

  return {
    syncedCount: synced.length,
    skippedCount: skipped.length,
    synced,
    skipped,
  };
}

function getCanonicalReportRows_() {
  return [
    ['process_overview', 'process_overview_page', 'page', 'Process overview', 'Overview for process setup, steps, tables, and current configuration readiness.', '', '', '', '', '', '', 1, 'Active', 'Report planner template v1', 'Default overview page for process template.', ''],
    ['process_overview', 'step_total_card', 'card', 'Total steps', 'Count all configured process steps.', 'Steps', 'Step', 'row_count', '', '', 1, 10, 'Active', 'Report planner template v1', 'Quick size of the operation flow.', 'tpl_step_total'],
    ['process_overview', 'process_step_card', 'card', 'Process steps', 'Count steps marked as process context.', 'Process steps', 'Step', 'row_count', '', 'Context~Process|process', 1, 20, 'Active', 'Report planner template v1', 'Operational steps available for the app.', 'tpl_step_process_total'],
    ['process_overview', 'table_total_card', 'card', 'Tables', 'Count configured tables from table index.', 'Tables', 'Table index', 'row_count', '', '', 1, 30, 'Active', 'Report planner template v1', 'Data objects registered for this process template.', 'tpl_table_total'],
    ['process_overview', 'tables_by_group', 'group', 'Tables by group', 'Distribution of table index by group.', 'Table group', 'Table index', 'group_count', 'Group', '', 8, 40, 'Active', 'Report planner template v1', 'Helps understand setup coverage by business/system area.', 'tpl_table_by_group'],
    ['process_overview', 'steps_by_category', 'group', 'Steps by category', 'Distribution of steps by category.', 'Step category', 'Step', 'group_count', 'Category', '', 8, 50, 'Active', 'Report planner template v1', 'Shows operation area density.', 'tpl_step_by_category'],
    ['process_overview', 'step_snapshot_table', 'table', 'Step snapshot', 'Latest process steps for quick review.', 'Step snapshot', 'Step', 'row_count', '', '', 10, 60, 'Active', 'Report planner template v1', 'Preview table for the process flow.', ''],
    ['step_overview', 'step_overview_page', 'page', 'Step overview', 'Focused view for process operation steps and ownership setup.', '', '', '', '', '', '', 100, 'Active', 'Report planner template v1', 'Step-specific overview page.', ''],
    ['step_overview', 'step_total_card', 'card', 'Total steps', 'Count all configured steps.', 'Steps', 'Step', 'row_count', '', '', 1, 110, 'Active', 'Report planner template v1', 'Reusable step count metric.', 'tpl_step_total'],
    ['step_overview', 'steps_by_operation', 'group', 'Steps by operation', 'Distribution of steps by operation.', 'Operation', 'Step', 'group_count', 'Operation', '', 8, 120, 'Active', 'Report planner template v1', 'Main operation grouping.', 'tpl_step_by_operation'],
    ['step_overview', 'steps_by_table', 'group', 'Steps by related table', 'Distribution of steps by relating table.', 'Relating table', 'Step', 'group_count', 'Relating table', '', 8, 130, 'Active', 'Report planner template v1', 'Shows table dependency footprint.', 'tpl_step_by_table'],
    ['step_overview', 'step_table', 'table', 'Step list', 'Step rows used by the app navigation and workflow.', 'Step list', 'Step', 'row_count', '', '', 12, 140, 'Active', 'Report planner template v1', 'Preview table for steps.', ''],
    ['table_index_overview', 'table_index_overview_page', 'page', 'Table index overview', 'Focused view for data object setup and native table readiness.', '', '', '', '', '', '', 200, 'Active', 'Report planner template v1', 'Table-index overview page.', ''],
    ['table_index_overview', 'table_total_card', 'card', 'Total tables', 'Count all tables in table index.', 'Tables', 'Table index', 'row_count', '', '', 1, 210, 'Active', 'Report planner template v1', 'Reusable table count metric.', 'tpl_table_total'],
    ['table_index_overview', 'tables_by_context', 'group', 'Tables by context', 'Distribution of tables by context.', 'Context', 'Table index', 'group_count', 'Context', '', 8, 220, 'Active', 'Report planner template v1', 'Data object context mix.', 'tpl_table_by_context'],
    ['table_index_overview', 'tables_by_type', 'group', 'Tables by data type', 'Distribution of tables by data type.', 'Data type', 'Table index', 'group_count', 'Data type', '', 8, 230, 'Active', 'Report planner template v1', 'Setup/reference/transaction split.', 'tpl_table_by_type'],
    ['table_index_overview', 'table_index_table', 'table', 'Table index', 'Configured tables and links.', 'Table index', 'Table index', 'row_count', '', '', 12, 240, 'Active', 'Report planner template v1', 'Preview table for table index.', ''],
  ];
}

function getCanonicalMetricRows_() {
  return [
    ['tpl_step_total', 'Steps', 'Step', 'row_count', '', '', 'number', 'Active', 'Metric engine v2', 'Total rows in Step table.'],
    ['tpl_step_process_total', 'Process steps', 'Step', 'row_count', '', 'Context~Process|process', 'number', 'Active', 'Metric engine v2', 'Steps marked as process context.'],
    ['tpl_step_by_operation', 'Operation', 'Step', 'group_count', 'Operation', '', 'group', 'Active', 'Metric engine v2', 'Group steps by operation.'],
    ['tpl_step_by_category', 'Category', 'Step', 'group_count', 'Category', '', 'group', 'Active', 'Metric engine v2', 'Group steps by category.'],
    ['tpl_step_by_table', 'Relating table', 'Step', 'group_count', 'Relating table', '', 'group', 'Active', 'Metric engine v2', 'Group steps by related table.'],
    ['tpl_table_total', 'Tables', 'Table index', 'row_count', '', '', 'number', 'Active', 'Metric engine v2', 'Total rows in Table index.'],
    ['tpl_table_by_group', 'Table group', 'Table index', 'group_count', 'Group', '', 'group', 'Active', 'Metric engine v2', 'Group tables by group.'],
    ['tpl_table_by_context', 'Context', 'Table index', 'group_count', 'Context', '', 'group', 'Active', 'Metric engine v2', 'Group tables by context.'],
    ['tpl_table_by_type', 'Data type', 'Table index', 'group_count', 'Data type', '', 'group', 'Active', 'Metric engine v2', 'Group tables by data type.'],
    ['tpl_pm_total', 'PM items', 'PM', 'row_count', '', '', 'number', 'Draft', 'Metric engine v2', 'Optional PM prompt/task configuration count.'],
  ];
}

function seedConfigAppTables_() {
  const registry = getConfigAppRegistry_(true);
  const seeded = [];
  const skipped = [];
  const missing = [];

  Object.keys(registry.items || {}).forEach((configKey) => {
    const item = registry.items[configKey];
    if (!item || !item.table) {
      missing.push({ key: configKey, label: item ? item.label : configKey });
      return;
    }

    const data = buildTableData_(item.table, true);
    if ((data.rows || []).length) {
      skipped.push({ key: configKey, label: item.label, reason: 'has_rows', rows: data.rows.length });
      return;
    }

    const records = getInitialConfigRecords_(item, data.headers || []);
    if (!records.length) {
      skipped.push({ key: configKey, label: item.label, reason: 'no_seed_template' });
      return;
    }

    const count = isNativeReadableTable_(item.table)
      ? importNativeRecords_(item.table, records)
      : importSystemRecords_(item.table, records);
    seeded.push({ key: configKey, label: item.label, table: item.table.name, rows: count });
  });

  return {
    seededCount: seeded.length,
    skippedCount: skipped.length,
    missingCount: missing.length,
    seeded,
    skipped,
    missing,
  };
}

function ensureSeedConfigAppTables_() {
  const seeded = seedConfigAppTables_();
  return {
    ensured: {
      createdCount: 0,
      skippedCount: 0,
      created: [],
      skipped: [],
      note: 'Table creation is disabled; only existing config tables are seeded.',
    },
    seeded,
  };
}

function upsertUserAppControlDefaults_(targetKeys) {
  targetKeys = targetKeys && targetKeys.length ? targetKeys : ['appConfig', 'view', 'component', 'dataServices'];
  const added = [];
  const skipped = [];

  targetKeys.forEach((configKey) => {
    const item = getConfigItemOrThrow_(configKey, true);
    if (!item || !item.table) {
      skipped.push({ key: configKey, reason: 'missing_table' });
      return;
    }

    const data = buildTableData_(item.table, true);
    const headers = data.headers || [];
    const keyHeader = findSeedKeyHeader_(headers);
    const existingKeys = (data.rows || []).reduce((acc, row) => {
      const record = row.record || {};
      const key = normalizeConfigSeedHeader_(keyHeader ? record[keyHeader] : recordValueForSeed_(record, ['config_key', 'key', 'id', 'item']));
      if (key) acc[key] = true;
      return acc;
    }, {});

    const records = getInitialConfigRecords_(item, headers).filter((record) => {
      const key = normalizeConfigSeedHeader_(keyHeader ? record[keyHeader] : recordValueForSeed_(record, ['config_key', 'key', 'id', 'item']));
      return key && !existingKeys[key];
    });

    if (!records.length) {
      skipped.push({ key: configKey, label: item.label, reason: 'already_seeded' });
      return;
    }

    const count = isNativeReadableTable_(item.table)
      ? importNativeRecords_(item.table, records)
      : importSystemRecords_(item.table, records);
    added.push({ key: configKey, label: item.label, rows: count });
  });

  return {
    addedCount: added.length,
    skippedCount: skipped.length,
    added,
    skipped,
  };
}

function findSeedKeyHeader_(headers) {
  return (headers || []).find((header) => (
    ['config_key', 'config_name', 'key', 'id', 'name', 'item', 'page_id', 'view_id', 'component_id', 'service_id'].includes(normalizeConfigSeedHeader_(header))
  )) || '';
}

function recordValueForSeed_(record, names) {
  const normalizedNames = names.map(normalizeConfigSeedHeader_);
  const keys = Object.keys(record || {});
  for (let index = 0; index < keys.length; index += 1) {
    if (normalizedNames.includes(normalizeConfigSeedHeader_(keys[index])) && record[keys[index]] !== '') {
      return record[keys[index]];
    }
  }
  return '';
}

function ensureMissingConfigNativeTables_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const registry = getConfigAppRegistry_(true);
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const existingSheetIds = spreadsheet.getSheets().reduce((acc, sheet) => {
      acc[sheet.getSheetId()] = true;
      return acc;
    }, {});
    const existingTableIds = getNativeSpreadsheetTables_(true).tables.reduce((acc, table) => {
      acc[table.tableId] = true;
      return acc;
    }, {});
    const created = [];
    const skipped = [];

    Object.keys(registry.items || {}).forEach((configKey, index) => {
      const item = registry.items[configKey];
      if (!item || (item.table && item.table.binding !== 'sheetRange')) return;

      const headers = getInitialConfigHeaders_(item.key);
      const records = getInitialConfigRecords_(item, headers);
      if (!headers.length || !records.length) {
        skipped.push({ key: configKey, label: item ? item.label : configKey, reason: 'no_seed_shape' });
        return;
      }

      const existingSheet = item.table && item.table.binding === 'sheetRange'
        ? spreadsheet.getSheetByName(item.table.sheetName)
        : findSheetByConfigAliases_(spreadsheet, item);
      const sheet = existingSheet || spreadsheet.insertSheet(item.label);
      let tableHeaders = headers;
      let rowCount = Math.max(records.length, 1);
      let columnCount = headers.length;

      if (sheet.getLastRow() > 1 && sheet.getLastColumn() > 0) {
        tableHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
          .map((header, headerIndex) => header || `Column ${headerIndex + 1}`);
        rowCount = sheet.getLastRow() - 1;
        columnCount = tableHeaders.length;
      } else {
        sheet.clear();
        sheet.setFrozenRows(1);

        const rows = [headers].concat(records.map((record) => (
          headers.map((header) => Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '')
        )));
        sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
        sheet.autoResizeColumns(1, headers.length);
        rowCount = records.length;
      }

      const sheetId = sheet.getSheetId();
      existingSheetIds[sheetId] = true;
      const tableId = makeUniqueConfigTableId_(item.key, existingTableIds, index + 1);
      existingTableIds[tableId] = true;

      sheetsApiBatchUpdate_(SPREADSHEET_ID, [{
        addTable: {
          table: {
            name: item.label,
            tableId,
            range: {
              sheetId,
              startRowIndex: 0,
              endRowIndex: rowCount + 1,
              startColumnIndex: 0,
              endColumnIndex: columnCount,
            },
            columnProperties: tableHeaders.map((header, columnIndex) => (
              getInitialConfigColumnProperty_(header, columnIndex)
            )),
          },
        },
      }]);

      created.push({
        key: item.key,
        label: item.label,
        sheetName: sheet.getName(),
        table: item.label,
        rows: rowCount,
        mode: existingSheet ? 'converted' : 'created',
      });
    });

    return {
      createdCount: created.length,
      skippedCount: skipped.length,
      created,
      skipped,
    };
  } finally {
    lock.releaseLock();
  }
}

function findSheetByConfigAliases_(spreadsheet, item) {
  const aliases = [item.label].concat(item.aliases || []);
  return aliases.reduce((found, alias) => found || spreadsheet.getSheetByName(alias), null);
}

function makeUniqueConfigTableId_(key, existingTableIds, ordinal) {
  const base = `config_${String(key || 'table').replace(/[^A-Za-z0-9_]/g, '_')}`;
  let candidate = base;
  let suffix = ordinal;
  while (existingTableIds[candidate]) {
    candidate = `${base}_${suffix}`;
    suffix += 1;
  }
  return candidate;
}

function getInitialConfigHeaders_(itemKey) {
  const headers = {
    appConfig: ['Config key', 'Category', 'Item', 'Value', 'Color', 'Detail', 'Sort order', 'Status', 'Note'],
    page: ['Page ID', 'Page type', 'Table name', 'Title column', 'Default table view', 'Data mode', 'Report name', 'Link URL', 'AI insight enable', 'AI insight column'],
    view: ['View ID', 'Page ID', 'Item', 'View type', 'Title column', 'Default group by', 'Column order', 'Hidden columns', 'Section layout', 'Filter fields', 'Default filters'],
    form: ['Form ID', 'Page ID', 'Item', 'Mode', 'Form type', 'Columns', 'Label position', 'Show empty fields', 'Main display name', 'Column order', 'Hidden columns', 'Section layout', 'Nested table', 'Status'],
    component: ['Component ID', 'View ID', 'Item', 'Component type', 'Enabled', 'Config value'],
    dataColumn: ['Column ID', 'Category', 'Item', 'Table name', 'Field', 'Label', 'Input type', 'Groupable', 'Default group', 'Group order', 'Hidden', 'Section', 'Sort order', 'Status', 'Note'],
    dataServices: ['Service ID', 'Category', 'Item', 'Service type', 'Table name', 'Detail', 'Sort order', 'Status', 'Note'],
    appPermission: ['Email', 'Role', 'Status'],
  };
  return headers[itemKey] || ['ID', 'Category', 'Item', 'Detail', 'Sort order', 'Status', 'Note'];
}

function getInitialConfigColumnProperty_(header, columnIndex) {
  const normalized = normalizeConfigSeedHeader_(header);
  const property = {
    columnIndex,
    columnName: header,
    columnType: 'TEXT',
  };
  if (normalized.indexOf('status') >= 0) {
    property.columnType = 'DROPDOWN';
    property.dataValidationRule = {
      condition: {
        type: 'ONE_OF_LIST',
        values: ['Active', 'Draft', 'Inactive'].map((value) => ({ userEnteredValue: value })),
      },
    };
  } else if (normalized.indexOf('sort') >= 0 || normalized.indexOf('order') >= 0) {
    property.columnType = 'DOUBLE';
  } else if (['enabled', 'groupable', 'default_group', 'hidden', 'is_default'].includes(normalized)) {
    property.columnType = 'DROPDOWN';
    property.dataValidationRule = {
      condition: {
        type: 'ONE_OF_LIST',
        values: ['TRUE', 'FALSE'].map((value) => ({ userEnteredValue: value })),
      },
    };
  } else if (normalized === 'page_type') {
    property.columnType = 'DROPDOWN';
    property.dataValidationRule = {
      condition: {
        type: 'ONE_OF_LIST',
        values: ['table', 'overview', 'report', 'custom'].map((value) => ({ userEnteredValue: value })),
      },
    };
  } else if (['view_type', 'default_view', 'default_view_type'].includes(normalized)) {
    property.columnType = 'DROPDOWN';
    property.dataValidationRule = {
      condition: {
        type: 'ONE_OF_LIST',
        values: ['table', 'detail', 'card', 'nested_form', 'form'].map((value) => ({ userEnteredValue: value })),
      },
    };
  } else if (normalized === 'ai_insight_enable') {
    property.columnType = 'DROPDOWN';
    property.dataValidationRule = {
      condition: {
        type: 'ONE_OF_LIST',
        values: ['yes', 'no'].map((value) => ({ userEnteredValue: value })),
      },
    };
  }
  return property;
}

function getInitialConfigRecords_(item, headers) {
  const templates = getConfigSeedTemplates_()[item.key] || [];
  if (isLeanKeyValueConfigHeaders_(headers)) return createLeanKeyValueConfigRecords_(item, headers, templates);
  return templates.map((template, index) => createSeedRecord_(item, headers, template, index + 1));
}

function isLeanKeyValueConfigHeaders_(headers) {
  const normalized = (headers || []).map(normalizeConfigSeedHeader_);
  return normalized.includes('config_group') &&
    normalized.includes('config_name') &&
    normalized.includes('config_value');
}

function createLeanKeyValueConfigRecords_(item, headers, templates) {
  const groupHeader = findConfigSeedHeaderByNames_(headers, ['config_group', 'group', 'category']) || headers[0];
  const nameHeader = findConfigSeedHeaderByNames_(headers, ['config_name', 'name', 'item', 'key']) || headers[1];
  const valueHeader = findConfigSeedHeaderByNames_(headers, ['config_value', 'value', 'setting_value']) || headers[2];
  const noteHeader = findConfigSeedHeaderByNames_(headers, ['note', 'detail', 'description']) || headers[3] || '';
  const rows = [];

  templates.forEach((template) => {
    getLeanConfigProperties_(item, template).forEach((property) => {
      const record = {};
      (headers || []).forEach((header) => { record[header] = ''; });
      record[groupHeader] = property.group;
      record[nameHeader] = property.name;
      record[valueHeader] = property.value == null ? '' : property.value;
      if (noteHeader) record[noteHeader] = property.note || template.detail || template.summary || '';
      rows.push(record);
    });
  });

  return rows;
}

function findConfigSeedHeaderByNames_(headers, names) {
  const normalizedNames = names.map(normalizeConfigSeedHeader_);
  return (headers || []).find((header) => normalizedNames.includes(normalizeConfigSeedHeader_(header))) || '';
}

function getLeanConfigProperties_(item, template) {
  if (item.key === 'appConfig') {
    return [{
      group: template.category || item.groupLabel || item.typeLabel || 'UI',
      name: template.label || titleCaseConfigName_(template.key),
      value: template.value || template.color || '',
      note: template.detail || template.summary || '',
    }];
  }

  const group = template.key || `${item.key}_config`;
  const props = [];
  const add = (name, value, note) => {
    if (value === undefined || value === null || value === '') return;
    props.push({ group, name, value, note });
  };

  const idLabels = {
    navigation: 'Nav ID',
    page: 'Page ID',
    view: 'View ID',
    form: 'Form ID',
    component: 'Component ID',
    dataServices: 'Service ID',
    dataColumn: 'Column ID',
    relationship: 'Relationship ID',
    metric: 'Metric key',
    reportPlanning: 'Report',
    tableIndex: 'Table',
  };

  add(idLabels[item.key] || 'Config key', template.key, `${item.label} config id.`);
  add('Item', template.label || template.item, `${item.label} label.`);
  add('Category', template.category || template.type, '');
  add('Parent ID', template.parent_id || template.parent, '');
  add('Nav ID', template.nav_id || template.nav, '');
  add('View ID', template.view_id || template.view_id_key, '');
  add('Page', template.page || template.page_id, '');
  add('Item type', template.item_type || template.type, '');
  add('Page type', template.page_type, '');
  add('Table name', template.table_name || template.table, '');
  add('Field', template.field, '');
  add('Label', template.label, '');
  add('Input type', template.input_type || template.type, '');
  add('Groupable', template.groupable, '');
  add('Default group', template.default_group, '');
  add('Group order', template.group_order, '');
  add('Hidden', template.hidden, '');
  add('View type', template.view_type || template.view, '');
  add('Title column', template.title_column || template.record_title_column || template.record_name_column || template.display_name_column || template.display_column || template.main_display_name || template.main_display_field || template.display_name || template.display_field || template.title_field, '');
  add('Default view', template.default_view || template.view_type || template.view, '');
  add('Default group by', template.default_group_by, '');
  add('Column order', template.column_order, '');
  add('Hidden columns', template.hidden_columns, '');
  add('Section layout', template.section_layout || template.sections, '');
  add('AI insight enable', template.ai_insight_enable || 'no', '');
  add('AI insight column', template.ai_insight_column, '');
  add('Filter fields', template.filter_fields || template.filter_columns || template.smart_filter_fields, '');
  add('Default filters', template.default_filters || template.predefined_filters || template.preset_filters || template.filter_config, '');
  add('Default sort', template.default_sort || template.default_sort_field || template.sort_field, '');
  add('Default sort direction', template.default_sort_direction || template.sort_direction, '');
  add('Mode', template.mode || template.form_mode || template.edit_mode, '');
  add('Form type', template.form_type || template.form_layout_type || template.layout_type, '');
  add('Columns', template.columns, '');
  add('Label position', template.label_position, '');
  add('Show empty fields', template.show_empty_fields, '');
  add('Main display name', template.main_display_name || template.main_display_field || template.display_name || template.display_field || template.title_field, '');
  add('Nested table', template.nested_table || template.nested_tables, '');
  add('Component type', template.component_type, '');
  add('Enabled', template.enabled, '');
  add('Config value', template.config_value || template.value, '');
  add('Service type', template.service_type, '');
  add('Enable create', template.enable_create, '');
  add('Enable edit', template.enable_edit, '');
  add('Enable delete', template.enable_delete, '');
  add('Enable import', template.enable_import, '');
  add('Icon', template.icon, '');
  add('Status', template.status || 'Active', 'Active config row.');
  add('Detail', template.detail || template.summary, '');

  return props;
}

function titleCaseConfigName_(value) {
  return String(value || '')
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/\b\w/g, (letter) => letter.toUpperCase());
}

function createSeedRecord_(item, headers, template, ordinal) {
  return (headers || []).reduce((record, header) => {
    const normalized = normalizeConfigSeedHeader_(header);
    record[header] = seedValueForHeader_(normalized, header, item, template, ordinal);
    return record;
  }, {});
}

function seedValueForHeader_(normalized, header, item, template, ordinal) {
  const direct = seedTemplateValue_(template, normalized);
  if (direct !== '') return direct;

  if (['id', 'key', 'code', 'config_key', 'nav_id', 'page_id', 'view_id', 'component_id', 'service_id', 'relationship_id', 'column_id'].includes(normalized)) {
    return template.key || `${item.key}_${ordinal}`;
  }
  if (['item', 'name', 'title', 'label', 'metric_key', 'table', 'table_name'].includes(normalized) || normalized.indexOf('item') >= 0) {
    return template.label || item.label;
  }
  if (normalized.indexOf('english') >= 0) return template.english || template.label || item.label;
  if (normalized.indexOf('category') >= 0 || normalized.indexOf('group') >= 0) return template.category || item.groupLabel || item.typeLabel || '';
  if (normalized.indexOf('type') >= 0) return template.type || item.groupLabel || 'Config';
  if (normalized.indexOf('status') >= 0) return template.status || 'Active';
  if (normalized === 'ai_insight_enable') return template.ai_insight_enable || 'no';
  if (normalized === 'ai_insight_column') return template.ai_insight_column || '';
  if (normalized.indexOf('sort') >= 0 || normalized.indexOf('order') >= 0) return template.sortOrder || ordinal * 10;
  if (normalized.indexOf('description') >= 0 || normalized.indexOf('detail') >= 0 || normalized.indexOf('summary') >= 0 || normalized.indexOf('note') >= 0) {
    return template.detail || template.summary || `Initial ${item.label} configuration.`;
  }
  if (normalized.indexOf('color') >= 0 || normalized.indexOf('accent') >= 0) return template.color || '#c89a5b';
  if (normalized === '_' || normalized === 'blank') return '';
  return '';
}

function seedTemplateValue_(template, normalized) {
  const aliases = {
    config_name: ['key', 'label', 'name', 'title', 'item'],
    config_value: ['value', 'setting_value', 'config'],
    parent_id: ['parent', 'parent_nav_id', 'parent_page_id'],
    page: ['target_page'],
    nav_id: ['nav', 'navigation_id', 'nav_item_id'],
    view_id: ['view', 'target_view'],
    page_type: ['type', 'page_kind'],
    table_name: ['table', 'source_table', 'target_table'],
    default_view: ['view_type', 'view', 'view_mode', 'display_type'],
    field: ['column', 'column_name', 'field_name'],
    view_type: ['view', 'view_mode', 'display_type'],
    component_type: ['component', 'component_kind'],
    enabled: ['active', 'visible'],
    service_type: ['service', 'service_name'],
  };
  const candidates = [normalized].concat(aliases[normalized] || []);
  for (let index = 0; index < candidates.length; index += 1) {
    const key = candidates[index];
    if (Object.prototype.hasOwnProperty.call(template, key) && template[key] !== '') return template[key];
  }
  return '';
}

function normalizeConfigSeedHeader_(value) {
  return String(value || '')
    .replace(/^[^A-Za-z0-9\u00C0-\u1EF9_]+/, '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9\u00C0-\u1EF9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function getConfigSeedTemplates_() {
  return {
    projectInfo: [
      seed_('project_context', 'Project context', 'General', 'Business context and project starter information.'),
    ],
    step: [
      seed_('step_template', 'Template step', 'General', 'Starter workflow step used by navigation and process setup.'),
    ],
    metric: [
      Object.assign(seed_('total_records', 'Total records', 'Report', 'Count records from a configured data table.'), {
        metric_type: 'row_count',
        table_name: 'Table index',
        field: '',
      }),
    ],
    reportPlanning: [
      Object.assign(seed_('starter_overview', 'Starter overview', 'Report', 'Initial report page with cards, groups, and table preview.'), {
        item_type: 'page',
        metric_type: 'row_count',
        table_name: 'Table index',
      }),
    ],
    tableIndex: [
      Object.assign(seed_('template_table', 'Template table', 'Data', 'Starter app data object registered for User app navigation.'), {
        context: 'process',
        data_type: 'transaction',
        table_name: 'Template table',
      }),
    ],
    dataColumn: [
      Object.assign(seed_('template_category', 'Category', 'Data', 'Primary category field for the starter template table.'), {
        table_name: 'Template table',
        field: 'Category',
        label: 'Category',
        type: 'Text',
        groupable: 'TRUE',
        default_group: 'TRUE',
        group_order: 10,
      }),
      Object.assign(seed_('template_item', 'Item', 'Data', 'Primary item/name field for the starter template table.'), {
        table_name: 'Template table',
        field: 'Item',
        label: 'Item',
        type: 'Text',
        groupable: 'FALSE',
        sortOrder: 20,
      }),
      Object.assign(seed_('template_status', 'Status', 'Data', 'Starter status field for filtering and reporting.'), {
        table_name: 'Template table',
        field: 'Status',
        label: 'Status',
        type: 'Dropdown',
        groupable: 'TRUE',
        sortOrder: 30,
      }),
    ],
    relationship: [
      Object.assign(seed_('template_parent_child', 'Template relationship', 'Data', 'Starter relationship definition for parent-child table binding.'), {
        source_table: 'Template table',
        target_table: 'Template table',
        type: 'Lookup',
      }),
    ],
    appConfig: [
      Object.assign(seed_('app_name', 'App name', 'UI', 'Controls the visible app name and page title.'), {
        value: 'Process App',
        color: '#7a4e1d',
      }),
      Object.assign(seed_('accent_color', 'Accent color', 'UI', 'Primary brand accent used by controls and active states.'), {
        value: '#c89a5b',
        color: '#c89a5b',
        sortOrder: 20,
      }),
      Object.assign(seed_('default_table_view', 'Default table view', 'UI', 'Controls the default user table view mode.'), {
        value: 'table',
        sortOrder: 30,
      }),
      Object.assign(seed_('default_landing_page', 'Default landing page', 'UI', 'Controls which navigation page opens when the User app starts.'), {
        value: 'operation.group.step',
        sortOrder: 35,
      }),
      Object.assign(seed_('show_search', 'Show search', 'UI', 'Show or hide the search input in User app tables.'), {
        value: 'TRUE',
        sortOrder: 40,
      }),
      Object.assign(seed_('show_status_filter', 'Show status filter', 'UI', 'Show or hide the status filter in User app tables.'), {
        value: 'TRUE',
        sortOrder: 50,
      }),
      Object.assign(seed_('enable_grouping', 'Enable grouping', 'UI', 'Allow table grouping controls in User app.'), {
        value: 'TRUE',
        sortOrder: 60,
      }),
      Object.assign(seed_('enable_create', 'Enable create', 'Service', 'Allow new rows from the User app.'), {
        value: 'TRUE',
        sortOrder: 70,
      }),
      Object.assign(seed_('enable_edit', 'Enable edit', 'Service', 'Allow record editing from the User app.'), {
        value: 'TRUE',
        sortOrder: 80,
      }),
      Object.assign(seed_('enable_delete', 'Enable delete', 'Service', 'Allow record clearing from the User app.'), {
        value: 'TRUE',
        sortOrder: 90,
      }),
      Object.assign(seed_('enable_import', 'Enable import', 'Service', 'Allow CSV import from the User app.'), {
        value: 'TRUE',
        sortOrder: 100,
      }),
    ],
    navigation: [
      Object.assign(seed_('overview', 'Overview', 'UI', 'Root overview navigation group.'), {
        item_type: 'group',
        parent_id: '',
        icon: 'overview',
      }),
      Object.assign(seed_('operation', 'Operation', 'UI', 'Root user workflow navigation group.'), {
        item_type: 'group',
        parent_id: '',
        icon: 'operation',
        sortOrder: 20,
      }),
      Object.assign(seed_('template_table_nav', 'Template table', 'UI', 'Starter table navigation item.'), {
        item_type: 'table',
        parent_id: 'operation',
        page: 'template_page',
        icon: 'table',
        sortOrder: 30,
      }),
    ],
    page: [
      Object.assign(seed_('template_page', 'Template page', 'UI', 'Starter page config for a table-driven user view.'), {
        page_type: 'table',
        table_name: 'Template table',
        default_view: 'detail',
      }),
    ],
    view: [
      Object.assign(seed_('template_table_view', 'Table view', 'UI', 'Lean default table view with search, filter, add, and edit controls.'), {
        page_id: 'template_page',
        view_type: 'table',
      }),
      Object.assign(seed_('template_detail_view', 'Detail view', 'UI', 'Lean detail view for inspecting and editing one selected record.'), {
        page_id: 'template_page',
        view_type: 'detail',
        default_group_by: 'Category',
        sortOrder: 20,
      }),
      Object.assign(seed_('template_card_view', 'Card view', 'UI', 'Card layout for visual browsing when a table has image/title fields.'), {
        page_id: 'template_page',
        view_type: 'card',
        default_group_by: 'Category',
        sortOrder: 30,
      }),
    ],
    component: [
      Object.assign(seed_('template_table_view.search', 'Search', 'UI', 'Search control for table/list views.'), {
        view_id: 'template_table_view',
        component_type: 'search',
        enabled: 'TRUE',
      }),
      Object.assign(seed_('template_table_view.status_filter', 'Status filter', 'UI', 'Status dropdown control for table/list views.'), {
        view_id: 'template_table_view',
        component_type: 'status_filter',
        enabled: 'TRUE',
        sortOrder: 12,
      }),
      Object.assign(seed_('template_table_view.view_switcher', 'View switcher', 'UI', 'Switch table/detail view modes.'), {
        view_id: 'template_table_view',
        component_type: 'view_switcher',
        enabled: 'TRUE',
        sortOrder: 14,
      }),
      Object.assign(seed_('template_table_view.grouping', 'Grouping', 'UI', 'Column grouping controls for tables.'), {
        view_id: 'template_table_view',
        component_type: 'grouping',
        enabled: 'TRUE',
        sortOrder: 16,
      }),
      Object.assign(seed_('template_table_view.open_sheet', 'Open sheet', 'UI', 'Open the source Google Sheet table.'), {
        view_id: 'template_table_view',
        component_type: 'open_sheet',
        enabled: 'TRUE',
        sortOrder: 18,
      }),
      Object.assign(seed_('template_table_view.new_row', 'New row', 'UI', 'Create a new row in the active table.'), {
        view_id: 'template_table_view',
        component_type: 'new_row',
        enabled: 'TRUE',
        sortOrder: 20,
      }),
      Object.assign(seed_('template_detail_view.master_list', 'Master list', 'UI', 'List pane for master-detail browsing.'), {
        view_id: 'template_detail_view',
        component_type: 'master_list',
        enabled: 'TRUE',
        sortOrder: 22,
      }),
      Object.assign(seed_('template_detail_view.grouping', 'Grouping', 'UI', 'Column grouping controls for detail lists.'), {
        view_id: 'template_detail_view',
        component_type: 'grouping',
        enabled: 'TRUE',
        sortOrder: 24,
      }),
      Object.assign(seed_('template_detail_view.edit', 'Edit', 'UI', 'Edit action for selected record detail.'), {
        view_id: 'template_detail_view',
        component_type: 'edit',
        enabled: 'TRUE',
        sortOrder: 26,
      }),
      Object.assign(seed_('template_card_view.card_image', 'Card image', 'UI', 'Image area for card browsing.'), {
        view_id: 'template_card_view',
        component_type: 'card_image',
        enabled: 'TRUE',
        sortOrder: 28,
      }),
      Object.assign(seed_('template_card_view.card_fields', 'Card fields', 'UI', 'Compact field summary for card browsing.'), {
        view_id: 'template_card_view',
        component_type: 'card_fields',
        enabled: 'TRUE',
        sortOrder: 30,
      }),
      Object.assign(seed_('template_card_view.edit', 'Edit', 'UI', 'Edit action for card records.'), {
        view_id: 'template_card_view',
        component_type: 'edit',
        enabled: 'TRUE',
        sortOrder: 32,
      }),
      Object.assign(seed_('template_form_view.save', 'Save', 'UI', 'Primary save action for data forms.'), {
        view_id: 'template_form_view',
        component_type: 'save',
        enabled: 'TRUE',
        sortOrder: 40,
      }),
      Object.assign(seed_('template_form_view.delete', 'Delete', 'UI', 'Delete/clear action for data forms.'), {
        view_id: 'template_form_view',
        component_type: 'delete',
        enabled: 'TRUE',
        sortOrder: 40,
      }),
      Object.assign(seed_('template_form_view.import', 'Import', 'UI', 'CSV import controls for data tables.'), {
        view_id: 'template_form_view',
        component_type: 'import',
        enabled: 'TRUE',
        sortOrder: 50,
      }),
    ],
    dataServices: [
      Object.assign(seed_('table_crud_service', 'Table CRUD service', 'Service', 'Reads, saves, imports, and clears records through native Google Sheets Tables.'), {
        service_type: 'crud',
        table_name: 'Template table',
      }),
      Object.assign(seed_('config_registry_service', 'Config registry service', 'Service', 'Discovers config tables and turns them into app behavior.'), {
        service_type: 'config',
        table_name: 'Table index',
        sortOrder: 20,
      }),
    ],
    appPermission: [
      {
        key: 'starter_admin',
        label: 'Starter admin',
        email: '',
        role: 'Admin',
        status: 'Draft',
        detail: 'Fill Email and set Status to Active to restrict Config app access to listed users.',
        summary: 'Fill Email and set Status to Active to restrict Config app access to listed users.',
      },
    ],
  };
}

function seed_(key, label, category, detail) {
  return {
    key,
    label,
    english: label,
    category,
    type: category,
    status: 'Active',
    sortOrder: 10,
    detail,
    summary: detail,
    color: '#c89a5b',
  };
}
