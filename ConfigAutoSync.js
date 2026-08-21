/**
 * Runtime config sync.
 *
 * Runs when Config app opens. It appends missing config rows derived from
 * live schema, Navigation, Page, and View rows, and clears stale generated
 * Page/View/Form/Component rows. It never creates native tables and keeps
 * custom View/Form/Component rows that do not match auto-generated IDs.
 */
const CONFIG_AUTO_VIEW_TEMPLATES = [
  { type: 'table', label: 'Table', sortOrder: 10 },
  { type: 'detail', label: 'Detail', sortOrder: 20 },
  { type: 'card', label: 'Card', sortOrder: 30 },
];

const CONFIG_AUTO_COMPONENT_TEMPLATES = {
  table: [
    ['open_sheet', 'Open sheet'],
    ['new_row', 'New row'],
    ['view_switcher', 'View switcher'],
    ['sort', 'Sort'],
    ['search', 'Search'],
    ['status_filter', 'Status filter'],
    ['grouping', 'Grouping'],
    ['edit', 'Edit'],
  ],
  detail: [
    ['open_sheet', 'Open sheet'],
    ['new_row', 'New row'],
    ['view_switcher', 'View switcher'],
    ['sort', 'Sort'],
    ['master_list', 'Master list'],
    ['grouping', 'Grouping'],
    ['edit', 'Edit'],
  ],
  card: [
    ['open_sheet', 'Open sheet'],
    ['new_row', 'New row'],
    ['view_switcher', 'View switcher'],
    ['sort', 'Sort'],
    ['search', 'Search'],
    ['status_filter', 'Status filter'],
    ['grouping', 'Grouping'],
    ['card_image', 'Card image'],
    ['card_fields', 'Card fields'],
    ['edit', 'Edit'],
  ],
};

function syncConfigAppGeneratedRows_() {
  const syncCacheKey = cacheKey_('registry', 'configAutoSync');
  const cachedSync = getCached_(syncCacheKey);
  if (cachedSync) {
    return Object.assign({}, cachedSync, { cached: true });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const cachedAfterLock = getCached_(syncCacheKey);
    if (cachedAfterLock) {
      return Object.assign({}, cachedAfterLock, { cached: true });
    }

    let registry = getConfigAppRegistry_(true);
    let items = registry.items || {};
    const pageSchema = ensureConfigPageSchema_(items.page);
    if (pageSchema.addedCount) {
      registry = getConfigAppRegistry_(true);
      items = registry.items || {};
    }
    const viewSchema = ensureConfigViewColumnOrderSchema_(items.view);
    if (viewSchema.addedCount) {
      registry = getConfigAppRegistry_(true);
      items = registry.items || {};
    }
    const formSchema = ensureConfigFormSchema_(items.form);
    if (formSchema.addedCount) {
      registry = getConfigAppRegistry_(true);
      items = registry.items || {};
    }
    const relationshipSchema = ensureConfigRelationshipSchema_(items.relationship);
    if (relationshipSchema.addedCount) {
      registry = getConfigAppRegistry_(true);
      items = registry.items || {};
    }
    const result = {
      viewSchema,
      pageSchema,
      formSchema,
      relationshipSchema,
      dataColumns: { addedCount: 0, skipped: [] },
      relationships: { addedCount: 0, skipped: [] },
      pages: syncConfigPagesFromNavigation_(items),
      views: { addedCount: 0, skipped: [] },
      forms: { addedCount: 0, skipped: [] },
      components: { addedCount: 0, skipped: [] },
    };

    result.dataColumns = syncConfigDataColumnsFromTables_(items);
    result.views = syncConfigViewsFromPages_(items);
    result.forms = syncConfigFormsFromPages_(items);
    result.relationships = syncConfigRelationshipsFromForms_(items);
    result.components = syncConfigComponentsFromViews_(items);

    result.deletedCount = ['dataColumns', 'relationships', 'pages', 'views', 'forms', 'components']
      .reduce((sum, key) => sum + Number((result[key] && result[key].deletedCount) || 0), 0);

    if (result.pageSchema.addedCount || result.viewSchema.addedCount || result.formSchema.addedCount || result.relationshipSchema.addedCount || result.dataColumns.addedCount || result.relationships.addedCount || result.pages.addedCount || result.views.addedCount || result.forms.addedCount || result.components.addedCount || result.deletedCount) {
      invalidateDataCache_();
    }
    result.addedCount = result.pageSchema.addedCount + result.viewSchema.addedCount + result.formSchema.addedCount + result.relationshipSchema.addedCount + result.dataColumns.addedCount + result.relationships.addedCount + result.pages.addedCount + result.views.addedCount + result.forms.addedCount + result.components.addedCount;
    result.generatedAt = new Date().toISOString();
    setCached_(syncCacheKey, result, 90);
    return result;
  } finally {
    lock.releaseLock();
  }
}

function ensureConfigPageSchema_(pageItem) {
  try {
    const requiredHeaders = ['Title column', 'Data mode', 'AI insight enable', 'AI insight column'];
    if (!pageItem || !pageItem.table) return { addedCount: 0, skipped: ['missing_page_table'] };
    const table = pageItem.table;
    const headers = getConfigAutoHeadersForTable_(table);
    const existing = headers.reduce((acc, header) => {
      acc[normalizeConfigSeedHeader_(header)] = true;
      return acc;
    }, {});
    const missing = requiredHeaders.filter((header) => !existing[normalizeConfigSeedHeader_(header)]);
    if (!missing.length) return { addedCount: 0, skipped: [] };

    if (isNativeReadableTable_(table)) {
      expandNativeConfigTableColumns_(table, headers.length, missing);
    } else {
      const sheet = getSheet_(table);
      const layout = getSystemTableLayout_(sheet, table);
      sheet.getRange(layout.headerRow, layout.dataStartColumn + headers.length, 1, missing.length).setValues([missing]);
      SpreadsheetApp.flush();
    }
    removeConfigSchemaCache_();
    return { addedCount: missing.length, addedHeaders: missing };
  } catch (error) {
    return { addedCount: 0, skipped: [error && error.message ? error.message : String(error)] };
  }
}

function ensureConfigRelationshipSchema_(relationshipItem) {
  try {
    const headers = ['Relationship ID', 'Item', 'Source table', 'Source field', 'Target table', 'Target field', 'Relationship type', 'Status'];
    if (relationshipItem && relationshipItem.dataConfig) return { addedCount: 0, skipped: ['data_config_schema_owned_by_source_file'] };
    if (!relationshipItem || !relationshipItem.table) return { addedCount: 0, skipped: ['missing_relationship_table'] };
    const table = relationshipItem.table;
    const current = getConfigAutoHeadersForTable_(table);
    const generic = current.length && current.every((header) => /^column\s*\d+$/i.test(String(header || '').trim()));
    if (generic || !current.length) {
      writeConfigAutoHeaderRow_(table, headers);
      removeConfigSchemaCache_();
      return { addedCount: headers.length, addedHeaders: headers };
    }

    const existing = current.reduce((acc, header) => {
      acc[normalizeConfigSeedHeader_(header)] = true;
      return acc;
    }, {});
    const missing = headers.filter((header) => !existing[normalizeConfigSeedHeader_(header)]);
    if (!missing.length) return { addedCount: 0, skipped: [] };
    appendConfigAutoHeaders_(table, current.length, missing);
    removeConfigSchemaCache_();
    return { addedCount: missing.length, addedHeaders: missing };
  } catch (error) {
    return { addedCount: 0, skipped: [error && error.message ? error.message : String(error)] };
  }
}

function ensureConfigFormSchema_(formItem) {
  try {
    const headers = ['Form ID', 'Page ID', 'Item', 'Mode', 'Form type', 'Columns', 'Label position', 'Show empty fields', 'Main display name', 'Column order', 'Hidden columns', 'Section layout', 'Nested table', 'Status'];
    if (!formItem || !formItem.table) return { addedCount: 0, skipped: ['missing_form_table'] };
    const table = formItem.table;
    const current = getConfigAutoHeadersForTable_(table);
    const generic = current.length && current.every((header) => /^column\s*\d+$/i.test(String(header || '').trim()));
    if (generic || !current.length) {
      writeConfigAutoHeaderRow_(table, headers);
      migrateConfigFormTypeValues_(table, headers);
      applyConfigFormValidation_(table, headers);
      removeConfigSchemaCache_();
      return { addedCount: headers.length, addedHeaders: headers };
    }

    const existing = current.reduce((acc, header) => {
      acc[normalizeConfigSeedHeader_(header)] = true;
      return acc;
    }, {});
    const missing = headers.filter((header) => !existing[normalizeConfigSeedHeader_(header)]);
    if (!missing.length) {
      migrateConfigFormTypeValues_(table, current);
      applyConfigFormValidation_(table, current);
      return { addedCount: 0, skipped: [] };
    }
    appendConfigAutoHeaders_(table, current.length, missing);
    const nextHeaders = current.concat(missing);
    migrateConfigFormTypeValues_(table, nextHeaders);
    applyConfigFormValidation_(table, nextHeaders);
    removeConfigSchemaCache_();
    return { addedCount: missing.length, addedHeaders: missing };
  } catch (error) {
    return { addedCount: 0, skipped: [error && error.message ? error.message : String(error)] };
  }
}

function syncFormConfigSchema_() {
  const registry = getConfigAppRegistry_(true);
  const items = registry.items || {};
  const formSchema = ensureConfigFormSchema_(items.form);
  if (formSchema.addedCount) invalidateDataCache_();
  return Object.assign({ ok: true, updatedAt: new Date().toISOString() }, formSchema);
}

function applyConfigFormValidation_(table, headers) {
  applyConfigAutoListValidation_(table, headers, ['form_type'], ['default', 'multi-step'], true);
  applyConfigAutoListValidation_(table, headers, ['mode', 'form_mode', 'edit_mode'], ['quick', 'full', 'nested'], true);
  applyConfigAutoListValidation_(table, headers, ['status'], ['Active', 'Draft', 'Inactive'], true);
}

function migrateConfigFormTypeValues_(table, headers) {
  const typeIndex = findConfigAutoHeaderIndex_(headers, ['form_type']);
  const modeIndex = findConfigAutoHeaderIndex_(headers, ['mode', 'form_mode', 'edit_mode']);
  if (typeIndex < 0 || modeIndex < 0) return { changedCount: 0 };
  const target = getConfigAutoValidationTarget_(table);
  if (!target || !target.sheet) return { changedCount: 0 };
  const rowCount = Math.max(0, Math.min(999, target.sheet.getLastRow() - target.headerRow));
  if (!rowCount) return { changedCount: 0 };
  const range = target.sheet.getRange(target.headerRow + 1, target.dataStartColumn, rowCount, headers.length);
  const values = range.getValues();
  let changedCount = 0;
  values.forEach((row) => {
    const rawType = row[typeIndex];
    const typeKey = normalizeUserConfigKey_(rawType);
    const isLegacyMode = typeKey === 'quick' || typeKey === 'full' || typeKey === 'advanced';
    if (!isLegacyMode) return;
    if (!String(row[modeIndex] || '').trim()) row[modeIndex] = typeKey === 'advanced' ? 'full' : typeKey;
    row[typeIndex] = 'default';
    changedCount += 1;
  });
  if (changedCount) {
    range.setValues(values);
    SpreadsheetApp.flush();
  }
  return { changedCount };
}

function findConfigAutoHeaderIndex_(headers, aliases) {
  return (headers || []).findIndex((header) => {
    const normalized = normalizeConfigSeedHeader_(header);
    return (aliases || []).some((alias) => normalizeConfigSeedHeader_(alias) === normalized);
  });
}

function ensureConfigViewColumnOrderSchema_(viewItem) {
  try {
    const requiredHeaders = ['Title column', 'Column order', 'Hidden columns', 'Section layout', 'Filter fields', 'Default filters', 'Default sort', 'Default sort direction'];
    if (!viewItem || !viewItem.table) return { addedCount: 0, skipped: ['missing_view_table'] };
    const table = viewItem.table;
    const headers = getConfigAutoHeadersForTable_(table);
    const existing = headers.reduce((acc, header) => {
      acc[normalizeConfigSeedHeader_(header)] = true;
      return acc;
    }, {});
    const missing = requiredHeaders.filter((header) => !existing[normalizeConfigSeedHeader_(header)]);
    if (!missing.length) return { addedCount: 0, skipped: [] };

    if (isNativeReadableTable_(table)) {
      expandNativeConfigTableColumns_(table, headers.length, missing);
    } else {
      const sheet = getSheet_(table);
      const layout = getSystemTableLayout_(sheet, table);
      sheet.getRange(layout.headerRow, layout.dataStartColumn + headers.length, 1, missing.length).setValues([missing]);
      SpreadsheetApp.flush();
    }
    removeConfigSchemaCache_();
    return { addedCount: missing.length, addedHeaders: missing };
  } catch (error) {
    return { addedCount: 0, skipped: [error && error.message ? error.message : String(error)] };
  }
}

function expandNativeConfigTableColumns_(table, currentWidth, headersToAppend) {
  const range = Object.assign({}, table.apiRange || {});
  const startColumnIndex = Number(range.startColumnIndex || 0);
  const endColumnIndex = startColumnIndex + currentWidth + headersToAppend.length;
  const sheet = getSheet_(table);
  if (sheet.getMaxColumns() < endColumnIndex) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), endColumnIndex - sheet.getMaxColumns());
  }

  const nextRange = Object.assign({}, range, { endColumnIndex });
  sheetsApiBatchUpdate_(getTableSpreadsheetId_(table), [
    {
      updateTable: {
        table: {
          tableId: table.apiTableId,
          range: nextRange,
        },
        fields: 'range',
      },
    },
    {
      updateCells: {
        range: {
          sheetId: getNativeSheetId_(table),
          startRowIndex: Number(range.startRowIndex || 0),
          endRowIndex: Number(range.startRowIndex || 0) + 1,
          startColumnIndex: startColumnIndex + currentWidth,
          endColumnIndex,
        },
        rows: [{
          values: headersToAppend.map((header) => ({ userEnteredValue: { stringValue: header } })),
        }],
        fields: 'userEnteredValue',
      },
    },
  ]);
}

function writeConfigAutoHeaderRow_(table, headers) {
  if (isNativeReadableTable_(table)) {
    const width = getNativeTableWidth_(table);
    if (width < headers.length) appendConfigAutoHeaders_(table, width, headers.slice(width));
    const range = table.apiRange || {};
    sheetsApiBatchUpdate_(getTableSpreadsheetId_(table), [{
      updateCells: {
        range: {
          sheetId: getNativeSheetId_(table),
          startRowIndex: Number(range.startRowIndex || 0),
          endRowIndex: Number(range.startRowIndex || 0) + 1,
          startColumnIndex: Number(range.startColumnIndex || 0),
          endColumnIndex: Number(range.startColumnIndex || 0) + headers.length,
        },
        rows: [{
          values: headers.map((header) => ({ userEnteredValue: { stringValue: header } })),
        }],
        fields: 'userEnteredValue',
      },
    }]);
    return;
  }
  const sheet = getSheet_(table);
  const layout = getSystemTableLayout_(sheet, table);
  sheet.getRange(layout.headerRow, layout.dataStartColumn, 1, headers.length).setValues([headers]);
  SpreadsheetApp.flush();
}

function appendConfigAutoHeaders_(table, currentWidth, headersToAppend) {
  if (!headersToAppend || !headersToAppend.length) return;
  if (isNativeReadableTable_(table)) {
    expandNativeConfigTableColumns_(table, currentWidth, headersToAppend);
    return;
  }
  const sheet = getSheet_(table);
  const layout = getSystemTableLayout_(sheet, table);
  sheet.getRange(layout.headerRow, layout.dataStartColumn + currentWidth, 1, headersToAppend.length).setValues([headersToAppend]);
  SpreadsheetApp.flush();
}

function applyConfigAutoListValidation_(table, headers, aliases, values, strict) {
  const headerIndex = findConfigAutoHeaderIndex_(headers, aliases);
  if (headerIndex < 0) return;
  const target = getConfigAutoValidationTarget_(table);
  if (!target || !target.sheet) return;
  const range = target.sheet.getRange(target.headerRow + 1, target.dataStartColumn + headerIndex, 999, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(!strict)
    .build();
  range.setDataValidation(rule);
}

function getConfigAutoValidationTarget_(table) {
  try {
    if (isNativeReadableTable_(table)) {
      const spreadsheet = SpreadsheetApp.openById(getTableSpreadsheetId_(table));
      const sheetId = getNativeSheetId_(table);
      const sheet = spreadsheet.getSheets().find((item) => Number(item.getSheetId()) === Number(sheetId));
      const range = table.apiRange || {};
      return sheet ? {
        sheet,
        headerRow: Number(range.startRowIndex || 0) + 1,
        dataStartColumn: Number(range.startColumnIndex || 0) + 1,
      } : null;
    }
    const sheet = getSheet_(table);
    const layout = getSystemTableLayout_(sheet, table);
    return {
      sheet,
      headerRow: layout.headerRow,
      dataStartColumn: layout.dataStartColumn,
    };
  } catch (error) {
    console.warn('Config validation skipped:', error);
    return null;
  }
}

function removeConfigSchemaCache_() {
  try {
    CacheService.getScriptCache().removeAll([
      cacheKey_('registry', 'tables'),
      cacheKey_('registry', 'nativeTables'),
      cacheKey_('registry', `nativeTables:${SPREADSHEET_ID}`),
      cacheKey_('registry', 'userAppConfig'),
    ]);
  } catch (error) {
    // Best-effort cache cleanup after table schema change.
  }
}

function syncConfigDataColumnsFromTables_(items) {
  const dataColumnItem = items.dataColumn;
  if (dataColumnItem && dataColumnItem.dataConfig) {
    return { addedCount: 0, skipped: ['data_config_columns_owned_by_source_file'] };
  }
  if (!dataColumnItem || !dataColumnItem.table) {
    return { addedCount: 0, skipped: ['missing_data_column_table'] };
  }

  const data = buildTableData_(dataColumnItem.table, true);
  const headers = data.headers || [];
  const records = getConfigAutoDataTables_().flatMap((table) => (
    getConfigAutoDataColumnRecordsForTable_(headers, table)
  ));
  const cleanup = cleanupStaleDataColumnRows_(dataColumnItem, data, records);
  const append = appendMissingDataColumnRows_(dataColumnItem, data, records);
  return combineConfigAutoSyncResult_(append, cleanup);
}

function syncConfigRelationshipsFromForms_(items) {
  const formItem = items.form;
  const pageItem = items.page;
  const relationshipItem = items.relationship;
  if (relationshipItem && relationshipItem.dataConfig) {
    return { addedCount: 0, skipped: ['data_config_relationships_saved_on_form_submit'] };
  }
  if (!formItem || !formItem.table || !pageItem || !pageItem.table || !relationshipItem || !relationshipItem.table) {
    return { addedCount: 0, skipped: ['missing_form_page_or_relationship_table'] };
  }

  const formData = buildTableData_(formItem.table, true);
  const pageData = buildTableData_(pageItem.table, true);
  const relationshipData = buildTableData_(relationshipItem.table, true);
  const records = getConfigAutoRelationshipRecordsFromForms_(relationshipData.headers || [], formData.rows || [], pageData.rows || []);
  const cleanup = cleanupStaleRelationshipRows_(relationshipItem, relationshipData);
  const append = appendMissingConfigRows_(relationshipItem, relationshipData, ['relationship_id', 'id', 'key', 'config_key'], records);
  return combineConfigAutoSyncResult_(append, cleanup);
}

function syncConfigPagesFromNavigation_(items) {
  const navItem = items.navigation;
  const pageItem = items.page;
  if (!navItem || !navItem.table || !pageItem || !pageItem.table) {
    return { addedCount: 0, skipped: ['missing_navigation_or_page_table'] };
  }

  const navData = buildTableData_(navItem.table, true);
  const pageData = buildTableData_(pageItem.table, true);
  const userTables = getConfigAutoUserTablesByName_();
  const records = getConfigAutoNavPageRequests_(navData.rows || [])
    .map((request) => createConfigAutoPageRecord_(pageData.headers || [], request, userTables));
  const cleanup = cleanupStaleConfigPageRows_(pageItem, pageData, records);
  const append = appendMissingConfigRows_(pageItem, pageData, ['page_id', 'page', 'id', 'key', 'config_key'], records);
  return combineConfigAutoSyncResult_(append, cleanup);
}

function getConfigAutoDataTables_() {
  const tables = getTables_(true);
  return Object.keys(tables)
    .map((key) => tables[key])
    .filter((table) => table && table.key && table.name)
    .filter((table) => table.type !== 'config' && table.type !== 'system')
    .filter((table) => table.binding || table.apiTableId || table.sheetName);
}

function getConfigAutoDataColumnRecordsForTable_(configHeaders, table) {
  const headers = getConfigAutoHeadersForTable_(table);
  if (!headers.length) return [];
  const validations = getConfigAutoValidationByHeader_(table, headers);
  return headers.map((header, index) => createConfigAutoDataColumnRecord_(configHeaders, table, header, index, validations[header] || null));
}

function syncConfigViewsFromPages_(items) {
  const pageItem = items.page;
  const viewItem = items.view;
  if (!pageItem || !pageItem.table || !viewItem || !viewItem.table) {
    return { addedCount: 0, skipped: ['missing_page_or_view_table'] };
  }

  const pageData = buildTableData_(pageItem.table, true);
  const viewData = buildTableData_(viewItem.table, true);
  const records = getConfigAutoTablePageRequests_(pageData.rows || [])
    .flatMap((page) => CONFIG_AUTO_VIEW_TEMPLATES.map((template) => (
      createConfigAutoViewRecord_(viewData.headers || [], page, template)
    )));
  const cleanup = cleanupStaleGeneratedConfigRows_(viewItem, viewData, records, ['view_id', 'view', 'id', 'key', 'config_key'], isConfigAutoGeneratedViewRow_);
  const append = appendMissingConfigRows_(viewItem, viewData, ['view_id', 'view', 'id', 'key', 'config_key'], records);
  return combineConfigAutoSyncResult_(append, cleanup);
}

function syncConfigFormsFromPages_(items) {
  const pageItem = items.page;
  const formItem = items.form;
  if (!pageItem || !pageItem.table || !formItem || !formItem.table) {
    return { addedCount: 0, skipped: ['missing_page_or_form_table'] };
  }

  const pageData = buildTableData_(pageItem.table, true);
  const formData = buildTableData_(formItem.table, true);
  const records = getConfigAutoTablePageRequests_(pageData.rows || [])
    .flatMap((page) => ([
      createConfigAutoFormRecord_(formData.headers || [], page, { type: 'quick', label: 'Quick form', columns: 1 }),
      createConfigAutoFormRecord_(formData.headers || [], page, { type: 'full', label: 'Full form', columns: 3 }),
      createConfigAutoFormRecord_(formData.headers || [], page, { type: 'nested', label: 'Nested form', columns: 2 }),
    ]));
  const cleanup = cleanupStaleGeneratedConfigRows_(formItem, formData, records, ['form_id', 'form', 'id', 'key', 'config_key'], isConfigAutoGeneratedFormRow_);
  const append = appendMissingConfigRows_(formItem, formData, ['form_id', 'form', 'id', 'key', 'config_key'], records);
  return combineConfigAutoSyncResult_(append, cleanup);
}

function syncConfigComponentsFromViews_(items) {
  const viewItem = items.view;
  const componentItem = items.component;
  if (!viewItem || !viewItem.table || !componentItem || !componentItem.table) {
    return { addedCount: 0, skipped: ['missing_view_or_component_table'] };
  }

  const viewData = buildTableData_(viewItem.table, true);
  const componentData = buildTableData_(componentItem.table, true);
  const records = getConfigAutoViewRequests_(viewData.rows || [])
    .flatMap((view) => (CONFIG_AUTO_COMPONENT_TEMPLATES[view.viewType] || [])
      .map(([componentType, label], index) => (
        createConfigAutoComponentRecord_(componentData.headers || [], view, componentType, label, index)
      )));
  const cleanup = cleanupStaleGeneratedConfigRows_(componentItem, componentData, records, ['component_id', 'component', 'id', 'key', 'config_key'], isConfigAutoGeneratedComponentRow_);
  const append = appendMissingConfigRows_(componentItem, componentData, ['component_id', 'component', 'id', 'key', 'config_key'], records);
  return combineConfigAutoSyncResult_(append, cleanup);
}

function getConfigAutoNavPageRequests_(rows) {
  return (rows || []).map((row, index) => {
    const record = row.record || {};
    if (!isConfigAutoActiveRecord_(record)) return null;

    const itemType = normalizeUserConfigKey_(recordValue_(record, ['item_type', 'navigation_type', 'type', 'component_type']));
    const navId = recordValue_(record, ['nav_id', 'navigation_id', 'id', 'key', 'config_key']);
    let pageId = recordValue_(record, ['page', 'page_id', 'target_page']);
    if (!pageId && isConfigAutoPageNavType_(itemType)) pageId = navId;
    if (!pageId || itemType === 'group') return null;

    const label = recordValue_(record, ['label', 'item', 'title', 'name']) || pageId;
    const pageType = inferConfigAutoPageType_(itemType, label, pageId);
    return {
      pageId,
      label,
      pageType,
      linkUrl: recordValue_(record, ['link_url', 'url', 'external_url', 'embed_url']),
      sortOrder: Number(recordValue_(record, ['sort_order', 'sort', 'order']) || ((index + 1) * 10)),
    };
  }).filter(Boolean);
}

function getConfigAutoTablePageRequests_(rows) {
  return (rows || []).map((row, index) => {
    const record = row.record || {};
    if (!isConfigAutoActiveRecord_(record)) return null;
    const pageId = recordValue_(record, ['page_id', 'page', 'id', 'key', 'config_key']);
    if (!pageId) return null;
    const pageType = inferConfigAutoPageType_(
      recordValue_(record, ['page_type', 'type', 'item_type']) || 'table',
      recordValue_(record, ['item', 'label', 'title', 'name']),
      pageId
    );
    if (pageType !== 'table') return null;
    return {
      pageId,
      label: recordValue_(record, ['item', 'label', 'title', 'name']) || pageId,
      tableName: recordValue_(record, ['table_name', 'table', 'source_table']),
      sortOrder: Number(recordValue_(record, ['sort_order', 'sort', 'order']) || ((index + 1) * 10)),
    };
  }).filter(Boolean);
}

function getConfigAutoViewRequests_(rows) {
  return (rows || []).map((row, index) => {
    const record = row.record || {};
    if (!isConfigAutoActiveRecord_(record)) return null;
    const viewId = recordValue_(record, ['view_id', 'view', 'id', 'key', 'config_key']);
    const viewType = normalizeUserViewType_(recordValue_(record, ['view_type', 'type', 'display_type']));
    if (!viewId || !CONFIG_AUTO_COMPONENT_TEMPLATES[viewType]) return null;
    return {
      viewId,
      viewType,
      pageId: recordValue_(record, ['page_id', 'page', 'parent_id']),
      label: recordValue_(record, ['item', 'label', 'title', 'name']) || titleCaseConfigName_(viewType),
      sortOrder: Number(recordValue_(record, ['sort_order', 'sort', 'order']) || ((index + 1) * 10)),
    };
  }).filter(Boolean);
}

function createConfigAutoPageRecord_(headers, request, userTables) {
  const record = createConfigAutoBlankRecord_(headers);
  const tableName = request.pageType === 'table' ? inferConfigAutoTableName_(request, userTables) : '';
  setConfigAutoValue_(record, headers, ['page_id', 'page', 'id', 'key', 'config_key'], request.pageId);
  setConfigAutoValue_(record, headers, ['item', 'label', 'title', 'name'], request.label);
  setConfigAutoValue_(record, headers, ['page_type', 'type', 'item_type'], request.pageType);
  setConfigAutoValue_(record, headers, ['table_name', 'table', 'source_table'], tableName);
  setConfigAutoValue_(record, headers, ['title_column', 'record_title_column', 'record_name_column', 'display_name_column'], '');
  setConfigAutoValue_(record, headers, ['default_table_view', 'default_view', 'default_view_type'], request.pageType === 'table' ? 'table' : '');
  setConfigAutoValue_(record, headers, ['report_name', 'report', 'overview_key', 'report_key'], request.pageType === 'report' ? request.pageId : '');
  setConfigAutoValue_(record, headers, ['link_url', 'url', 'embed_url', 'external_url'], request.linkUrl);
  setConfigAutoValue_(record, headers, ['sort_order', 'sort', 'order'], request.sortOrder);
  setConfigAutoValue_(record, headers, ['status'], 'Active');
  return record;
}

function createConfigAutoDataColumnRecord_(headers, table, field, index, validation) {
  const record = createConfigAutoBlankRecord_(headers);
  const inputType = inferFieldInputType_(field, validation);
  setConfigAutoValue_(record, headers, ['column_id', 'id', 'key', 'config_key'], `${table.key}.${normalizeUserConfigKey_(field)}`);
  setConfigAutoValue_(record, headers, ['category', 'config_group', 'group'], table.name);
  setConfigAutoValue_(record, headers, ['item', 'label', 'title', 'name'], field);
  setConfigAutoValue_(record, headers, ['table_name', 'table', 'source_table'], table.name);
  setConfigAutoValue_(record, headers, ['field', 'column', 'column_name', 'field_name'], field);
  setConfigAutoValue_(record, headers, ['display_label', 'label'], titleCaseConfigName_(field));
  setConfigAutoValue_(record, headers, ['input_type', 'data_type', 'type'], inputType);
  setConfigAutoValue_(record, headers, ['groupable', 'allow_group', 'enable_group'], isConfigAutoGroupableField_(field, inputType) ? 'TRUE' : 'FALSE');
  setConfigAutoValue_(record, headers, ['default_group', 'group_by_default', 'is_default_group'], 'FALSE');
  setConfigAutoValue_(record, headers, ['group_order', 'default_group_order'], '');
  setConfigAutoValue_(record, headers, ['hidden', 'hide', 'is_hidden'], 'FALSE');
  setConfigAutoValue_(record, headers, ['section'], '');
  setConfigAutoValue_(record, headers, ['sort_order', 'sort', 'order'], (index + 1) * 10);
  setConfigAutoValue_(record, headers, ['status'], 'Active');
  return record;
}

function createConfigAutoViewRecord_(headers, page, template) {
  const record = createConfigAutoBlankRecord_(headers);
  const viewId = `${page.pageId}.${template.type}`;
  setConfigAutoValue_(record, headers, ['view_id', 'view', 'id', 'key', 'config_key'], viewId);
  setConfigAutoValue_(record, headers, ['page_id', 'page', 'parent_id'], page.pageId);
  setConfigAutoValue_(record, headers, ['item', 'label', 'title', 'name'], template.label);
  setConfigAutoValue_(record, headers, ['view_type', 'type', 'display_type'], template.type);
  setConfigAutoValue_(record, headers, ['table_name', 'table', 'source_table'], page.tableName);
  setConfigAutoValue_(record, headers, ['title_column', 'record_title_column', 'record_name_column', 'display_name_column', 'display_name', 'display_field', 'main_display_name', 'title_field'], '');
  setConfigAutoValue_(record, headers, ['column_order', 'field_order', 'columns_order'], '');
  setConfigAutoValue_(record, headers, ['hidden_columns', 'hide_columns', 'hidden_fields'], '');
  setConfigAutoValue_(record, headers, ['section_layout', 'form_sections', 'sections'], '');
  setConfigAutoValue_(record, headers, ['default_sort', 'default_sort_field', 'sort_field'], '');
  setConfigAutoValue_(record, headers, ['default_sort_direction', 'sort_direction'], '');
  setConfigAutoValue_(record, headers, ['sort_order', 'sort', 'order'], template.sortOrder);
  setConfigAutoValue_(record, headers, ['status'], 'Active');
  return record;
}

function createConfigAutoFormRecord_(headers, page, template) {
  const record = createConfigAutoBlankRecord_(headers);
  const formId = `${page.pageId}.${template.type}_form`;
  setConfigAutoValue_(record, headers, ['form_id', 'form', 'id', 'key', 'config_key'], formId);
  setConfigAutoValue_(record, headers, ['page_id', 'page', 'parent_id'], page.pageId);
  setConfigAutoValue_(record, headers, ['item', 'label', 'title', 'name'], template.label);
  setConfigAutoValue_(record, headers, ['mode', 'form_mode', 'edit_mode'], template.type);
  setConfigAutoValue_(record, headers, ['form_type', 'form_layout_type', 'layout_type'], template.formType || 'default');
  setConfigAutoValue_(record, headers, ['columns', 'layout_columns', 'column_count'], template.columns);
  setConfigAutoValue_(record, headers, ['label_position', 'label_layout'], 'top');
  setConfigAutoValue_(record, headers, ['show_empty_fields', 'show_empty'], 'TRUE');
  setConfigAutoValue_(record, headers, ['column_order', 'field_order'], '');
  setConfigAutoValue_(record, headers, ['hidden_columns', 'hide_columns', 'hidden_fields'], '');
  setConfigAutoValue_(record, headers, ['section_layout', 'form_sections', 'sections'], '');
  setConfigAutoValue_(record, headers, ['nested_table', 'nested_tables', 'related_table'], '');
  setConfigAutoValue_(record, headers, ['main_display_name', 'main_display_field', 'display_name', 'display_field', 'title_field'], '');
  setConfigAutoValue_(record, headers, ['status'], 'Active');
  return record;
}

function appendMissingDataColumnRows_(item, data, records) {
  const headers = data.headers || [];
  if (!item || !item.table || !headers.length || !records.length) {
    return { addedCount: 0, skipped: [] };
  }

  const existing = (data.rows || []).reduce((acc, row) => {
    const record = row.record || {};
    const id = normalizeConfigAutoId_(recordValue_(record, ['column_id', 'id', 'key', 'config_key']));
    const pair = getConfigAutoDataColumnPairKey_(record);
    if (id) acc.ids[id] = true;
    if (pair) acc.pairs[pair] = true;
    return acc;
  }, { ids: {}, pairs: {} });

  const added = { ids: {}, pairs: {} };
  const missing = records.filter((record) => {
    const id = normalizeConfigAutoId_(recordValue_(record, ['column_id', 'id', 'key', 'config_key']));
    const pair = getConfigAutoDataColumnPairKey_(record);
    if ((!id && !pair) || existing.ids[id] || existing.pairs[pair] || added.ids[id] || added.pairs[pair]) return false;
    if (id) added.ids[id] = true;
    if (pair) added.pairs[pair] = true;
    return true;
  });

  if (!missing.length) return { addedCount: 0, skipped: [] };
  const addedCount = isNativeReadableTable_(item.table)
    ? importNativeRecords_(item.table, missing)
    : importSystemRecords_(item.table, missing);
  return {
    addedCount,
    addedKeys: missing.map((record) => getConfigAutoDataColumnPairKey_(record)).filter(Boolean),
  };
}

function getConfigAutoDataColumnPairKey_(record) {
  const tableName = recordValue_(record, ['table_name', 'table', 'source_table']);
  const field = recordValue_(record, ['field', 'column', 'column_name', 'field_name', 'item', 'label']);
  return tableName && field ? `${normalizeUserConfigKey_(tableName)}::${normalizeUserConfigKey_(field)}` : '';
}

function cleanupStaleDataColumnRows_(item, data, expectedRecords) {
  const expected = (expectedRecords || []).reduce((acc, record) => {
    const pair = getConfigAutoDataColumnPairKey_(record);
    if (pair) acc[pair] = true;
    return acc;
  }, {});
  const staleRows = ((data && data.rows) || []).filter((row) => {
    const record = row.record || {};
    if (isConfigAutoManualKeepRow_(record)) return false;
    const pair = getConfigAutoDataColumnPairKey_(record);
    return !!pair && !expected[pair];
  });
  return clearConfigAutoRows_(item, staleRows, ['column_id', 'id', 'key', 'config_key']);
}

function getConfigAutoRelationshipRecordsFromForms_(headers, formRows, pageRows) {
  const tableLookup = getConfigAutoDataTableLookup_();
  const tablePages = getConfigAutoTablePageRequests_(pageRows || []).reduce((acc, page) => {
    acc[normalizeConfigAutoId_(page.pageId)] = page;
    return acc;
  }, {});
  const seen = {};

  return (formRows || []).flatMap((row) => {
    const record = row.record || {};
    if (!isConfigAutoActiveRecord_(record)) return [];

    const formId = recordValue_(record, ['form_id', 'form', 'id', 'key', 'config_key']);
    const pageId = recordValue_(record, ['page_id', 'page', 'parent_id']) || inferPageIdFromFormId_(formId);
    const page = tablePages[normalizeConfigAutoId_(pageId)] || {};
    const sourceTable = resolveConfigAutoDataTable_(
      recordValue_(record, ['table_name', 'table', 'source_table']) || page.tableName,
      tableLookup
    );
    if (!sourceTable) return [];

    const sourceHeaders = getConfigAutoHeadersForTable_(sourceTable);
    const nestedTables = parseUserConfigList_(recordValue_(record, ['nested_table', 'nested_tables', 'related_table', 'related_tables']));
    return nestedTables.map((targetName) => {
      const targetTable = resolveConfigAutoDataTable_(targetName, tableLookup);
      if (!targetTable) return null;

      const targetHeaders = getConfigAutoHeadersForTable_(targetTable);
      const sourceField = guessConfigAutoRelationshipField_(sourceHeaders, targetHeaders, true);
      const targetField = guessConfigAutoRelationshipField_(targetHeaders, sourceHeaders, false);
      if (!sourceField || !targetField) return null;

      const key = [
        sourceTable.name,
        targetTable.name,
        sourceField,
        targetField,
      ].map(normalizeConfigAutoId_).join('::');
      if (seen[key]) return null;
      seen[key] = true;

      return createRelationshipConfigRecord_(headers, {
        sourceTable: sourceTable.name,
        targetTable: targetTable.name,
        sourceField,
        targetField,
      });
    }).filter(Boolean);
  });
}

function cleanupStaleRelationshipRows_(item, data) {
  const tableLookup = getConfigAutoDataTableLookup_();
  const staleRows = ((data && data.rows) || []).filter((row) => {
    const record = row.record || {};
    if (isConfigAutoManualKeepRow_(record)) return false;

    const sourceName = recordValue_(record, ['source_table', 'parent_table', 'from_table', 'table_name', 'table']);
    const targetName = recordValue_(record, ['target_table', 'child_table', 'to_table', 'related_table']);
    const sourceField = recordValue_(record, ['source_field', 'parent_field', 'from_field', 'key_field', 'field']);
    const targetField = recordValue_(record, ['target_field', 'child_field', 'to_field', 'lookup_field', 'related_field']);
    if (!sourceName && !targetName && !sourceField && !targetField) return false;

    const sourceTable = resolveConfigAutoDataTable_(sourceName, tableLookup);
    const targetTable = resolveConfigAutoDataTable_(targetName, tableLookup);
    if (!sourceTable || !targetTable) return true;
    return !hasConfigAutoTableField_(sourceTable, sourceField) || !hasConfigAutoTableField_(targetTable, targetField);
  });
  return clearConfigAutoRows_(item, staleRows, ['relationship_id', 'id', 'key', 'config_key']);
}

function getConfigAutoDataTableLookup_() {
  return getConfigAutoDataTables_().reduce((acc, table) => {
    getConfigAutoTableLookupNames_(table).forEach((name) => {
      if (name) acc[name] = table;
    });
    return acc;
  }, {});
}

function getConfigAutoTableLookupNames_(table) {
  const names = [
    table && table.name,
    table && table.key,
    table && table.apiTableName,
    table && table.sheetName,
  ].concat((table && table.aliases) || []);
  return names
    .flatMap((name) => [name, stripTableLabelIcon_(name)])
    .map(normalizeUserConfigKey_)
    .filter(Boolean);
}

function resolveConfigAutoDataTable_(tableName, tableLookup) {
  return tableLookup[normalizeUserConfigKey_(tableName)] || tableLookup[normalizeUserConfigKey_(stripTableLabelIcon_(tableName))] || null;
}

function guessConfigAutoRelationshipField_(primaryColumns, otherColumns, preferId) {
  const otherKeys = (otherColumns || []).reduce((acc, column) => {
    acc[normalizeUserConfigKey_(column)] = true;
    return acc;
  }, {});
  const exact = (primaryColumns || []).find((column) => otherKeys[normalizeUserConfigKey_(column)]);
  if (exact) return exact;
  const idField = (primaryColumns || []).find((column) => /^id$|_id$/i.test(normalizeUserConfigKey_(column)));
  return idField || (preferId ? ((primaryColumns || [])[0] || '') : '');
}

function hasConfigAutoTableField_(table, field) {
  if (!table || !field) return false;
  const fieldKey = normalizeUserConfigKey_(field);
  return getConfigAutoHeadersForTable_(table).some((header) => normalizeUserConfigKey_(header) === fieldKey);
}

function isConfigAutoManualKeepRow_(record) {
  const status = normalizeUserConfigKey_(recordValue_(record, ['status']));
  return ['manual', 'keep'].includes(status);
}

function createConfigAutoComponentRecord_(headers, view, componentType, label, index) {
  const record = createConfigAutoBlankRecord_(headers);
  setConfigAutoValue_(record, headers, ['component_id', 'component', 'id', 'key', 'config_key'], `${view.viewId}.${componentType}`);
  setConfigAutoValue_(record, headers, ['view_id', 'view', 'parent_id'], view.viewId);
  setConfigAutoValue_(record, headers, ['item', 'label', 'title', 'name'], label);
  setConfigAutoValue_(record, headers, ['component_type', 'type', 'display_type'], componentType);
  setConfigAutoValue_(record, headers, ['enabled', 'active', 'visible'], 'TRUE');
  setConfigAutoValue_(record, headers, ['sort_order', 'sort', 'order'], view.sortOrder + index);
  setConfigAutoValue_(record, headers, ['status'], 'Active');
  return record;
}

function appendMissingConfigRows_(item, data, keyNames, records) {
  const headers = data.headers || [];
  if (!item || !item.table || !headers.length || !records.length) {
    return { addedCount: 0, skipped: [] };
  }

  const existingKeys = (data.rows || []).reduce((acc, row) => {
    const key = normalizeConfigAutoId_(recordValue_(row.record || {}, keyNames));
    if (key) acc[key] = true;
    return acc;
  }, {});
  const addedKeys = {};
  const missing = records.filter((record) => {
    const key = normalizeConfigAutoId_(recordValue_(record, keyNames));
    if (!key || existingKeys[key] || addedKeys[key]) return false;
    addedKeys[key] = true;
    return true;
  });

  if (!missing.length) return { addedCount: 0, skipped: [] };
  const addedCount = isNativeReadableTable_(item.table)
    ? importNativeRecords_(item.table, missing)
    : importSystemRecords_(item.table, missing);
  return {
    addedCount,
    addedKeys: missing.map((record) => recordValue_(record, keyNames)).filter(Boolean),
  };
}

function combineConfigAutoSyncResult_(appendResult, cleanupResult) {
  const append = appendResult || { addedCount: 0, skipped: [] };
  const cleanup = cleanupResult || { deletedCount: 0, skipped: [] };
  return Object.assign({}, append, {
    deletedCount: cleanup.deletedCount || 0,
    deletedKeys: cleanup.deletedKeys || [],
    skipped: (append.skipped || []).concat(cleanup.skipped || []),
  });
}

function cleanupStaleConfigPageRows_(item, data, expectedRecords) {
  const expected = getConfigAutoExpectedKeys_(expectedRecords, ['page_id', 'page', 'id', 'key', 'config_key']);
  const staleRows = ((data && data.rows) || []).filter((row) => {
    const record = row.record || {};
    const key = normalizeConfigAutoId_(recordValue_(record, ['page_id', 'page', 'id', 'key', 'config_key']));
    if (!key || expected[key]) return false;
    return isConfigAutoRemovablePageRow_(record);
  });
  return clearConfigAutoRows_(item, staleRows, ['page_id', 'page', 'id', 'key', 'config_key']);
}

function cleanupStaleGeneratedConfigRows_(item, data, expectedRecords, keyNames, generatedPredicate) {
  const expected = getConfigAutoExpectedKeys_(expectedRecords, keyNames);
  const staleRows = ((data && data.rows) || []).filter((row) => {
    const record = row.record || {};
    const key = normalizeConfigAutoId_(recordValue_(record, keyNames));
    if (!key || expected[key]) return false;
    return generatedPredicate(record);
  });
  return clearConfigAutoRows_(item, staleRows, keyNames);
}

function getConfigAutoExpectedKeys_(records, keyNames) {
  return (records || []).reduce((acc, record) => {
    const key = normalizeConfigAutoId_(recordValue_(record, keyNames));
    if (key) acc[key] = true;
    return acc;
  }, {});
}

function clearConfigAutoRows_(item, rows, keyNames) {
  if (!item || !item.table || !rows || !rows.length) return { deletedCount: 0, deletedKeys: [], skipped: [] };
  const deletedKeys = [];
  const skipped = [];
  rows.forEach((row) => {
    try {
      const key = recordValue_(row.record || {}, keyNames);
      if (isNativeReadableTable_(item.table)) clearNativeRecord_(item.table, { rowNumber: row.rowNumber });
      else clearSystemRecord_(item.table, { rowNumber: row.rowNumber });
      if (key) deletedKeys.push(key);
    } catch (error) {
      skipped.push(error && error.message ? error.message : String(error));
    }
  });
  return {
    deletedCount: deletedKeys.length,
    deletedKeys,
    skipped,
  };
}

function isConfigAutoRemovablePageRow_(record) {
  const pageType = normalizeUserConfigKey_(recordValue_(record, ['page_type', 'type', 'item_type']));
  const status = normalizeUserConfigKey_(recordValue_(record, ['status']));
  if (['draft', 'custom', 'manual', 'keep'].includes(status) || pageType === 'custom') return false;
  return ['table', 'report', 'overview', 'dashboard', 'external', 'external_link', 'link', 'embed', 'embedded', 'iframe', 'form', ''].includes(pageType);
}

function isConfigAutoGeneratedViewRow_(record) {
  const viewId = recordValue_(record, ['view_id', 'view', 'id', 'key', 'config_key']);
  const pageId = recordValue_(record, ['page_id', 'page', 'parent_id']);
  const viewType = normalizeUserViewType_(recordValue_(record, ['view_type', 'type', 'display_type']));
  if (!viewId || !pageId || !CONFIG_AUTO_VIEW_TEMPLATES.some((template) => template.type === viewType)) return false;
  return normalizeConfigAutoId_(viewId) === normalizeConfigAutoId_(`${pageId}.${viewType}`);
}

function isConfigAutoGeneratedFormRow_(record) {
  const formId = recordValue_(record, ['form_id', 'form', 'id', 'key', 'config_key']);
  const pageId = recordValue_(record, ['page_id', 'page', 'parent_id']) || inferPageIdFromFormId_(formId);
  const rawFormType = recordValue_(record, ['form_type', 'form_layout_type', 'layout_type']);
  const mode = recordValue_(record, ['mode', 'form_mode', 'edit_mode', 'type'])
    || (typeof isUserFormModeValue_ === 'function' && isUserFormModeValue_(rawFormType) ? rawFormType : '')
    || formId;
  const formType = normalizeUserFormType_(mode);
  if (!formId || !pageId || !['quick', 'full', 'nested'].includes(formType)) return false;
  return normalizeConfigAutoId_(formId) === normalizeConfigAutoId_(`${pageId}.${formType}_form`);
}

function isConfigAutoGeneratedComponentRow_(record) {
  const componentId = recordValue_(record, ['component_id', 'component', 'id', 'key', 'config_key']);
  const viewId = recordValue_(record, ['view_id', 'view', 'parent_id']);
  const componentType = normalizeUserConfigKey_(recordValue_(record, ['component_type', 'type', 'display_type']));
  if (!componentId || !viewId || !componentType) return false;
  const knownComponentTypes = Object.keys(CONFIG_AUTO_COMPONENT_TEMPLATES)
    .flatMap((viewType) => CONFIG_AUTO_COMPONENT_TEMPLATES[viewType].map(([type]) => type));
  if (!knownComponentTypes.includes(componentType)) return false;
  return normalizeConfigAutoId_(componentId) === normalizeConfigAutoId_(`${viewId}.${componentType}`);
}

function createConfigAutoBlankRecord_(headers) {
  return (headers || []).reduce((record, header) => {
    record[header] = '';
    return record;
  }, {});
}

function setConfigAutoValue_(record, headers, names, value) {
  if (value === undefined || value === null || value === '') return;
  const header = findConfigAutoHeader_(headers, names);
  if (header) record[header] = value;
}

function findConfigAutoHeader_(headers, names) {
  const normalizedNames = names.map(normalizeConfigSeedHeader_);
  return (headers || []).find((header) => normalizedNames.includes(normalizeConfigSeedHeader_(header))) || '';
}

function isConfigAutoActiveRecord_(record) {
  const status = normalizeUserConfigKey_(recordValue_(record, ['status']));
  return !status || status === 'active' || status === 'enabled' || status === 'true';
}

function isConfigAutoPageNavType_(itemType) {
  return ['page', 'table', 'report', 'overview', 'dashboard', 'external', 'external_link', 'link', 'embed', 'iframe', 'form'].includes(itemType);
}

function inferConfigAutoPageType_(itemType, label, pageId) {
  const hint = normalizeUserConfigKey_([itemType, label, pageId].filter(Boolean).join(' '));
  const type = normalizeUserConfigKey_(itemType);
  const labelKey = normalizeUserConfigKey_(label);
  const pageKey = normalizeUserConfigKey_(pageId);
  if (['report', 'overview', 'dashboard'].includes(type) || labelKey === 'overview' || pageKey === 'overview') return 'report';
  if (['external', 'external_link', 'link'].includes(type) || hint.indexOf('external') >= 0) return 'external';
  if (['embed', 'embedded', 'embed_link', 'iframe'].includes(type) || hint.indexOf('embed') >= 0) return 'embed';
  if (['form', 'input_form'].includes(type)) return 'form';
  return 'table';
}

function getConfigAutoUserTablesByName_() {
  const tables = getTables_(true);
  return Object.keys(tables).reduce((acc, key) => {
    const table = tables[key];
    if (!table || table.type === 'config') return acc;
    acc[normalizeUserConfigKey_(table.name)] = table.name;
    acc[normalizeUserConfigKey_(stripTableLabelIcon_(table.name))] = table.name;
    (table.aliases || []).forEach((alias) => {
      acc[normalizeUserConfigKey_(alias)] = table.name;
      acc[normalizeUserConfigKey_(stripTableLabelIcon_(alias))] = table.name;
    });
    return acc;
  }, {});
}

function getConfigAutoHeadersForTable_(table) {
  try {
    if (isNativeReadableTable_(table)) return getNativeHeaders_(table);
    const sheet = getSheet_(table);
    return getSystemTableLayout_(sheet, table).headers || [];
  } catch (error) {
    return [];
  }
}

function getConfigAutoValidationByHeader_(table, headers) {
  try {
    const nativeValidation = readNativeTableColumnValidationByHeader_(table, headers);
    const cellValidation = readCellValidationByHeader_(table, headers);
    return Object.assign({}, cellValidation, nativeValidation);
  } catch (error) {
    return {};
  }
}

function isConfigAutoGroupableField_(field, inputType) {
  const key = normalizeUserConfigKey_(field);
  if (!key) return false;
  if (['note', 'notes', 'detail', 'details', 'description', 'summary', 'comment', 'comments'].some((part) => key.indexOf(part) >= 0)) return false;
  if (['url', 'link', 'photo', 'image', 'picture', 'thumbnail', 'file', 'attachment'].some((part) => key.indexOf(part) >= 0)) return false;
  if (inputType === 'number') return false;
  return true;
}

function inferConfigAutoTableName_(request, userTables) {
  const candidates = [
    request.label,
    request.pageId,
    String(request.pageId || '').split('.').pop(),
    String(request.pageId || '').split('_').pop(),
  ].filter(Boolean);
  for (let index = 0; index < candidates.length; index += 1) {
    const key = normalizeUserConfigKey_(candidates[index]);
    if (userTables[key]) return userTables[key];
  }
  return '';
}

function normalizeConfigAutoId_(value) {
  return normalizeUserConfigKey_(value);
}
