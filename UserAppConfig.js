/**
 * User app configuration adapter.
 *
 * Reads the Config app tables and exposes a compact UI/service config for the
 * User app. Keep this layer tolerant of lean table shapes so config tables can
 * evolve without breaking the app shell.
 */
function getUserAppConfig_(forceRefresh, options) {
  const deferDataConfig = !!(options && options.deferDataConfig);
  const cacheKey = cacheKey_('registry', deferDataConfig ? 'userAppConfig:lite' : 'userAppConfig');
  if (!forceRefresh) {
    const cached = getCached_(cacheKey);
    if (cached) return cached;
  }

  const tables = getTables_(forceRefresh);
  const tableNameToKey = Object.keys(tables).reduce((acc, key) => {
    getTableLookupAliases_(tables[key]).forEach((alias) => {
      acc[normalizeUserConfigKey_(alias)] = key;
    });
    return acc;
  }, {});
  const context = { tables, tableNameToKey };
  const pages = buildUserPageConfig_(readUserConfigRecords_('page', forceRefresh), context);
  context.pages = pages;
  const views = buildUserViewConfig_(readUserConfigRecords_('view', forceRefresh), context);
  context.views = views;
  const relationships = deferDataConfig
    ? createDeferredRelationshipConfig_()
    : buildUserRelationshipConfig_(readUserConfigRecords_('relationship', forceRefresh), context);
  context.relationships = relationships;
  const config = {
    app: buildUserAppSettings_(readUserConfigRecords_('appConfig', forceRefresh)),
    pages,
    views,
    forms: deferDataConfig
      ? createDeferredFormConfig_()
      : buildUserFormConfig_(readUserConfigRecords_('form', forceRefresh), context),
    relationships,
    components: buildUserComponentConfig_(readUserConfigRecords_('component', forceRefresh), context),
    dataColumns: deferDataConfig
      ? createDeferredDataColumnConfig_()
      : buildUserDataColumnConfig_(readUserConfigRecords_('dataColumn', forceRefresh), context),
    dataServices: buildUserDataServiceConfig_(readUserConfigRecords_('dataServices', forceRefresh), context),
    deferredDataConfig: deferDataConfig,
    generatedAt: new Date().toISOString(),
  };

  setCached_(cacheKey, config, UI_CONFIG_CACHE_TTL_SECONDS);
  return config;
}

function createDeferredRelationshipConfig_() {
  return { byId: {}, byPair: {}, list: [], deferred: true };
}

function createDeferredDataColumnConfig_() {
  return { byTableKey: {}, deferred: true };
}

function createDeferredFormConfig_() {
  return { byId: {}, byPageId: {}, byViewId: {}, deferred: true };
}

function readUserConfigRecords_(configKey, forceRefresh) {
  try {
    const item = getConfigItemOrThrow_(configKey, forceRefresh);
    if (!item || !item.table) return [];
    const rows = item.dataConfig
      ? getConfigTableData(configKey, forceRefresh).rows || []
      : buildTableRowsData_(item.table).rows;
    const records = rows.map((row) => row.record || {});
    return normalizeUserConfigRecords_(configKey, records);
  } catch (error) {
    logAppDiagnostic_('warn', 'user_config_records_skipped', {
      configKey: configKey || '',
      forceRefresh: !!forceRefresh,
    }, error);
    return [];
  }
}

function normalizeUserConfigRecords_(configKey, records) {
  if (configKey === 'appConfig' || !isKeyValueConfigShape_(records)) return records;
  const grouped = {};
  records.forEach((record) => {
    const group = recordValue_(record, ['config_group', 'group', 'category', 'section']);
    const name = recordValue_(record, ['config_name', 'name', 'item', 'key']);
    const value = recordValue_(record, ['config_value', 'value', 'setting_value']);
    if (!group || !name) return;
    grouped[group] = grouped[group] || getDefaultGroupedConfigRecord_(configKey, group);
    grouped[group][name] = value;
  });
  return Object.keys(grouped).map((key) => grouped[key]);
}

function isKeyValueConfigShape_(records) {
  return (records || []).some((record) => (
    recordValue_(record, ['config_group', 'group', 'category', 'section']) &&
    recordValue_(record, ['config_name', 'name', 'item', 'key'])
  ));
}

function getDefaultGroupedConfigRecord_(configKey, group) {
  const record = {
    'Config group': group,
    Status: 'Active',
  };
  if (configKey === 'page') {
    record['Page ID'] = group;
    record.Item = group;
  } else if (configKey === 'view') {
    record['View ID'] = group;
    record.Item = group;
  } else if (configKey === 'component') {
    record['Component ID'] = group;
    record.Item = group;
  } else if (configKey === 'form') {
    record['Form ID'] = group;
    record.Item = group;
  } else if (configKey === 'dataServices') {
    record['Service ID'] = group;
    record.Item = group;
  } else if (configKey === 'dataColumn') {
    record['Column ID'] = group;
    record.Item = group;
  } else if (configKey === 'navigation') {
    record['Nav ID'] = group;
    record.Item = group;
  }
  return record;
}

function buildUserAppSettings_(records) {
  const settings = {
    appName: '',
    logoUrl: '',
    accentColor: '',
    defaultLandingPage: '',
    defaultTableView: 'table',
    showSearch: true,
    showStatusFilter: true,
    enableGrouping: true,
    enableCreate: true,
    enableEdit: true,
    enableDelete: true,
    enableImport: true,
    enablePrefetch: true,
    prefetchLimit: 4,
    prefetchIdleMs: 10000,
    prefetchGapMs: 3000,
    showOpenSheet: true,
    showViewSwitcher: true,
  };

  records
    .filter(isActiveUserConfigRecord_)
    .forEach((record) => {
      const key = normalizeUserConfigKey_(recordValue_(record, ['config_key', 'config_name', 'key', 'id', 'item', 'name', 'label']));
      const value = recordValue_(record, ['value', 'config_value', 'setting_value', 'color', 'url', 'detail']);
      if (!key || value === '') return;
      if (['app_name', 'process_name', 'project_name'].includes(key)) settings.appName = value;
      if (['logo_url', 'brand_logo', 'logo'].includes(key)) settings.logoUrl = value;
      if (['accent_color', 'primary_color', 'theme_color', 'brand_color'].includes(key)) settings.accentColor = value;
      if (['default_landing_page', 'landing_page', 'start_page', 'home_page', 'default_page'].includes(key)) settings.defaultLandingPage = value;
      if (['default_table_view', 'table_view', 'default_view'].includes(key)) settings.defaultTableView = normalizeUserViewType_(value);
      if (['show_search', 'search_enabled'].includes(key)) settings.showSearch = configBool_(value, settings.showSearch);
      if (['show_status_filter', 'status_filter_enabled'].includes(key)) settings.showStatusFilter = configBool_(value, settings.showStatusFilter);
      if (['enable_grouping', 'grouping_enabled'].includes(key)) settings.enableGrouping = configBool_(value, settings.enableGrouping);
      if (['enable_create', 'create_enabled', 'new_row_enabled'].includes(key)) settings.enableCreate = configBool_(value, settings.enableCreate);
      if (['enable_edit', 'edit_enabled'].includes(key)) settings.enableEdit = configBool_(value, settings.enableEdit);
      if (['enable_delete', 'delete_enabled'].includes(key)) settings.enableDelete = configBool_(value, settings.enableDelete);
      if (['enable_import', 'import_enabled'].includes(key)) settings.enableImport = configBool_(value, settings.enableImport);
      if (['enable_prefetch', 'prefetch_enabled', 'navigation_prefetch_enabled'].includes(key)) settings.enablePrefetch = configBool_(value, settings.enablePrefetch);
      if (['prefetch_limit', 'navigation_prefetch_limit', 'prefetch_pages'].includes(key)) settings.prefetchLimit = clampUserConfigNumber_(value, settings.prefetchLimit, 0, 8);
      if (['prefetch_idle_ms', 'navigation_prefetch_idle_ms', 'prefetch_delay_ms'].includes(key)) settings.prefetchIdleMs = clampUserConfigNumber_(value, settings.prefetchIdleMs, 5000, 30000);
      if (['prefetch_gap_ms', 'navigation_prefetch_gap_ms'].includes(key)) settings.prefetchGapMs = clampUserConfigNumber_(value, settings.prefetchGapMs, 1500, 10000);
      if (['show_open_sheet', 'open_sheet_enabled'].includes(key)) settings.showOpenSheet = configBool_(value, settings.showOpenSheet);
      if (['show_view_switcher', 'view_switcher_enabled'].includes(key)) settings.showViewSwitcher = configBool_(value, settings.showViewSwitcher);
    });

  return settings;
}

function clampUserConfigNumber_(value, defaultValue, min, max) {
  const number = Number(value);
  if (!Number.isFinite(number)) return defaultValue;
  return Math.max(min, Math.min(max, Math.round(number)));
}

function buildUserPageConfig_(records, context) {
  return records
    .filter(isActiveUserConfigRecord_)
    .map((record) => {
      const tableName = recordValue_(record, ['table_name', 'table', 'source_table']);
      const tableKey = resolveUserConfigTableKey_(tableName, context);
      const pageId = recordValue_(record, ['page_id', 'id', 'key', 'config_key']) || toTableKey_(recordValue_(record, ['item', 'title', 'label']) || tableName || 'page');
      const pageType = normalizeUserPageType_(recordValue_(record, ['page_type', 'type', 'item_type']));
      return {
        id: pageId,
        pageType,
        tableName,
        tableKey,
        reportName: recordValue_(record, ['report_name', 'report', 'overview_key', 'report_key']),
        linkUrl: normalizeUserLinkUrl_(recordValue_(record, ['link_url', 'url', 'embed_url', 'external_url'])),
        label: recordValue_(record, ['item', 'title', 'label', 'name']) || tableName || pageId,
        subtitle: recordValue_(record, ['subtitle', 'detail', 'summary', 'description', 'note']),
        titleColumn: recordValue_(record, [
          'title_column',
          'record_title_column',
          'record_name_column',
          'display_name_column',
          'display_column',
          'display_name',
          'display_field',
          'main_display_name',
          'main_display_field',
          'title_field',
          'record_title_field',
          'primary_field',
          'name_field',
        ]),
        aiInsightEnabled: configBool_(recordValue_(record, ['ai_insight_enable', 'insight_enable', 'enable_ai_insight']), false),
        aiInsightColumn: recordValue_(record, ['ai_insight_column', 'insight_column', 'ai_column', 'insight_field']),
        viewType: normalizeUserViewType_(recordValue_(record, ['default_view', 'default_view_type', 'default_table_view', 'view_type', 'type'])),
        defaultGroupBy: parseUserConfigList_(recordValue_(record, ['default_group_by', 'default_groups', 'default_group_fields', 'group_by', 'group_fields'])),
        dataMode: normalizeUserDataMode_(recordValue_(record, ['data_mode', 'data_mode_key', 'loading_mode', 'table_data_mode'])),
        sortOrder: Number(recordValue_(record, ['sort_order', 'sort', 'order']) || 0),
      };
    })
    .reduce((acc, page) => {
      if (page.tableKey) acc.byTableKey[page.tableKey] = page;
      if (page.id) acc.byId[page.id] = page;
      return acc;
    }, { byTableKey: {}, byId: {} });
}

function normalizeUserDataMode_(value) {
  const key = normalizeUserConfigKey_(value);
  if (['paged', 'page', 'server', 'server_side', 'transaction', 'transactions'].includes(key)) return 'paged';
  return 'full';
}

function normalizeUserPageType_(value) {
  const key = normalizeUserConfigKey_(value);
  if (['report', 'overview', 'dashboard'].includes(key)) return 'report';
  if (['external', 'external_link', 'link'].includes(key)) return 'external';
  if (['embed', 'embedded', 'embeded', 'embed_link', 'embedded_link', 'embeded_link', 'iframe'].includes(key)) return 'embed';
  if (['form', 'input_form'].includes(key)) return 'form';
  return 'table';
}

function normalizeUserLinkUrl_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  if (/^https?:\/\//i.test(text)) return text;
  return `https://${text}`;
}

function buildUserViewConfig_(records, context) {
  return records
    .filter(isActiveUserConfigRecord_)
    .map((record) => {
      const pageId = recordValue_(record, ['page_id', 'page', 'parent_id']);
      const page = pageId && context.pages && context.pages.byId ? context.pages.byId[pageId] : null;
      const tableName = recordValue_(record, ['table_name', 'table', 'source_table']) || (page && page.tableName) || '';
      const tableKey = resolveUserConfigTableKey_(tableName, context);
      const type = normalizeUserConfigKey_(recordValue_(record, ['view_type', 'type', 'display_type'])) || 'table';
      return {
        id: recordValue_(record, ['view_id', 'id', 'key', 'config_key']) || `${tableKey || toTableKey_(tableName)}_${type}`,
        pageId,
        tableName,
        tableKey,
        label: recordValue_(record, ['item', 'title', 'label', 'name']) || type,
        type: normalizeUserViewType_(type),
        displayNameColumn: recordValue_(record, [
          'title_column',
          'record_title_column',
          'record_name_column',
          'display_name_column',
          'display_column',
          'display_name',
          'display_field',
          'main_display_name',
          'main_display_field',
          'title_field',
          'record_title_field',
          'primary_field',
          'name_field',
        ]),
        defaultGroupBy: parseUserConfigList_(recordValue_(record, ['default_group_by', 'default_groups', 'default_group_fields', 'group_by', 'group_fields'])),
        columnOrder: parseUserConfigList_(recordValue_(record, ['column_order', 'field_order', 'columns', 'visible_columns'])),
        hiddenColumns: parseUserConfigList_(recordValue_(record, ['hidden_columns', 'hide_columns', 'hidden_fields'])),
        sections: parseUserFormSections_(recordValue_(record, ['section_layout', 'form_sections', 'sections'])),
        columns: clampFormColumnCount_(recordValue_(record, ['columns', 'layout_columns', 'column_count']), 2),
        labelPosition: normalizeUserConfigKey_(recordValue_(record, ['label_position', 'label_layout'])) || 'top',
        showEmptyFields: configBool_(recordValue_(record, ['show_empty_fields', 'show_empty']), true),
        filterFields: parseUserConfigList_(recordValue_(record, ['filter_fields', 'filter_columns', 'smart_filter_fields', 'filters'])),
        defaultFilters: parseUserViewFilters_(recordValue_(record, ['default_filters', 'predefined_filters', 'preset_filters', 'filter_config'])),
        defaultSort: parseUserViewSort_(
          recordValue_(record, ['default_sort', 'default_sort_field', 'sort_field', 'order_by']),
          recordValue_(record, ['default_sort_direction', 'sort_direction', 'order_direction'])
        ),
        sortOrder: Number(recordValue_(record, ['sort_order', 'sort', 'order']) || 0),
      };
    })
    .reduce((acc, view) => {
      if (view.tableKey) {
        acc.byTableKey[view.tableKey] = acc.byTableKey[view.tableKey] || [];
        acc.byTableKey[view.tableKey].push(view);
        acc.byTableKey[view.tableKey].sort((a, b) => a.sortOrder - b.sortOrder);
      }
      if (view.pageId) {
        acc.byPageId[view.pageId] = acc.byPageId[view.pageId] || [];
        acc.byPageId[view.pageId].push(view);
        acc.byPageId[view.pageId].sort((a, b) => a.sortOrder - b.sortOrder);
      }
      if (view.id) acc.byId[view.id] = view;
      return acc;
    }, { byTableKey: {}, byPageId: {}, byId: {} });
}

function buildUserComponentConfig_(records, context) {
  return records
    .map((record) => {
      const type = normalizeUserConfigKey_(recordValue_(record, ['component_type', 'type', 'item', 'label', 'name']));
      const viewId = recordValue_(record, ['view_id', 'view']);
      const linkedView = viewId && context && context.views && context.views.byId ? context.views.byId[viewId] : null;
      const rawViewType = recordValue_(record, ['view_type', 'view_type_key']) || (linkedView && linkedView.type) || '';
      const viewType = rawViewType ? normalizeUserViewType_(rawViewType) : '';
      return {
        id: recordValue_(record, ['component_id', 'id', 'key', 'config_key']) || type,
        viewId,
        type,
        label: recordValue_(record, ['item', 'label', 'title', 'name']) || type,
        enabled: isActiveUserConfigRecord_(record) &&
          configBool_(recordValue_(record, ['enabled', 'active', 'visible']), true) &&
          !configBool_(recordValue_(record, ['hidden', 'disabled']), false),
        viewType,
        configValue: recordValue_(record, ['config_value', 'value', 'setting_value']),
        sortOrder: Number(recordValue_(record, ['sort_order', 'sort', 'order']) || 0),
      };
    })
    .filter((component) => component.type)
    .reduce((acc, component) => {
      if (component.id) acc.byId[component.id] = component;
      if (component.viewId) {
        acc.byViewId[component.viewId] = acc.byViewId[component.viewId] || {};
        acc.byViewId[component.viewId][component.type] = component;
      }
      if (component.viewType) {
        acc.byViewType[component.viewType] = acc.byViewType[component.viewType] || {};
        acc.byViewType[component.viewType][component.type] = component;
      }
      acc.byType[component.type] = component;
      acc.list.push(component);
      return acc;
    }, { byId: {}, byViewId: {}, byViewType: {}, byType: {}, list: [] });
}

function buildUserFormConfig_(records, context) {
  return records
    .filter(isActiveUserConfigRecord_)
    .map((record) => {
      const viewId = recordValue_(record, ['view_id', 'view']);
      const pageId = recordValue_(record, ['page_id', 'page', 'parent_id']) || inferPageIdFromFormId_(recordValue_(record, ['form_id', 'id', 'key', 'config_key']));
      const rawFormLayoutType = recordValue_(record, ['form_type', 'form_layout_type', 'layout_type', 'flow_type']);
      const rawMode = recordValue_(record, ['mode', 'form_mode', 'edit_mode', 'type'])
        || (isUserFormModeValue_(rawFormLayoutType) ? rawFormLayoutType : '')
        || recordValue_(record, ['item', 'title', 'label', 'name'])
        || recordValue_(record, ['form_id', 'id', 'key', 'config_key']);
      const formType = normalizeUserFormType_(rawMode);
      const formLayoutType = normalizeUserFormLayoutType_(rawFormLayoutType);
      const linkedView = viewId && context && context.views && context.views.byId ? context.views.byId[viewId] : null;
      const columns = clampFormColumnCount_(recordValue_(record, ['columns', 'layout_columns', 'column_count']), formType === 'full' || (linkedView && linkedView.type === 'detail') ? 3 : 1);
      const columnOrder = parseUserConfigList_(recordValue_(record, ['column_order', 'field_order', 'columns_order']));
      const hiddenColumns = parseUserConfigList_(recordValue_(record, ['hidden_columns', 'hide_columns', 'hidden_fields']));
      const nestedTables = parseUserConfigList_(recordValue_(record, ['nested_table', 'nested_tables', 'related_table', 'related_tables']));
      const nestedTablesExplicit = nestedTables.length > 0;
      const sections = parseUserFormSections_(recordValue_(record, ['section_layout', 'form_sections', 'sections']));
      const mainDisplayName = recordValue_(record, [
        'main_display_name',
        'main_display_field',
        'display_name',
        'display_field',
        'title_field',
        'record_title_field',
        'primary_field',
        'name_field',
      ]);
      return {
        id: recordValue_(record, ['form_id', 'id', 'key', 'config_key']) || (pageId ? `${pageId}.${formType}_form` : viewId ? `${viewId}.form` : 'form'),
        pageId,
        viewId,
        type: formType,
        label: recordValue_(record, ['item', 'title', 'label', 'name']) || 'Form',
        columns,
        labelPosition: normalizeUserConfigKey_(recordValue_(record, ['label_position', 'label_layout'])) || 'top',
        showEmptyFields: configBool_(recordValue_(record, ['show_empty_fields', 'show_empty']), true),
        formLayoutType,
        columnOrder,
        hiddenColumns,
        mainDisplayName,
        sections,
        nestedTables: uniqueUserConfigList_(nestedTables),
        nestedTablesExplicit,
      };
    })
    .filter((form) => !isLegacyViewFormConfig_(form))
    .reduce((acc, form) => {
      if (form.id) acc.byId[form.id] = form;
      if (form.pageId) {
        acc.byPageId[form.pageId] = acc.byPageId[form.pageId] || {};
        acc.byPageId[form.pageId][form.type || 'quick'] = form;
      }
      if (form.viewId) acc.byViewId[form.viewId] = form;
      return acc;
    }, { byId: {}, byPageId: {}, byViewId: {} });
}

function getUserTableNameByKey_(tableKey, context) {
  const table = context && context.tables ? context.tables[tableKey] : null;
  return table ? table.name || table.apiTableName || table.sheetName || tableKey : tableKey;
}

function uniqueUserConfigList_(items) {
  const seen = {};
  return (items || []).filter((item) => {
    const key = normalizeUserConfigKey_(item);
    if (!key || seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function normalizeUserFormType_(value) {
  const key = normalizeUserConfigKey_(value);
  if (key.indexOf('nested') >= 0 || key.indexOf('related') >= 0 || key.indexOf('subform') >= 0 || key.indexOf('sub_form') >= 0) return 'nested';
  if (key.indexOf('full') >= 0 || key.indexOf('advanced') >= 0) return 'full';
  return 'quick';
}

function isUserFormModeValue_(value) {
  const key = normalizeUserConfigKey_(value);
  return key.indexOf('full') >= 0 || key.indexOf('advanced') >= 0 || key.indexOf('quick') >= 0 || key.indexOf('nested') >= 0 || key.indexOf('related') >= 0 || key.indexOf('subform') >= 0 || key.indexOf('sub_form') >= 0;
}

function normalizeUserFormLayoutType_(value) {
  const key = normalizeUserConfigKey_(value);
  if (['multi_step', 'multistep', 'flow', 'stepper', 'wizard'].includes(key)) return 'multi-step';
  return 'default';
}

function inferPageIdFromFormId_(formId) {
  const text = String(formId || '');
  return text.replace(/\.(quick|full|nested)_form$/i, '');
}

function isLegacyViewFormConfig_(form) {
  const id = String((form && form.id) || '');
  return /\.(table|detail|card)\.form$/i.test(id);
}

function clampFormColumnCount_(value, defaultValue) {
  const count = Number(value || defaultValue || 1);
  if (!Number.isFinite(count)) return defaultValue || 1;
  return Math.max(1, Math.min(3, Math.round(count)));
}

function buildUserDataColumnConfig_(records, context) {
  return records
    .filter(isActiveUserConfigRecord_)
    .map((record) => {
      const tableName = recordValue_(record, ['table_name', 'table', 'source_table']);
      const field = recordValue_(record, ['field', 'column', 'column_name', 'field_name', 'item', 'label']);
      const tableKey = resolveUserConfigTableKey_(tableName, context);
      return {
        tableName,
        tableKey,
        field,
        fieldKey: normalizeUserConfigKey_(field),
        label: recordValue_(record, ['label', 'display_label', 'english_translation', 'item']) || field,
        inputType: normalizeUserFieldInputType_(recordValue_(record, ['input_type', 'input_control', 'control_type', 'input'])),
        dataType: normalizeUserFieldDataType_(recordValue_(record, ['data_type', 'type'])),
        required: configBool_(recordValue_(record, ['required', 'is_required', 'mandatory']), false),
        hidden: configBool_(recordValue_(record, ['hidden', 'hide', 'is_hidden']), false),
        groupable: configBool_(recordValue_(record, ['groupable', 'allow_group', 'enable_group']), true),
        defaultGroup: configBool_(recordValue_(record, ['default_group', 'group_by_default', 'default_group_by', 'is_default_group']), false),
        groupOrder: Number(recordValue_(record, ['group_order', 'default_group_order', 'group_sort_order']) || 0),
        sortOrder: Number(recordValue_(record, ['sort_order', 'sort', 'order']) || 0),
        section: recordValue_(record, ['section', 'group', 'category']),
      };
    })
    .filter((column) => column.tableKey && column.fieldKey)
    .reduce((acc, column) => {
      acc.byTableKey[column.tableKey] = acc.byTableKey[column.tableKey] || {};
      acc.byTableKey[column.tableKey][column.fieldKey] = column;
      return acc;
    }, { byTableKey: {} });
}

function buildUserDataServiceConfig_(records, context) {
  return records
    .filter(isActiveUserConfigRecord_)
    .map((record) => {
      const tableName = recordValue_(record, ['table_name', 'table', 'source_table']);
      const tableKey = resolveUserConfigTableKey_(tableName, context);
      const type = normalizeUserConfigKey_(recordValue_(record, ['service_type', 'type', 'item', 'label']));
      return {
        id: recordValue_(record, ['service_id', 'id', 'key', 'config_key']) || `${tableKey || 'app'}_${type}`,
        tableKey,
        tableName,
        type,
        enableRead: configBool_(recordValue_(record, ['enable_read', 'read_enabled']), true),
        enableCreate: configBool_(recordValue_(record, ['enable_create', 'create_enabled']), true),
        enableEdit: configBool_(recordValue_(record, ['enable_edit', 'edit_enabled']), true),
        enableDelete: configBool_(recordValue_(record, ['enable_delete', 'delete_enabled']), true),
        enableImport: configBool_(recordValue_(record, ['enable_import', 'import_enabled']), true),
      };
    })
    .reduce((acc, service) => {
      if (service.tableKey) acc.byTableKey[service.tableKey] = Object.assign(acc.byTableKey[service.tableKey] || {}, service);
      if (service.type) acc.byType[service.type] = service;
      return acc;
    }, { byTableKey: {}, byType: {} });
}

function buildUserRelationshipConfig_(records, context) {
  return records
    .filter(isActiveUserConfigRecord_)
    .map((record) => {
      const fromColumn = parseDataConfigColumnId_(recordValue_(record, ['from_column_id', 'from_column', 'source_column_id']));
      const toColumn = parseDataConfigColumnId_(recordValue_(record, ['to_column_id', 'to_column', 'target_column_id']));
      const sourceTable = recordValue_(record, ['source_table', 'parent_table', 'from_table', 'table_name', 'table']) || fromColumn.table;
      const targetTable = recordValue_(record, ['target_table', 'child_table', 'to_table', 'related_table']) || toColumn.table;
      const sourceTableKey = resolveUserConfigTableKey_(sourceTable, context);
      const targetTableKey = resolveUserConfigTableKey_(targetTable, context);
      const sourceField = recordValue_(record, ['source_field', 'parent_field', 'from_field', 'key_field', 'field']) || fromColumn.column;
      const targetField = recordValue_(record, ['target_field', 'child_field', 'to_field', 'lookup_field', 'related_field']) || toColumn.column;
      return {
        id: recordValue_(record, ['relationship_id', 'id', 'key', 'config_key']) || `${sourceTableKey}.${targetTableKey}.${normalizeUserConfigKey_(sourceField)}.${normalizeUserConfigKey_(targetField)}`,
        label: recordValue_(record, ['item', 'label', 'title', 'name']) || `${sourceTable} to ${targetTable}`,
        sourceTable,
        sourceTableKey,
        sourceField,
        targetTable,
        targetTableKey,
        targetField,
        type: normalizeUserConfigKey_(recordValue_(record, ['relationship_type', 'type'])) || 'lookup',
      };
    })
    .filter((relationship) => relationship.sourceTableKey && relationship.targetTableKey && relationship.sourceField && relationship.targetField)
    .reduce((acc, relationship) => {
      if (relationship.id) acc.byId[relationship.id] = relationship;
      const pairKey = `${relationship.sourceTableKey}::${relationship.targetTableKey}`;
      acc.byPair[pairKey] = acc.byPair[pairKey] || [];
      acc.byPair[pairKey].push(relationship);
      acc.list.push(relationship);
      return acc;
    }, { byId: {}, byPair: {}, list: [] });
}

function applyUserFieldConfig_(table, fields) {
  if (!table || table.type === 'config') return fields;
  const config = getUserAppConfig_(false, { deferDataConfig: false });
  const tableColumns = config.dataColumns.byTableKey[table.key] || {};
  return fields
    .map((field) => {
      const column = tableColumns[normalizeUserConfigKey_(field.name)] || null;
      if (!column) return field;
      const hasValidationOptions = !!(field.validation && Array.isArray(field.validation.options) && field.validation.options.length);
      const configuredInputType = column.inputType || (!hasValidationOptions ? column.dataType : '');
      return Object.assign({}, field, {
        label: column.label || field.label || field.name,
        inputType: configuredInputType || field.inputType,
        required: column.required,
        hidden: column.hidden,
        groupable: column.groupable,
        defaultGroup: column.defaultGroup,
        groupOrder: column.groupOrder,
        sortOrder: column.sortOrder || field.index,
        section: column.section || '',
      });
    })
    .sort((a, b) => Number(a.sortOrder == null ? a.index : a.sortOrder) - Number(b.sortOrder == null ? b.index : b.sortOrder));
}

function normalizeUserFieldInputType_(value) {
  const normalized = normalizeUserConfigKey_(value);
  if (['select', 'dropdown', 'drop_down', 'choice'].includes(normalized)) return 'select';
  if (['datalist', 'autocomplete', 'auto_complete', 'suggestion', 'suggestions'].includes(normalized)) return 'datalist';
  if (['textarea', 'multiline', 'multi_line', 'long_text'].includes(normalized)) return 'textarea';
  if (['date', 'number', 'text', 'url', 'email', 'tel', 'phone'].includes(normalized)) return normalized === 'phone' ? 'tel' : normalized;
  return '';
}

function normalizeUserFieldDataType_(value) {
  const normalized = normalizeUserConfigKey_(value);
  if (['date', 'datetime', 'time'].includes(normalized)) return 'date';
  if (['number', 'numeric', 'currency', 'amount', 'price', 'quantity', 'qty'].includes(normalized)) return 'number';
  return '';
}

function resolveUserConfigTableKey_(tableName, context) {
  if (!tableName) return '';
  return (context.tableNameToKey || {})[normalizeUserConfigKey_(tableName)] || '';
}

function isActiveUserConfigRecord_(record) {
  const status = normalizeUserConfigKey_(recordValue_(record, ['status']));
  return !status || status === 'active' || status === 'enabled' || status === 'true';
}

function recordValue_(record, names) {
  const normalizedNames = names.map(normalizeUserConfigKey_);
  const keys = Object.keys(record || {});
  for (let index = 0; index < keys.length; index += 1) {
    const key = keys[index];
    if (normalizedNames.includes(normalizeUserConfigKey_(key)) && record[key] !== '') return record[key];
  }
  return '';
}

function configBool_(value, defaultValue) {
  if (value === '' || value == null) return defaultValue;
  const text = normalizeUserConfigKey_(value);
  if (['false', 'no', 'n', '0', 'off', 'disabled', 'inactive', 'hide', 'hidden', 'optional', 'not_required', 'notrequired', 'khong_bat_buoc'].includes(text)) return false;
  if (['true', 'yes', 'y', '1', 'on', 'enabled', 'active', 'show', 'visible', 'required', 'mandatory', 'must', 'bat_buoc'].includes(text)) return true;
  return defaultValue;
}

function normalizeUserViewType_(value) {
  const type = normalizeUserConfigKey_(value);
  if (['detail', 'master_detail', 'masterdetail', 'list_detail', 'list_master_detail'].includes(type)) return 'detail';
  if (['card', 'cards', 'gallery'].includes(type)) return 'card';
  if (['nested_form', 'nestedform', 'nested', 'related_form', 'related_table', 'subform', 'sub_form'].includes(type)) return 'nested_form';
  if (['form', 'input_form'].includes(type)) return 'form';
  return 'table';
}

function parseUserConfigList_(value) {
  return String(value || '')
    .split(/[,;\n|]+/)
    .map((item) => item.trim())
    .filter(Boolean);
}

function parseUserViewFilters_(value) {
  const text = String(value || '').trim();
  if (!text) return [];

  const jsonFilters = parseUserViewFiltersJson_(text);
  if (jsonFilters.length) return jsonFilters;

  return text
    .split(/[;\n]+/)
    .map((part) => part.trim())
    .filter(Boolean)
    .map((part) => {
      const match = part.match(/^(.+?)(?:=>|=|:)(.+)$/);
      if (!match) return null;
      const field = match[1].trim();
      const values = match[2]
        .split(/[|,]+/)
        .map((item) => item.trim())
        .filter(Boolean);
      if (!field || !values.length) return null;
      return {
        field,
        fieldKey: normalizeUserConfigKey_(field),
        values,
      };
    })
    .filter(Boolean);
}

function parseUserViewSort_(fieldValue, directionValue) {
  let field = String(fieldValue || '').trim();
  if (!field) return null;
  let direction = normalizeUserSortDirection_(directionValue);
  const prefixedDesc = field.charAt(0) === '-';
  if (prefixedDesc) {
    direction = -1;
    field = field.slice(1).trim();
  }
  const suffix = field.match(/^(.+?)(?:\s*(?:=>|:|\|)\s*|\s+)(asc|ascending|a_z|az|desc|descending|z_a|za)$/i);
  if (suffix) {
    field = suffix[1].trim();
    direction = normalizeUserSortDirection_(suffix[2]);
  }
  if (!field) return null;
  return {
    field,
    fieldKey: normalizeUserConfigKey_(field),
    direction: direction || 1,
  };
}

function normalizeUserSortDirection_(value) {
  const key = normalizeUserConfigKey_(value);
  if (['desc', 'descending', 'z_a', 'za', 'down', 'reverse', 'giam', 'giảm'].includes(key)) return -1;
  if (['asc', 'ascending', 'a_z', 'az', 'up', 'tang', 'tăng'].includes(key)) return 1;
  return 0;
}

function parseUserViewFiltersJson_(text) {
  try {
    const parsed = JSON.parse(text);
    if (Array.isArray(parsed)) {
      return parsed.map((item) => {
        const field = item && (item.field || item.column || item.name);
        const rawValues = item && (item.values || item.value);
        const values = Array.isArray(rawValues) ? rawValues : String(rawValues || '').split(/[|,]+/);
        return field ? {
          field: String(field).trim(),
          fieldKey: normalizeUserConfigKey_(field),
          values: values.map((value) => String(value || '').trim()).filter(Boolean),
        } : null;
      }).filter((item) => item && item.values.length);
    }
    if (parsed && typeof parsed === 'object') {
      return Object.keys(parsed).map((field) => {
        const rawValues = Array.isArray(parsed[field]) ? parsed[field] : String(parsed[field] || '').split(/[|,]+/);
        return {
          field,
          fieldKey: normalizeUserConfigKey_(field),
          values: rawValues.map((value) => String(value || '').trim()).filter(Boolean),
        };
      }).filter((item) => item.values.length);
    }
  } catch (error) {
    return [];
  }
  return [];
}

function parseUserFormSections_(value) {
  return normalizeUserFormSectionLayoutText_(value)
    .split(/\n+/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const parts = line.split(':');
      const title = (parts.shift() || '').trim();
      const fields = parseUserConfigList_(parts.join(':'));
      return { title, fields };
    })
    .filter((section) => section.title);
}

function normalizeUserFormSectionLayoutText_(value) {
  const text = String(value || '').replace(/\r\n?/g, '\n').trim();
  if (!text) return '';
  return text.replace(
    /([^\n])\s+((?:step|section|phase|part|bước|buoc)\s*\d+\s*:)/gi,
    '$1\n$2'
  );
}

function normalizeUserConfigKey_(value) {
  return String(value || '')
    .replace(/^[^A-Za-z0-9\u00C0-\u1EF9_]+/, '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9\u00C0-\u1EF9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}
