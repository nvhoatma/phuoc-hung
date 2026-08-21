/**
 * Navigation builder.
 *
 * Turns config rows into the tree consumed by ClientNavigation.html.
 * Keep label/tree decisions here; client navigation should only render
 * the tree it receives.
 */
function getNavigation(forceRefresh) {
  assertSpreadsheetAccess_();
  const cacheKey = cacheKey_('registry', 'navigation');
  if (!forceRefresh) {
    const cached = getCached_(cacheKey);
    if (cached) return cached;
  }
  try {
    const navigation = buildNavigationFromConfig_(!!forceRefresh);
    setCached_(cacheKey, navigation, UI_CONFIG_CACHE_TTL_SECONDS);
    return navigation;
  } catch (error) {
    console.warn('Navigation config error:', error && error.message ? error.message : error);
    return [];
  }
}

function buildNavigationFromConfig_(forceRefresh) {
  const userConfig = getUserAppConfig_(!!forceRefresh, { deferDataConfig: true });
  const rows = readNavigationConfigRows_(!!forceRefresh)
    .filter((row) => normalizeNavText_(getNavConfigValue_(row, ['status']) || 'Active') === 'active')
    .sort((a, b) => getNavSortOrder_(a) - getNavSortOrder_(b));

  const items = rows.map((row) => toNavigationItem_(row, userConfig));
  const byId = items.reduce((acc, item) => {
    acc[item.id] = item;
    return acc;
  }, {});

  const roots = [];
  items.forEach((item) => {
    const parentId = item.parentId;
    delete item.parentId;
    if (parentId && byId[parentId]) {
      byId[parentId].children = byId[parentId].children || [];
      byId[parentId].children.push(item);
      return;
    }
    roots.push(item);
  });

  return roots;
}

function readNavigationConfigRows_(forceRefresh) {
  const records = readUserConfigRecords_('navigation', !!forceRefresh);
  if (records && records.length) return records;

  const nativeTable = getNativeTableByName_('Nav config', !!forceRefresh) ||
    getNativeTableByName_('Setting Nav', !!forceRefresh) ||
    getNativeTableByName_('Navigation config', !!forceRefresh);
  if (!nativeTable) return [];
  const sheetMeta = getNativeSheetsById_(!!forceRefresh)[String(nativeTable.range.sheetId)] || {};
  const table = {
    key: 'config_navigation',
    name: nativeTable.name,
    sheetName: sheetMeta.title || 'Navigation',
    gid: nativeTable.range.sheetId,
    binding: 'nativeTable',
    apiTableId: nativeTable.tableId,
    apiTableName: nativeTable.name,
    apiRange: nativeTable.range,
    columnProperties: getNativeTableColumnProperties_(nativeTable),
    columns: getNativeTableColumns_(nativeTable),
  };
  return buildTableRowsData_(table).rows.map((row) => row.record || {});
}

function toNavigationItem_(row, userConfig) {
  const pageId = getNavConfigValue_(row, ['page', 'page_id', 'target_page']) || '';
  const page = pageId ? resolveNavPageConfig_(userConfig, pageId) : null;
  const tableName = getNavConfigValue_(row, ['table_name', 'table', 'source_table']) || '';
  const explicitType = getNavConfigValue_(row, ['item_type', 'navigation_type', 'type', 'component_type']);
  const pageType = page ? page.pageType : '';
  const itemType = normalizeNavText_(pageType || explicitType || (pageId || tableName ? 'table' : 'group'));
  const label = getNavConfigValue_(row, ['label', 'item', 'title', 'name']) || (page && page.label) || tableName || pageId || 'Navigation item';
  const id = getNavConfigValue_(row, ['nav_id', 'navigation_id', 'id', 'key', 'config_key']) ||
    normalizeName(label).replace(/[^a-z0-9]+/g, '-');
  const item = {
    id,
    parentId: getNavConfigValue_(row, ['parent_id', 'parent', 'parent_nav_id']) || '',
    label,
    icon: normalizeNavIcon_(getNavConfigValue_(row, ['icon', 'category', 'group']), id),
    pageId,
  };

  if (itemType === 'group' && !page) {
    item.children = [];
    return item;
  }

  if (itemType === 'page' && !page) {
    return applyMissingPageNavigationFallback_(item, row);
  }

  if (itemType === 'overview' || itemType === 'report') {
    item.overviewKey = (page && page.reportName) || getNavConfigValue_(row, ['overview_key', 'report_key', 'report']) || pageId || '';
    item.disabled = !item.overviewKey;
    return item;
  }

  if (itemType === 'external' || itemType === 'embed') {
    item.pageType = itemType;
    item.linkUrl = page && page.linkUrl ? page.linkUrl : getNavConfigValue_(row, ['link_url', 'url', 'external_url', 'embed_url']);
    item.disabled = !item.linkUrl;
    return item;
  }

  if (itemType === 'table' || itemType === 'form') {
    const resolvedTableName = (page && page.tableName) || tableName || label || '';
    const tableKey = (page && page.tableKey) || getTableNameToKey_()[normalizeName(resolvedTableName)] || null;
    item.pageType = itemType;
    item.tableName = resolvedTableName;
    item.tableKey = tableKey;
    item.disabled = isNavTrue_(getNavConfigValue_(row, ['disabled_if_missing_binding']), true) && !tableKey;
    return item;
  }

  item.disabled = true;
  return item;
}

function applyMissingPageNavigationFallback_(item, row) {
  const pageId = item.pageId || '';
  const label = item.label || '';
  const linkUrl = getNavConfigValue_(row, ['link_url', 'url', 'external_url', 'embed_url']);
  const inferredType = inferMissingNavPageType_(label, pageId, linkUrl);

  if (inferredType === 'report') {
    item.overviewKey = getNavConfigValue_(row, ['overview_key', 'report_key', 'report']) || pageId || normalizeName(label);
    item.disabled = !item.overviewKey;
    return item;
  }

  if (inferredType === 'external' || inferredType === 'embed') {
    item.pageType = inferredType;
    item.linkUrl = linkUrl;
    item.disabled = !item.linkUrl;
    return item;
  }

  const explicitTableName = getNavConfigValue_(row, ['table_name', 'table', 'source_table']);
  const tableCandidates = [
    explicitTableName,
    label,
    pageId,
    String(pageId || '').split('.').pop(),
    String(pageId || '').split('_').pop(),
  ].filter(Boolean);
  const tableKey = resolveNavigationTableKey_(tableCandidates);
  item.pageType = 'table';
  item.tableName = explicitTableName || tableCandidates.find((candidate) => resolveNavigationTableKey_([candidate])) || label || pageId;
  item.tableKey = tableKey;
  item.disabled = isNavTrue_(getNavConfigValue_(row, ['disabled_if_missing_binding']), true) && !tableKey;
  return item;
}

function inferMissingNavPageType_(label, pageId, linkUrl) {
  const hint = normalizeNavHeader_([label, pageId].filter(Boolean).join(' '));
  if (hint.indexOf('overview') >= 0 || hint.indexOf('dashboard') >= 0 || hint.indexOf('report') >= 0) return 'report';
  if (linkUrl) return hint.indexOf('embed') >= 0 || hint.indexOf('iframe') >= 0 ? 'embed' : 'external';
  return 'table';
}

function resolveNavigationTableKey_(candidates) {
  const tableNameToKey = getTableNameToKey_();
  for (let index = 0; index < (candidates || []).length; index += 1) {
    const candidate = candidates[index];
    const normalized = normalizeName(candidate);
    if (tableNameToKey[normalized]) return tableNameToKey[normalized];
  }
  return null;
}

function resolveNavPageConfig_(userConfig, pageId) {
  const pages = userConfig && userConfig.pages && userConfig.pages.byId ? userConfig.pages.byId : {};
  if (pages[pageId]) return pages[pageId];
  const normalizedPageId = normalizeNavHeader_(pageId);
  return Object.keys(pages).reduce((found, key) => (
    found || (normalizeNavHeader_(key) === normalizedPageId ? pages[key] : null)
  ), null);
}

function getNavSortOrder_(row) {
  return Number(getNavConfigValue_(row, ['sort_order', 'sort', 'order']) || 0);
}

function normalizeNavHeader_(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g, '_');
}

function normalizeNavText_(value) {
  return String(value || '').trim().toLowerCase();
}

function getNavConfigValue_(row, names) {
  for (let index = 0; index < names.length; index += 1) {
    const name = names[index];
    if (Object.prototype.hasOwnProperty.call(row, name) && row[name] !== '') return row[name];
  }

  const normalizedNames = names.map(normalizeNavHeader_);
  const keys = Object.keys(row || {});
  for (let index = 0; index < keys.length; index += 1) {
    if (normalizedNames.includes(normalizeNavHeader_(keys[index])) && row[keys[index]] !== '') {
      return row[keys[index]];
    }
  }
  return '';
}

function normalizeNavIcon_(icon, id) {
  const text = normalizeNavText_(icon);
  if (['overview', 'operation', 'business', 'settings'].includes(text)) return text;

  const normalizedId = normalizeNavText_(id);
  if (normalizedId === 'overview') return 'overview';
  if (normalizedId === 'operation') return 'operation';
  if (normalizedId === 'business_setup') return 'business';
  if (normalizedId === 'system_setup') return 'settings';
  return icon || '';
}

function isNavTrue_(value, defaultValue) {
  if (value === '' || value == null) return defaultValue;
  return ['true', 'yes', 'y', '1'].includes(normalizeNavText_(value));
}
