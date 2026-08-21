/**
 * Overview/report aggregation service.
 *
 * Builds summary cards, grouped data, chart/table payloads, and filtered
 * overview responses. Keep reporting calculations here; ClientViews.html
 * should only render the payload.
 */
function getOverviewData(overviewKey, forceRefresh, selectedFilters) {
  assertSpreadsheetAccess_();
  const normalizedFilters = normalizeOverviewSelectedFilters_(selectedFilters);
  const cacheKey = cacheKey_('overview', `${overviewKey}:${serializeOverviewFilters_(normalizedFilters)}`);
  if (!forceRefresh) {
    const cached = getCached_(cacheKey);
    if (cached) return sanitizeClientPayload_(cached);
  }

  const data = buildOverviewData_(overviewKey, forceRefresh, normalizedFilters);
  setCached_(cacheKey, data);
  return sanitizeClientPayload_(data);
}

function buildOverviewData_(overviewKey, forceRefresh, selectedFilters) {
  const rows = getOverviewConfigRows_(!!forceRefresh)
    .filter((row) => normalizeName(getOverviewConfigValue_(row, ['overview_key', 'report', 'report_key'])) === normalizeName(overviewKey))
    .filter((row) => normalizeName(getOverviewConfigValue_(row, ['status']) || 'Active') === 'active')
    .sort((a, b) => Number(getOverviewConfigValue_(a, ['sort_order']) || 0) - Number(getOverviewConfigValue_(b, ['sort_order']) || 0));
  if (!rows.length) return emptyOverviewData_(overviewKey);

  const page = rows.find((row) => normalizeReportItemType_(getOverviewConfigValue_(row, ['item_type'])) === 'page') || {};
  const dataCache = { __forceRefresh: !!forceRefresh };
  const reportFilters = hydrateOverviewReportFilters_(
    getOverviewReportFilters_(rows),
    dataCache,
    forceRefresh,
    selectedFilters
  );
  const cards = rows
    .filter((row) => normalizeReportItemType_(getOverviewConfigValue_(row, ['item_type'])) === 'card')
    .map((row) => {
      const definition = applyOverviewReportFilters_(getOverviewMetricDefinition_(row, dataCache), reportFilters, dataCache, forceRefresh);
      return metric_(getOverviewConfigValue_(row, ['label']) || definition.label || definition.item_id, evaluateOverviewMetric_(definition, dataCache, forceRefresh));
    });
  const groups = rows
    .filter((row) => normalizeReportItemType_(getOverviewConfigValue_(row, ['item_type'])) === 'group')
    .map((row) => {
      const definition = applyOverviewReportFilters_(getOverviewMetricDefinition_(row, dataCache), reportFilters, dataCache, forceRefresh);
      return group_(getOverviewConfigValue_(row, ['label']) || definition.label || definition.field || definition.item_id, evaluateOverviewGroup_(definition, dataCache, forceRefresh));
    });
  const tableRow = rows.find((row) => normalizeReportItemType_(getOverviewConfigValue_(row, ['item_type'])) === 'table');

  return {
    key: overviewKey,
    title: getOverviewConfigValue_(page, ['title']) || overviewKey,
    subtitle: getOverviewConfigValue_(page, ['subtitle']) || '',
    filters: reportFilters,
    cards,
    groups,
    table: tableRow ? buildOverviewPreviewTable_(tableRow, dataCache, forceRefresh, reportFilters) : null,
  };
}

function metric_(label, value) {
  return { label, value };
}

function group_(label, values) {
  return { label, values };
}

function getOverviewConfigRows_(forceRefresh) {
  return readUserConfigRecords_('reportPlanning', !!forceRefresh);
}

function getMetricConfigRows_(forceRefresh) {
  return readUserConfigRecords_('metric', !!forceRefresh);
}

function getOverviewKeys_() {
  const keys = getOverviewConfigRows_(false)
    .filter((row) => normalizeName(getOverviewConfigValue_(row, ['status']) || 'Active') === 'active')
    .map((row) => getOverviewConfigValue_(row, ['overview_key', 'report', 'report_key']))
    .filter(Boolean);
  return Array.from(new Set(keys));
}

function normalizeOverviewHeader_(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g, '_');
}

function getSupportedReportItemTypes_() {
  return [
    'page', 'overview_page', 'dashboard_page', 'report_page',
    'filter', 'report_filter', 'filter_bar', 'quick_filter',
    'select_filter', 'dropdown_filter', 'date_filter', 'search_filter',
    'card', 'metric_card', 'kpi_card', 'summary_card', 'stat_card',
    'number_card', 'progress_card', 'status_card', 'tile',
    'group', 'group_chart', 'bar_chart', 'column_chart', 'pie_chart',
    'donut_chart', 'line_chart', 'trend_chart', 'distribution_chart',
    'breakdown_chart', 'status_breakdown',
    'table', 'data_table', 'detail_table', 'preview_table',
    'snapshot_table', 'list', 'record_list', 'pivot_table',
  ];
}

function getReportItemTypeAliasMap_() {
  return {
    overview: 'page',
    overview_page: 'page',
    main_page: 'page',
    dashboard: 'page',
    dashboard_page: 'page',
    report: 'page',
    report_page: 'page',
    filter_bar: 'filter',
    report_filter: 'filter',
    quick_filter: 'filter',
    select_filter: 'filter',
    dropdown_filter: 'filter',
    date_filter: 'filter',
    search_filter: 'filter',
    slicer: 'filter',
    metric: 'card',
    metric_card: 'card',
    kpi: 'card',
    kpi_card: 'card',
    card_metric: 'card',
    summary_card: 'card',
    stat: 'card',
    stat_card: 'card',
    number_card: 'card',
    progress_card: 'card',
    status_card: 'card',
    tile: 'card',
    value: 'card',
    number: 'card',
    chart: 'group',
    group_chart: 'group',
    bar_chart: 'group',
    column_chart: 'group',
    pie_chart: 'group',
    donut_chart: 'group',
    line_chart: 'group',
    trend_chart: 'group',
    distribution: 'group',
    distribution_chart: 'group',
    breakdown_chart: 'group',
    status_breakdown: 'group',
    group_count: 'group',
    group_by: 'group',
    segment: 'group',
    breakdown: 'group',
    list: 'table',
    data_table: 'table',
    detail_table: 'table',
    preview_table: 'table',
    table_preview: 'table',
    records: 'table',
    record_list: 'table',
    rows: 'table',
    snapshot: 'table',
    snapshot_table: 'table',
    pivot: 'table',
    pivot_table: 'table',
  };
}

function normalizeReportItemType_(value) {
  const type = normalizeOverviewHeader_(value);
  return getReportItemTypeAliasMap_()[type] || type;
}

function getSupportedMetricTypes_() {
  return [
    'row_count', 'count', 'count_rows', 'record_count',
    'count_matching', 'count_if', 'filtered_count',
    'group_count', 'count_by', 'group_by', 'breakdown',
    'unique_count', 'count_distinct', 'distinct_count',
    'sum', 'total',
    'average', 'avg', 'mean',
    'median', 'min', 'max',
    'count_not_blank', 'not_blank_count',
    'count_blank', 'blank_count',
    'formula', 'calc', 'calculated',
  ];
}

function getMetricTypeAliasMap_() {
  return {
    count: 'row_count',
    count_rows: 'row_count',
    count_row: 'row_count',
    record_count: 'row_count',
    records: 'row_count',
    rows: 'row_count',
    rowcount: 'row_count',
    count_all: 'row_count',
    count_by: 'group_count',
    count_group: 'group_count',
    grouped_count: 'group_count',
    group: 'group_count',
    group_by: 'group_count',
    breakdown: 'group_count',
    distribution: 'group_count',
    count_distinct: 'unique_count',
    distinct_count: 'unique_count',
    distinct: 'unique_count',
    unique: 'unique_count',
    unique_rows: 'unique_count',
    total: 'sum',
    subtotal: 'sum',
    add: 'sum',
    avg: 'average',
    mean: 'average',
    avg_value: 'average',
    median_value: 'median',
    calculated: 'formula',
    calc: 'formula',
    expression: 'formula',
    custom_formula: 'formula',
    count_if: 'count_matching',
    countif: 'count_matching',
    filtered_count: 'count_matching',
    matching_count: 'count_matching',
    not_blank_count: 'count_not_blank',
    count_filled: 'count_not_blank',
    filled_count: 'count_not_blank',
    blank_count: 'count_blank',
    count_empty: 'count_blank',
    empty_count: 'count_blank',
  };
}

function normalizeMetricType_(value) {
  const type = normalizeOverviewHeader_(value || 'row_count');
  return getMetricTypeAliasMap_()[type] || type || 'row_count';
}

function getOverviewConfigValue_(row, names) {
  for (let index = 0; index < names.length; index += 1) {
    const name = names[index];
    if (Object.prototype.hasOwnProperty.call(row, name) && row[name] !== '') return row[name];
  }

  const normalizedNames = names.map(normalizeOverviewHeader_);
  const keys = Object.keys(row || {});
  for (let index = 0; index < keys.length; index += 1) {
    if (normalizedNames.includes(normalizeOverviewHeader_(keys[index])) && row[keys[index]] !== '') {
      return row[keys[index]];
    }
  }
  return '';
}

function getOverviewMetricDefinition_(row, dataCache) {
  if (!dataCache.__metricConfigByKey) {
    dataCache.__metricConfigByKey = getMetricConfigRows_(dataCache.__forceRefresh)
      .filter((metric) => normalizeName(getOverviewConfigValue_(metric, ['status']) || 'Active') === 'active')
      .reduce((acc, metric) => {
        acc[normalizeName(getOverviewConfigValue_(metric, ['metric_key']))] = metric;
        return acc;
      }, {});
  }

  const rowMetricKey = getOverviewConfigValue_(row, ['metric_key']);
  const itemId = getOverviewConfigValue_(row, ['item_id', 'item']);
  const metric = dataCache.__metricConfigByKey[normalizeName(rowMetricKey)] || {};
  return {
    key: rowMetricKey || getOverviewConfigValue_(metric, ['metric_key']) || itemId,
    label: getOverviewConfigValue_(metric, ['label']) || getOverviewConfigValue_(row, ['label']) || getOverviewConfigValue_(row, ['title']) || '',
    table_name: getOverviewConfigValue_(metric, ['source_table']) || getOverviewConfigValue_(row, ['table_name']),
    metric_type: normalizeMetricType_(getOverviewConfigValue_(metric, ['metric_type']) || getOverviewConfigValue_(row, ['metric_type'])),
    field: getOverviewConfigValue_(metric, ['field_expression']) || getOverviewConfigValue_(row, ['field']),
    match_pattern: getOverviewConfigValue_(metric, ['filter_expression']) || getOverviewConfigValue_(row, ['match_pattern']),
    format: getOverviewConfigValue_(metric, ['format']) || '',
    limit: getOverviewConfigValue_(row, ['limit']),
    title: getOverviewConfigValue_(row, ['title']),
    item_id: itemId,
  };
}

function getOverviewReportFilters_(rows) {
  return rows
    .filter((row) => normalizeReportItemType_(getOverviewConfigValue_(row, ['item_type'])) === 'filter')
    .map((row) => ({
      key: getOverviewConfigValue_(row, ['item_id', 'item']) || getOverviewConfigValue_(row, ['label']) || '',
      label: getOverviewConfigValue_(row, ['label']) || getOverviewConfigValue_(row, ['title']) || 'Filter',
      title: getOverviewConfigValue_(row, ['title']) || '',
      table_name: getOverviewConfigValue_(row, ['table_name']),
      field: getOverviewConfigValue_(row, ['field']),
      match_pattern: getOverviewConfigValue_(row, ['match_pattern']),
      input_type: normalizeOverviewHeader_(getOverviewConfigValue_(row, ['filter_type', 'control_type', 'input_type']) || 'select'),
      default_value: getOverviewConfigValue_(row, ['default_value', 'value']),
    }));
}

function hydrateOverviewReportFilters_(filters, dataCache, forceRefresh, selectedFilters) {
  return (filters || []).map((filter) => {
    const value = getOverviewSelectedFilterValue_(filter, selectedFilters);
    return Object.assign({}, filter, {
      value,
      options: buildOverviewFilterOptions_(filter, dataCache, forceRefresh),
    });
  });
}

function getOverviewSelectedFilterValue_(filter, selectedFilters) {
  const key = normalizeName(filter.key || filter.label || filter.field);
  const value = selectedFilters && Object.prototype.hasOwnProperty.call(selectedFilters, key)
    ? selectedFilters[key]
    : filter.default_value;
  if (Array.isArray(value)) return value.filter(Boolean);
  return value == null ? '' : String(value).trim();
}

function buildOverviewFilterOptions_(filter, dataCache, forceRefresh) {
  if (!filter.table_name || !filter.field) return [];
  const data = getOverviewTableData_(filter.table_name, dataCache, forceRefresh);
  if (!data) return [];
  const counts = countBy_(data.rows, filter.field);
  return Object.keys(counts)
    .filter((value) => value && value !== 'Blank')
    .sort((a, b) => a.localeCompare(b, 'vi', { numeric: true }))
    .map((value) => ({ label: value, value, count: counts[value] }));
}

function applyOverviewReportFilters_(definition, filters, dataCache, forceRefresh) {
  const data = getOverviewTableData_(definition.table_name, dataCache, forceRefresh);
  const headers = data && data.headers ? data.headers : [];
  const activeFilters = (filters || []).filter(isOverviewFilterActive_);
  const missingFilters = activeFilters.filter((filter) => !overviewFilterAppliesToHeaders_(filter, headers));
  const extraFilters = activeFilters
    .filter((filter) => overviewFilterAppliesToHeaders_(filter, headers))
    .map((filter) => getOverviewFilterExpression_(filter))
    .filter(Boolean);
  if (missingFilters.length) extraFilters.push('__overview_filter_scope__~__no_match__');
  if (!extraFilters.length) return definition;
  return Object.assign({}, definition, {
    match_pattern: [definition.match_pattern].concat(extraFilters).filter(Boolean).join(';'),
  });
}

function isOverviewFilterActive_(filter) {
  if (!filter) return false;
  if (filter.match_pattern) return true;
  if (Array.isArray(filter.value)) return filter.value.some((value) => String(value || '').trim() !== '');
  return String(filter.value || '').trim() !== '';
}

function overviewFilterAppliesToHeaders_(filter, headers) {
  if (!filter.field) return false;
  const normalizedField = normalizeOverviewHeader_(filter.field);
  return (headers || []).some((header) => normalizeOverviewHeader_(header) === normalizedField);
}

function getOverviewFilterExpression_(filter) {
  if (filter.match_pattern) return filter.match_pattern;
  if (!filter.field || filter.value === '' || filter.value == null) return '';
  const values = Array.isArray(filter.value) ? filter.value : [filter.value];
  const pattern = values
    .map((value) => String(value || '').trim())
    .filter(Boolean)
    .map(escapeRegex_)
    .join('|');
  return pattern ? `${filter.field}~^(${pattern})$` : '';
}

function normalizeOverviewSelectedFilters_(selectedFilters) {
  return Object.keys(selectedFilters || {}).reduce((acc, key) => {
    const value = selectedFilters[key];
    if (value == null || value === '' || value === '__all__') return acc;
    if (Array.isArray(value)) {
      const values = value.filter((item) => item != null && item !== '' && item !== '__all__');
      if (values.length) acc[normalizeName(key)] = values;
      return acc;
    }
    acc[normalizeName(key)] = String(value);
    return acc;
  }, {});
}

function serializeOverviewFilters_(filters) {
  return Object.keys(filters || {})
    .sort()
    .map((key) => `${key}=${Array.isArray(filters[key]) ? filters[key].join('|') : filters[key]}`)
    .join('&');
}

function escapeRegex_(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function evaluateOverviewMetric_(definition, dataCache, forceRefresh) {
  const data = getOverviewTableData_(definition.table_name, dataCache, forceRefresh);
  if (!data) return 0;

  const metricType = normalizeMetricType_(definition.metric_type);
  const rows = filterRowsByConfig_(data.rows, definition.match_pattern);
  if (metricType === 'formula') {
    return evaluateMetricFormula_(rows, definition.field);
  }
  if (metricType === 'sum') {
    return formatMetricNumber_(sumField_(rows, definition.field));
  }
  if (metricType === 'average' || metricType === 'avg') {
    return formatMetricNumber_(averageField_(rows, definition.field));
  }
  if (metricType === 'median') {
    return formatMetricNumber_(medianField_(rows, definition.field));
  }
  if (metricType === 'min') {
    return formatMetricNumber_(minField_(rows, definition.field));
  }
  if (metricType === 'max') {
    return formatMetricNumber_(maxField_(rows, definition.field));
  }
  if (metricType === 'count_matching') {
    return countMatching_(data.rows, definition.field, patternFromConfig_(definition.match_pattern));
  }
  if (metricType === 'count_not_blank') {
    return countField_(rows, definition.field);
  }
  if (metricType === 'count_blank') {
    return Math.max(0, rows.length - countField_(rows, definition.field));
  }
  if (metricType === 'group_count') {
    return Object.keys(countBy_(rows, definition.field)).filter((key) => key !== 'Blank').length;
  }
  if (metricType === 'unique_count') {
    return uniqueCount_(rows, definition.field);
  }
  return rows.length;
}

function evaluateOverviewGroup_(definition, dataCache, forceRefresh) {
  const data = getOverviewTableData_(definition.table_name, dataCache, forceRefresh);
  if (!data) return {};
  return countBy_(filterRowsByConfig_(data.rows, definition.match_pattern), definition.field);
}

function buildOverviewPreviewTable_(row, dataCache, forceRefresh, reportFilters) {
  const definition = applyOverviewReportFilters_(getOverviewMetricDefinition_(row, dataCache), reportFilters, dataCache, forceRefresh);
  const data = getOverviewTableData_(definition.table_name, dataCache, forceRefresh);
  const limit = Number(getOverviewConfigValue_(row, ['limit']) || 8);
  if (!data) {
    return { title: getOverviewConfigValue_(row, ['title']) || definition.label || definition.item_id, headers: [], rows: [] };
  }
  return {
    title: getOverviewConfigValue_(row, ['title']) || definition.label || data.table.name,
    headers: data.headers,
    rows: filterRowsByConfig_(data.rows, definition.match_pattern).slice(0, limit > 0 ? limit : 8),
  };
}

function getOverviewTableData_(tableName, dataCache, forceRefresh) {
  if (!dataCache.__tableNameToKey) dataCache.__tableNameToKey = getTableNameToKey_(forceRefresh);
  const cleanTableName = typeof stripTableLabelIcon_ === 'function' ? stripTableLabelIcon_(tableName) : tableName;
  const tableKey = dataCache.__tableNameToKey[normalizeName(tableName)]
    || dataCache.__tableNameToKey[normalizeName(cleanTableName)];
  if (!tableKey) return null;
  if (!dataCache[tableKey]) dataCache[tableKey] = getTableData(tableKey, forceRefresh);
  return dataCache[tableKey];
}

function patternFromConfig_(pattern) {
  try {
    return new RegExp(pattern || '.', 'i');
  } catch (error) {
    return /./i;
  }
}

function filterRowsByConfig_(rows, filterExpression) {
  const filters = parseFilterExpression_(filterExpression);
  if (!filters.length) return rows;
  return rows.filter((row) => filters.every((filter) => filter.pattern.test(getRecordFieldValue_(row.record, filter.field) || '')));
}

function parseFilterExpression_(filterExpression) {
  return String(filterExpression || '').split(/[;&]/)
    .map((part) => part.trim())
    .filter((part) => part.includes('~'))
    .map((part) => {
      const separatorIndex = part.indexOf('~');
      return {
        field: part.slice(0, separatorIndex).trim(),
        pattern: patternFromConfig_(part.slice(separatorIndex + 1).trim()),
      };
    });
}

function evaluateMetricFormula_(rows, expression) {
  const formula = String(expression || '').trim();
  if (!formula) return 0;

  const value = formula
    .replace(/sum\(([^)]*)\)/gi, (_, field) => String(sumField_(rows, field)))
    .replace(/avg\(([^)]*)\)/gi, (_, field) => String(averageField_(rows, field)))
    .replace(/average\(([^)]*)\)/gi, (_, field) => String(averageField_(rows, field)))
    .replace(/median\(([^)]*)\)/gi, (_, field) => String(medianField_(rows, field)))
    .replace(/min\(([^)]*)\)/gi, (_, field) => String(minField_(rows, field)))
    .replace(/max\(([^)]*)\)/gi, (_, field) => String(maxField_(rows, field)))
    .replace(/count\(([^)]*)\)/gi, (_, field) => String(countField_(rows, field)))
    .replace(/unique\(([^)]*)\)/gi, (_, field) => String(uniqueCount_(rows, field)));

  if (!/^[\d+\-*/().\s]+$/.test(value)) return 0;
  try {
    return formatMetricNumber_(Function(`"use strict"; return (${value});`)());
  } catch (error) {
    return 0;
  }
}

function sumField_(rows, fieldExpression) {
  return getFieldNames_(fieldExpression).reduce((total, field) => (
    total + rows.reduce((sum, row) => {
      const value = parseMetricNumber_(getRecordFieldValue_(row.record, field));
      return Number.isFinite(value) ? sum + value : sum;
    }, 0)
  ), 0);
}

function averageField_(rows, fieldExpression) {
  const values = getFieldNames_(fieldExpression).flatMap((field) => (
    rows.map((row) => parseMetricNumber_(getRecordFieldValue_(row.record, field))).filter((value) => Number.isFinite(value))
  ));
  if (!values.length) return 0;
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function minField_(rows, fieldExpression) {
  const values = getFieldNames_(fieldExpression).flatMap((field) => (
    rows.map((row) => parseMetricNumber_(getRecordFieldValue_(row.record, field))).filter((value) => Number.isFinite(value))
  ));
  return values.length ? Math.min.apply(null, values) : 0;
}

function medianField_(rows, fieldExpression) {
  const values = getFieldNames_(fieldExpression).flatMap((field) => (
    rows.map((row) => parseMetricNumber_(getRecordFieldValue_(row.record, field))).filter((value) => Number.isFinite(value))
  )).sort((a, b) => a - b);
  if (!values.length) return 0;
  const middle = Math.floor(values.length / 2);
  return values.length % 2 ? values[middle] : (values[middle - 1] + values[middle]) / 2;
}

function maxField_(rows, fieldExpression) {
  const values = getFieldNames_(fieldExpression).flatMap((field) => (
    rows.map((row) => parseMetricNumber_(getRecordFieldValue_(row.record, field))).filter((value) => Number.isFinite(value))
  ));
  return values.length ? Math.max.apply(null, values) : 0;
}

function countField_(rows, fieldExpression) {
  const fields = getFieldNames_(fieldExpression);
  if (!fields.length) return rows.length;
  return rows.filter((row) => fields.some((field) => getRecordFieldValue_(row.record, field) !== '')).length;
}

function getFieldNames_(fieldExpression) {
  return String(fieldExpression || '').split(/[,+]/).map((field) => field.trim()).filter(Boolean);
}

function getRecordFieldValue_(record, field) {
  if (!record || !field) return '';
  if (Object.prototype.hasOwnProperty.call(record, field)) return record[field];

  const normalizedField = normalizeOverviewHeader_(field);
  const keys = Object.keys(record);
  for (let index = 0; index < keys.length; index += 1) {
    if (normalizeOverviewHeader_(keys[index]) === normalizedField) return record[keys[index]];
  }
  return '';
}

function parseMetricNumber_(value) {
  if (value == null || value === '') return NaN;
  const normalized = String(value).replace(/,/g, '').trim();
  const number = Number(normalized);
  return Number.isFinite(number) ? number : NaN;
}

function formatMetricNumber_(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return 0;
  return Number.isInteger(number) ? number : Number(number.toFixed(1));
}

function emptyOverviewData_(overviewKey) {
  return {
    key: overviewKey,
    title: overviewKey,
    subtitle: '',
    cards: [],
    groups: [],
    table: {
      title: '',
      headers: [],
      rows: [],
    },
  };
}

function countBy_(rows, header) {
  return rows.reduce((acc, row) => {
    const value = getRecordFieldValue_(row.record, header) || 'Blank';
    acc[value] = (acc[value] || 0) + 1;
    return acc;
  }, {});
}

function countMatching_(rows, header, pattern) {
  return rows.filter((row) => pattern.test(getRecordFieldValue_(row.record, header) || '')).length;
}

function uniqueCount_(rows, header) {
  return Object.keys(countBy_(rows, header)).filter((key) => key !== 'Blank').length;
}
