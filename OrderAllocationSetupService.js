/**
 * Order-allocation Setup preview and batch creation.
 * Server writes are rebuilt from the clicked lsx_line row, never client rows.
 */
const ORDER_ALLOCATION_SETUP_NAV_ID_ = 'operation.lập_kế_hoạch.phân_bổ_đơn_hàng';
const ORDER_ALLOCATION_ROUTINGS_ = ['Sơ chế', 'Tinh chế', 'Lắp ráp', 'Nguội', 'Sơn', 'Bao bì'];

function getOrderAllocationSetupPreview(payload) {
  assertSpreadsheetAccess_();
  const context = buildOrderAllocationSetupContext_(payload);
  return sanitizeClientPayload_({
    lsxLineId: context.lsxLineId,
    bomHeader: context.bomHeader,
    sourceRowNumber: context.sourceRowNumber,
    materialRows: context.materialRows,
    resourceRows: context.resourceRows,
    materialStats: context.materialStats,
    resourceStats: context.resourceStats,
    complete: context.complete,
  });
}

function saveOrderAllocationSetup(payload) {
  assertSpreadsheetAccess_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const context = buildOrderAllocationSetupContext_(payload);
    const expectedBomHeader = String(payload && payload.bomHeader || '').trim();
    if (expectedBomHeader && expectedBomHeader !== context.bomHeader) {
      throw new Error('BOM header đã thay đổi sau khi mở Setup. Hãy đóng popup và kiểm tra lại dữ liệu.');
    }
    const materialImported = importOrderAllocationRecords_(context.materialTable, context.materialRecords);
    const resourceImported = importOrderAllocationRecords_(context.resourceTable, context.resourceRecords);
    SpreadsheetApp.flush();
    invalidateTableDataCache_(context.materialTable.key);
    invalidateTableDataCache_(context.resourceTable.key);
    return sanitizeClientPayload_({
      ok: true,
      lsxLineId: context.lsxLineId,
      bomHeader: context.bomHeader,
      materialImported,
      resourceImported,
      alreadyComplete: materialImported === 0 && resourceImported === 0,
      materialTableKey: context.materialTable.key,
      resourceTableKey: context.resourceTable.key,
    });
  } finally {
    lock.releaseLock();
  }
}

function buildOrderAllocationSetupContext_(payload) {
  const request = payload || {};
  if (request.navId !== ORDER_ALLOCATION_SETUP_NAV_ID_) {
    throw new Error('Setup chỉ được phép chạy tại NAV Phân bổ đơn hàng.');
  }
  const tableKey = String(request.tableKey || '').trim();
  const rowNumber = Number(request.rowNumber || 0);
  if (!tableKey || !rowNumber) throw new Error('Không xác định được dòng LSX đã chọn.');

  const lineTable = getOrderAllocationTable_(['lsx_line', 'LSX line'], false);
  if (lineTable.key !== tableKey) throw new Error('Setup phải được mở trực tiếp từ một record của bảng lsx_line.');
  const lineRow = readSavedRecordRow_(lineTable, rowNumber);
  if (!lineRow || !lineRow.record) throw new Error('Không đọc được dòng lsx_line đã chọn.');
  const lsxLineId = String(recordValue_(lineRow.record, ['LSX line ID', 'lsx_line_id']) || '').trim();
  const bomHeader = String(recordValue_(lineRow.record, ['BOM header', 'bom_header']) || '').trim();
  if (!lsxLineId) throw new Error('Dòng đã chọn chưa có LSX line ID.');
  if (!bomHeader) throw new Error('Dòng đã chọn chưa có BOM header.');
  const expectedLineId = String(request.lsxLineId || '').trim();
  if (expectedLineId && expectedLineId !== lsxLineId) {
    throw new Error('LSX line đã thay đổi sau khi mở Setup. Hãy chọn lại record.');
  }

  const bomLineTable = getOrderAllocationReadTable_(
    ['bom_line', 'BOM line', '↳ bom_line'],
    lineTable.spreadsheetId || SPREADSHEET_ID
  );
  const materialTable = getOrderAllocationTable_(['lsx_material_requirement'], false);
  const resourceTable = getOrderAllocationTable_(['lsx_resource_requirement'], false);
  assertTableWritable_(materialTable);
  assertTableWritable_(resourceTable);
  const bomLines = readBomLinesForHeader_(bomLineTable, bomHeader);
  if (!bomLines.matchedCount) throw new Error(`Không tìm thấy BOM line nào cho BOM header ${bomHeader}.`);

  const materialHeaders = getOrderAllocationHeaders_(materialTable);
  const resourceHeaders = getOrderAllocationHeaders_(resourceTable);
  const materialBomLineHeader = requireOrderAllocationHeader_(materialHeaders, ['BOM line', 'bom_line']);
  const materialLsxLineHeader = requireOrderAllocationHeader_(materialHeaders, ['LSX line ID', 'lsx_line_id']);
  const resourceLsxLineHeader = requireOrderAllocationHeader_(resourceHeaders, ['LSX line ID', 'lsx_line_id']);
  const resourceRoutingHeader = requireOrderAllocationHeader_(resourceHeaders, ['Routing']);
  const existingMaterial = readExistingOrderAllocationValues_(
    materialTable,
    ['LSX line ID', 'lsx_line_id'],
    ['BOM line', 'bom_line'],
    lsxLineId
  );
  const existingResource = readExistingOrderAllocationValues_(
    resourceTable,
    ['LSX line ID', 'lsx_line_id'],
    ['Routing'],
    lsxLineId
  );
  const missingBomLineIds = bomLines.eligibleIds.filter((bomLineId) => !existingMaterial[bomLineId]);
  const missingRoutings = ORDER_ALLOCATION_ROUTINGS_.filter((routing) => !existingResource[routing]);
  const materialRecords = missingBomLineIds.map((bomLineId) => ({
    [materialBomLineHeader]: bomLineId,
    [materialLsxLineHeader]: lsxLineId,
  }));
  const resourceRecords = missingRoutings.map((routing) => ({
    [resourceLsxLineHeader]: lsxLineId,
    [resourceRoutingHeader]: routing,
  }));

  return {
    lsxLineId,
    bomHeader,
    sourceRowNumber: rowNumber,
    materialTable,
    resourceTable,
    materialRecords,
    resourceRecords,
    materialRows: missingBomLineIds.map((bomLineId) => ({ bomLine: bomLineId, lsxLineId })),
    resourceRows: missingRoutings.map((routing) => ({ lsxLineId, routing })),
    materialStats: {
      matched: bomLines.matchedCount,
      excludedClusters: bomLines.excludedCount,
      eligible: bomLines.eligibleIds.length,
      existing: bomLines.eligibleIds.length - missingBomLineIds.length,
      missing: missingBomLineIds.length,
    },
    resourceStats: {
      total: ORDER_ALLOCATION_ROUTINGS_.length,
      existing: ORDER_ALLOCATION_ROUTINGS_.length - missingRoutings.length,
      missing: missingRoutings.length,
    },
    complete: missingBomLineIds.length === 0 && missingRoutings.length === 0,
  };
}

function getOrderAllocationTable_(aliases, forceRefresh) {
  const nameToKey = getTableNameToKey_(!!forceRefresh);
  for (let index = 0; index < (aliases || []).length; index += 1) {
    const alias = aliases[index];
    const key = nameToKey[normalizeName(alias)] || nameToKey[normalizeConfigHeader_(alias)];
    if (key) return getTableOrThrow_(key, !!forceRefresh);
  }
  if (!forceRefresh) return getOrderAllocationTable_(aliases, true);
  throw new Error(`Không tìm thấy bảng ${String((aliases || [])[0] || '')} trong Table index.`);
}

function getOrderAllocationReadTable_(aliases, spreadsheetId) {
  try {
    return getOrderAllocationTable_(aliases, false);
  } catch (error) {
    const appSpreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const indexRows = [
      readTableIndexRows_(appSpreadsheet, TABLE_INDEX_ALL_SHEET_NAME, TABLE_INDEX_ALL_SHEET_ID, {}),
      readTableIndexRows_(appSpreadsheet, TABLE_INDEX_SHEET_NAME, TABLE_INDEX_SHEET_ID, {}),
    ].flat();
    const lookupNames = (aliases || []).map(normalizeTableLookupName_);
    const matchingRows = indexRows.filter((row) => (
      lookupNames.indexOf(normalizeTableLookupName_(row.tableName)) >= 0
    ));
    const indexRow = matchingRows.find((row) => row.parsedLink && row.parsedLink.spreadsheetId) || matchingRows[0];
    if (indexRow) {
      const indexedTable = resolveProcessTableFromIndexRow_(
        appSpreadsheet,
        indexRow,
        buildTableIndexSourceHints_(indexRows),
        false,
        {}
      );
      if (indexedTable && !indexedTable.accessFallbackReason) return indexedTable;
    }

    const registry = getNativeSpreadsheetTablesById_(spreadsheetId || SPREADSHEET_ID, false);
    const nativeLookupNames = (aliases || []).reduce((names, alias) => names.concat(getNativeTableLookupNames_(alias)), []);
    const nativeTable = (registry.tables || []).find((candidate) => (
      nativeLookupNames.indexOf(normalizeTableLookupName_(candidate.name)) >= 0
    ));
    if (!nativeTable) throw error;
    const sheetMeta = (registry.sheetsById || {})[String(nativeTable.range.sheetId)] || {};
    return createNativeProcessTableConfig_({
      tableName: String((aliases || [])[0] || nativeTable.name),
      nativeTable,
      sheetMeta,
      spreadsheetId: spreadsheetId || SPREADSHEET_ID,
      source: 'Order allocation Setup',
      inherited: (spreadsheetId || SPREADSHEET_ID) !== SPREADSHEET_ID,
      readOnly: true,
    });
  }
}

function getOrderAllocationHeaders_(table) {
  if (isNativeReadableTable_(table)) return getNativeHeaderList_(table, readNativeTableHeaderRow_(table));
  return getSystemTableLayout_(getSheet_(table), table).headers;
}

function requireOrderAllocationHeader_(headers, aliases) {
  const normalizedAliases = (aliases || []).map(normalizeUserConfigKey_);
  const header = (headers || []).find((candidate) => normalizedAliases.indexOf(normalizeUserConfigKey_(candidate)) >= 0);
  if (!header) throw new Error(`Thiếu cột bắt buộc: ${String((aliases || [])[0] || '')}.`);
  return header;
}

function readBomLinesForHeader_(table, bomHeader) {
  const rows = readOrderAllocationColumns_(table, [
    ['BOM line ID', 'bom_line_id', 'BOM line'],
    ['BOM header', 'bom_header'],
    ['Level'],
  ]);
  const unique = {};
  return rows.reduce((result, values) => {
    if (String(values[1] || '').trim() !== bomHeader) return result;
    result.matchedCount += 1;
    const bomLineId = String(values[0] || '').trim();
    const isCluster = normalizeUserConfigKey_(values[2]) === normalizeUserConfigKey_('Cụm');
    if (isCluster) {
      result.excludedCount += 1;
      return result;
    }
    if (!bomLineId || unique[bomLineId]) return result;
    unique[bomLineId] = true;
    result.eligibleIds.push(bomLineId);
    return result;
  }, { matchedCount: 0, excludedCount: 0, eligibleIds: [] });
}

function readExistingOrderAllocationValues_(table, parentAliases, valueAliases, parentValue) {
  const rows = readOrderAllocationColumns_(table, [parentAliases, valueAliases]);
  return rows.reduce((existing, values) => {
    if (String(values[0] || '').trim() !== parentValue) return existing;
    const value = String(values[1] || '').trim();
    if (value) existing[value] = true;
    return existing;
  }, {});
}

function readOrderAllocationColumns_(table, aliasGroups) {
  const headers = getOrderAllocationHeaders_(table);
  const resolvedHeaders = (aliasGroups || []).map((aliases) => requireOrderAllocationHeader_(headers, aliases));
  if (!isNativeReadableTable_(table)) {
    return (buildTableRowsData_(table, { limit: 50000, fromStart: true }).rows || [])
      .map((row) => resolvedHeaders.map((header) => String((row.record || {})[header] || '').trim()));
  }

  const indexes = resolvedHeaders.map((header) => headers.indexOf(header));
  const firstIndex = Math.min.apply(null, indexes);
  const lastIndex = Math.max.apply(null, indexes);
  const sheet = getNativeSheet_(table);
  const range = table.apiRange || {};
  const dataStartRow = Number(range.startRowIndex || 0) + 2;
  const tableEndRow = Number(range.endRowIndex || 0);
  const dataEndRow = tableEndRow > 0
    ? Math.min(Number(sheet.getLastRow() || 0), tableEndRow)
    : Number(sheet.getLastRow() || 0);
  if (dataEndRow < dataStartRow) return [];
  const startColumn = Number(range.startColumnIndex || 0) + firstIndex + 1;
  const values = sheet
    .getRange(dataStartRow, startColumn, dataEndRow - dataStartRow + 1, lastIndex - firstIndex + 1)
    .getDisplayValues();
  return values.map((row) => indexes.map((index) => String(row[index - firstIndex] || '').trim()));
}

function importOrderAllocationRecords_(table, records) {
  return isNativeReadableTable_(table) ? importNativeRecords_(table, records) : importSystemRecords_(table, records);
}
