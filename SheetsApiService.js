/**
 * Advanced Sheets API wrappers.
 *
 * Keep direct Sheets.Spreadsheets calls here so DataService/Config can stay
 * focused on app semantics instead of API call shape.
 */
function sheetsApiGet_(spreadsheetId, fields) {
  const options = fields ? { fields } : {};
  try {
    return runSheetsApiRead_(() => Sheets.Spreadsheets.get(spreadsheetId, options));
  } catch (error) {
    logAppDiagnostic_('error', 'sheets_api_read_failed', {
      operation: 'Sheets.Spreadsheets.get',
      spreadsheetId: spreadsheetId || '',
      fields: fields || '',
    }, error);
    throw error;
  }
}

function sheetsApiGetByDataFilter_(spreadsheetId, dataFilters, fields) {
  const options = fields ? { fields } : {};
  try {
    return runSheetsApiRead_(() => Sheets.Spreadsheets.getByDataFilter({
      dataFilters,
      includeGridData: true,
    }, spreadsheetId, options));
  } catch (error) {
    logAppDiagnostic_('error', 'sheets_api_read_failed', {
      operation: 'Sheets.Spreadsheets.getByDataFilter',
      spreadsheetId: spreadsheetId || '',
      dataFilterCount: (dataFilters || []).length,
      fields: fields || '',
    }, error);
    throw error;
  }
}

function sheetsApiBatchUpdate_(spreadsheetId, requests) {
  return Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheetId);
}

function runSheetsApiRead_(callback) {
  const delays = [0, 250, 800];
  let lastError = null;
  for (let index = 0; index < delays.length; index += 1) {
    if (delays[index]) Utilities.sleep(delays[index]);
    try {
      return callback();
    } catch (error) {
      lastError = error;
      if (!isQuotaError_(error)) throw error;
    }
  }
  throw toFriendlyQuotaError_(lastError);
}
