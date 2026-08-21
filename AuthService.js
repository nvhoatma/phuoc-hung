/**
 * Spreadsheet access guard.
 *
 * All app rendering and includes should pass through this guard so users
 * without spreadsheet access see AccessDenied.html instead of partial UI.
 */
const ACCESS_CACHE_TTL_SECONDS = 900;
const PERMISSION_CACHE_TTL_SECONDS = 300;
let ACCESS_EXECUTION_CACHE_ = null;

function getSpreadsheetAccess_() {
  if (ACCESS_EXECUTION_CACHE_) return ACCESS_EXECUTION_CACHE_;

  const userKey = getActiveUserKey_();
  const userEmail = getActiveUserEmail_();
  const cacheKey = cacheKey_('access', userEmail || userKey || 'anonymous');
  const cached = getCached_(cacheKey);
  if (cached && cached.ok) {
    ACCESS_EXECUTION_CACHE_ = cached;
    return cached;
  }

  try {
    // Keep webapp.executeAs = USER_ACCESSING so this reflects the viewer's Sheet access.
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const access = {
      ok: true,
      spreadsheetId: spreadsheet.getId(),
      spreadsheetUrl: spreadsheet.getUrl(),
      userKey,
      userEmail,
    };
    setCached_(cacheKey, access, ACCESS_CACHE_TTL_SECONDS);
    ACCESS_EXECUTION_CACHE_ = access;
    return access;
  } catch (error) {
    if (isQuotaError_(error) && cached && cached.ok) {
      ACCESS_EXECUTION_CACHE_ = Object.assign({}, cached, {
        transient: true,
        quotaLimited: true,
        detail: error && error.message ? error.message : String(error),
      });
      return ACCESS_EXECUTION_CACHE_;
    }
    const quotaLimited = isQuotaError_(error);
    ACCESS_EXECUTION_CACHE_ = {
      ok: false,
      transient: quotaLimited,
      quotaLimited,
      spreadsheetId: SPREADSHEET_ID,
      spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit`,
      userKey,
      userEmail,
      message: quotaLimited
        ? quotaErrorMessage_()
        : 'You need access to the Google Sheet data file before opening this app.',
      detail: error && error.message ? error.message : String(error),
    };
    return ACCESS_EXECUTION_CACHE_;
  }
}

function assertSpreadsheetAccess_() {
  const access = getSpreadsheetAccess_();
  if (!access.ok) {
    throw new Error(access.message);
  }
  return access;
}

function getActiveUserKey_() {
  try {
    return Session.getTemporaryActiveUserKey();
  } catch (error) {
    return '';
  }
}

function getActiveUserEmail_() {
  try {
    return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  } catch (error) {
    return '';
  }
}

function getAppSettingPermission_(forceRefresh) {
  const access = assertSpreadsheetAccess_();
  const userEmail = access.userEmail || getActiveUserEmail_();
  const cacheKey = cacheKey_('permission', userEmail || access.userKey || 'anonymous');
  const cached = getCached_(cacheKey);
  if (!forceRefresh && cached) return cached;

  try {
    const table = getNativeTableByName_('App setting permission', !!forceRefresh) ||
      getNativeTableByName_('Setting permission', !!forceRefresh) ||
      getNativeTableByName_('App permission', !!forceRefresh);
    const tableFound = !!table;
    const permissionRows = table ? buildTableRowsData_(toPermissionRuntimeTable_(table)).rows.map((row) => row.record || {}) : [];

    const activeEmails = permissionRows
      .filter((record) => {
        const status = normalizeConfigHeader_(recordValueForPermission_(record, ['status', 'active', 'enabled']));
        return ['active', 'enabled', 'true', 'yes'].includes(status);
      })
      .map((record) => normalizePermissionEmail_(recordValueForPermission_(record, ['email', 'user_email', 'gmail', 'account'])))
      .filter(Boolean);

    const configured = activeEmails.length > 0;
    const canAccessConfig = !configured || (!!userEmail && activeEmails.includes(normalizePermissionEmail_(userEmail)));
    const permission = {
      canAccessConfig,
      userEmail,
      permissionConfigured: configured,
      tableFound,
      message: canAccessConfig
        ? ''
        : userEmail
        ? 'Your email is not allowed to open Config app.'
        : 'Config app requires a signed-in email listed in App setting permission.',
    };
    setCached_(cacheKey, permission, PERMISSION_CACHE_TTL_SECONDS);
    return permission;
  } catch (error) {
    if (isQuotaError_(error) && cached) {
      return Object.assign({}, cached, { stale: true, quotaLimited: true });
    }
    throw toFriendlyQuotaError_(error);
  }
}

function toPermissionRuntimeTable_(nativeTable) {
  return {
    key: 'config_appPermission',
    name: nativeTable.name,
    type: 'config',
    spreadsheetId: SPREADSHEET_ID,
    binding: 'nativeTable',
    apiTableId: nativeTable.tableId,
    apiTableName: nativeTable.name,
    apiRange: nativeTable.range,
    columnProperties: getNativeTableColumnProperties_(nativeTable),
    columns: getNativeTableColumns_(nativeTable),
  };
}

function assertConfigAppAccess_() {
  const permission = getAppSettingPermission_(false);
  if (!permission.canAccessConfig) {
    throw new Error(permission.message || 'You do not have permission to open Config app.');
  }
  return permission;
}

function recordValueForPermission_(record, names) {
  const normalizedNames = (names || []).map(normalizeConfigHeader_);
  const keys = Object.keys(record || {});
  for (let index = 0; index < keys.length; index += 1) {
    const key = keys[index];
    if (normalizedNames.includes(normalizeConfigHeader_(key)) && record[key] !== '') return record[key];
  }
  return '';
}

function normalizePermissionEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function renderAccessDenied_(access) {
  const template = HtmlService.createTemplateFromFile('AccessDenied');
  template.access = access;
  return template
    .evaluate()
    .setTitle('Access required')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
