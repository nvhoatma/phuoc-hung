function normalizeName(name) {
  return String(name || '').trim().toLowerCase();
}

function appDiagnosticError_(error) {
  return {
    name: error && error.name ? String(error.name) : 'Error',
    message: error && error.message ? String(error.message) : String(error || ''),
    stack: error && error.stack ? String(error.stack) : '',
  };
}

let APP_DIAGNOSTIC_TRACE_ID_ = '';

function getAppDiagnosticTraceId_() {
  if (!APP_DIAGNOSTIC_TRACE_ID_) {
    APP_DIAGNOSTIC_TRACE_ID_ = Utilities.getUuid();
  }
  return APP_DIAGNOSTIC_TRACE_ID_;
}

function logAppDiagnostic_(level, event, details, error) {
  let userEmail = '';
  try {
    userEmail = typeof getActiveUserEmail_ === 'function' ? getActiveUserEmail_() : '';
  } catch (ignored) {
    userEmail = '';
  }

  const payload = Object.assign({
    diagnostic: true,
    event: String(event || 'app_diagnostic'),
    level: String(level || 'info'),
    timestamp: new Date().toISOString(),
    traceId: getAppDiagnosticTraceId_(),
    userEmail,
  }, details || {});
  if (error) payload.error = appDiagnosticError_(error);

  const method = console && typeof console[level] === 'function' ? console[level] : console.log;
  method.call(console, JSON.stringify(payload));
  return payload;
}

function sanitizeClientPayload_(value, seen) {
  if (value == null) return value;
  if (value instanceof Date) return value.toISOString();
  if (typeof value !== 'object') return value;

  const visited = seen || [];
  if (visited.indexOf(value) >= 0) return null;
  const nextSeen = visited.concat([value]);

  if (Array.isArray(value)) {
    return value.map((item) => sanitizeClientPayload_(item, nextSeen));
  }

  return Object.keys(value).reduce((acc, key) => {
    let safeKey = String(key || 'blank');
    if (safeKey === '__proto__' || safeKey === 'constructor' || /__$/.test(safeKey)) {
      safeKey = `_${safeKey.replace(/_+$/g, '') || 'blank'}`;
    }
    if (Object.prototype.hasOwnProperty.call(acc, safeKey)) safeKey = `${safeKey}_field`;
    acc[safeKey] = sanitizeClientPayload_(value[key], nextSeen);
    return acc;
  }, {});
}

function isQuotaError_(error) {
  const message = String(error && error.message ? error.message : error || '').toLowerCase();
  return message.indexOf('quota exceeded') >= 0 ||
    message.indexOf('rate_limit') >= 0 ||
    message.indexOf('resource_exhausted') >= 0 ||
    message.indexOf('too many requests') >= 0 ||
    message.indexOf('429') >= 0;
}

function quotaErrorMessage_() {
  return 'Google Sheets read quota is temporarily busy. Please wait a few seconds and reload.';
}

function toFriendlyQuotaError_(error) {
  if (!isQuotaError_(error)) return error;
  const detail = error && error.message ? error.message : String(error || '');
  return new Error(`${quotaErrorMessage_()} ${detail}`);
}
