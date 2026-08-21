function normalizeName(name) {
  return String(name || '').trim().toLowerCase();
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
