function normalizeKey(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function digitsOnly(value) {
  return String(value ?? '').replace(/\D+/g, '');
}

function normalizeWhitespace(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function normalizeTextPreservingCase(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function normalizeUppercaseText(value) {
  return normalizeTextPreservingCase(value).toUpperCase();
}

function formatCnpj(value) {
  const digits = digitsOnly(value);
  if (digits.length !== 14) {
    return normalizeTextPreservingCase(value);
  }

  return digits.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5');
}

function parseBrazilianMoney(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }

  const text = normalizeTextPreservingCase(value);
  if (!text) {
    return null;
  }

  const sanitized = text
    .replace(/[R$\s]/g, '')
    .replace(/\.(?=\d{3}(?:\D|$))/g, '')
    .replace(',', '.');

  const numeric = Number(sanitized);
  return Number.isFinite(numeric) ? numeric : null;
}

function parseBrazilianDate(value) {
  if (!value) {
    return null;
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(
      value.getUTCFullYear(),
      value.getUTCMonth(),
      value.getUTCDate(),
      value.getUTCHours(),
      value.getUTCMinutes(),
      value.getUTCSeconds(),
      value.getUTCMilliseconds(),
    );
  }

  const text = normalizeTextPreservingCase(value);
  const match = text.match(
    /(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/,
  );

  if (!match) {
    return null;
  }

  const day = Number(match[1]);
  const month = Number(match[2]) - 1;
  const year = Number(match[3].length === 2 ? `20${match[3]}` : match[3]);
  const hour = Number(match[4] ?? 0);
  const minute = Number(match[5] ?? 0);
  const second = Number(match[6] ?? 0);

  const date = new Date(year, month, day, hour, minute, second);
  return Number.isNaN(date.getTime()) ? null : date;
}

function formatDate(date, withTime = false) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) {
    return '';
  }

  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const year = String(date.getFullYear());
  const base = `${month}/${day}/${year}`;

  if (!withTime) {
    return base;
  }

  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  return `${base} ${hours}:${minutes}`;
}

function normalizeBooleanLike(value) {
  const key = normalizeKey(value);
  if (!key) {
    return '';
  }

  if (['sim', 's', 'yes', 'true', '1'].includes(key)) {
    return 'Sim';
  }

  if (['nao', 'n', 'no', 'false', '0'].includes(key)) {
    return 'Não';
  }

  return normalizeTextPreservingCase(value);
}

function normalizeDocumentNumber(value) {
  const cleaned = normalizeTextPreservingCase(value);
  if (!cleaned) {
    return '';
  }

  const digits = digitsOnly(cleaned);
  if (digits && digits.length === cleaned.replace(/\s+/g, '').length) {
    return digits;
  }

  return cleaned;
}

function safeNumberCompare(left, right) {
  if (left == null || right == null) {
    return false;
  }

  return Math.abs(Number(left) - Number(right)) < 0.01;
}

module.exports = {
  digitsOnly,
  formatCnpj,
  formatDate,
  normalizeBooleanLike,
  normalizeDocumentNumber,
  normalizeKey,
  normalizeTextPreservingCase,
  normalizeUppercaseText,
  normalizeWhitespace,
  parseBrazilianDate,
  parseBrazilianMoney,
  safeNumberCompare,
};
