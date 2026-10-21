'use strict';
const DATE_LIKE = /^(?:\d{1,2}[-/.]\d{1,2}[-/.]\d{4}|\d{4}[-/.]\d{1,2}[-/.]\d{1,2})/;

/** Masks emails and phone-like numbers (9 to 15 digits). Dates such as 14-10-2026 10:30 are kept. */
function redact(s) {
  return String(s ?? '')
    .replace(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, '[email removed]')
    .replace(/(?<!\d)\+?\d(?:[\s().-]?\d){8,}(?!\d)/g, (m) => {
      const digits = m.replace(/\D/g, '').length;
      if (digits < 9 || digits > 15 || DATE_LIKE.test(m)) return m;
      return '[number removed]';
    });
}
module.exports = { redact };
