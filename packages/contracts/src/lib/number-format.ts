/**
 * Excel number format codes, the subset charts use: what an axis tick, a
 * data label or a numeric category reads like ("0%", "#,##0.00",
 * "$#,##0;[Red]-$#,##0", "m/d/yyyy", "General").
 *
 * Sections (positive;negative;zero;text), literals ("…", \c, _x, *x),
 * bracketed colour/condition/locale tokens (dropped; a [$€-…] currency keeps
 * its symbol), percent, thousands grouping and scaling commas, scientific
 * notation, and dates/times from the 1900 serial. Separators are the en-US
 * ones ("." decimal, "," thousands).
 */

/** Split on ';' outside quotes and escapes. */
function sections(code: string): string[] {
  const out: string[] = [];
  let cur = '';
  let quoted = false;
  for (let i = 0; i < code.length; i++) {
    const c = code[i];
    if (c === '"') quoted = !quoted;
    if (!quoted && c === '\\') {
      cur += c + (code[i + 1] ?? '');
      i++;
      continue;
    }
    if (!quoted && c === ';') {
      out.push(cur);
      cur = '';
      continue;
    }
    cur += c;
  }
  out.push(cur);
  return out;
}

type Token = { lit: string } | { code: string };

/** Literal runs vs format characters, brackets resolved. */
function tokenize(section: string): Token[] {
  const out: Token[] = [];
  const lit = (s: string) => out.push({ lit: s });
  for (let i = 0; i < section.length; i++) {
    const c = section[i];
    if (c === '"') {
      const end = section.indexOf('"', i + 1);
      lit(section.slice(i + 1, end < 0 ? undefined : end));
      i = end < 0 ? section.length : end;
    } else if (c === '\\') {
      lit(section[i + 1] ?? '');
      i++;
    } else if (c === '_') {
      lit(' '); // a space the width of the next character
      i++;
    } else if (c === '*') {
      i++; // fill-repeat: no width to fill here
    } else if (c === '[') {
      const end = section.indexOf(']', i);
      const inner = section.slice(i + 1, end < 0 ? undefined : end);
      // [$€-407] → "€"; colours, conditions and bare locales vanish.
      if (inner.startsWith('$')) lit(inner.slice(1).split('-')[0]);
      i = end < 0 ? section.length : end;
    } else {
      out.push({ code: c });
    }
  }
  return out;
}

const DATE_CHARS = /[ymdhs]/i;

/** "General": up to 11 significant characters, exponent for the extremes. */
function general(v: number): string {
  if (v === 0) return '0';
  const abs = Math.abs(v);
  if (abs >= 1e11 || abs < 1e-4) {
    const [m, e] = v.toExponential(5).split('e');
    const mant = m.replace(/\.?0+$/, '');
    const exp = Number(e);
    return `${mant}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
  }
  if (Number.isInteger(v)) return String(v);
  const digits = Math.max(0, 10 - Math.floor(Math.log10(abs)));
  return String(Number(v.toFixed(Math.min(digits, 15))));
}

/** Round half away from zero to `places` decimals. */
function round(v: number, places: number): number {
  const f = 10 ** places;
  return (Math.sign(v) * Math.round(Math.abs(v) * f + 1e-9)) / f;
}

function numberSection(v: number, tokens: Token[]): string {
  const codes = tokens.map((t) => ('code' in t ? t.code : '')).join('');
  const percent = (codes.match(/%/g) ?? []).length;
  let value = v * 100 ** percent;
  // Scientific notation.
  const eIdx = codes.search(/[eE][+-]/);
  // The digit pattern: placeholders, '.', ','.
  const isDigit = (c: string) => c === '0' || c === '#' || c === '?';
  // The mantissa's digits only: an exponent's own 0s come after the E.
  const eToken = tokens.findIndex(
    (t, i) =>
      'code' in t &&
      /[eE]/.test(t.code) &&
      /[+-]/.test((tokens[i + 1] as { code?: string })?.code ?? ''),
  );
  const digitTokens = tokens
    .map((t, i) => ({ t, i }))
    .filter(
      ({ t, i }) =>
        (eToken < 0 || i < eToken) &&
        'code' in t &&
        (isDigit(t.code) || t.code === '.' || t.code === ','),
    );
  const pattern = digitTokens
    .map(({ t }) => (t as { code: string }).code)
    .join('');
  // Trailing commas after the last digit placeholder scale by 1000 each.
  const trailing = pattern.match(/,+$/)?.[0].length ?? 0;
  value /= 1000 ** trailing;
  const core = pattern.slice(0, pattern.length - trailing);
  const [intPat, fracPat = ''] = core.split('.');
  const grouping = intPat.includes(',');
  const intDigits = intPat.replace(/,/g, '');
  const minInt = (intDigits.match(/0/g) ?? []).length;
  const fracMax = (fracPat.match(/[0#?]/g) ?? []).length;
  const fracMin = (fracPat.match(/0/g) ?? []).length;

  let body: string;
  if (eIdx >= 0) {
    const exp = value === 0 ? 0 : Math.floor(Math.log10(Math.abs(value)));
    const mant = round(value / 10 ** exp, fracMax);
    const expDigits = (codes.slice(eIdx + 2).match(/0/g) ?? []).length || 1;
    const sign =
      codes[eIdx + 1] === '+' || exp < 0 ? (exp < 0 ? '-' : '+') : '';
    body = `${formatDigits(Math.abs(mant), minInt, fracMin, fracMax, false, core.includes('.'))}E${sign}${String(Math.abs(exp)).padStart(expDigits, '0')}`;
  } else {
    body = formatDigits(
      Math.abs(round(value, fracMax)),
      minInt,
      fracMin,
      fracMax,
      grouping,
      core.includes('.'),
    );
  }

  // Re-assemble: literals and non-digit codes keep their places around the
  // digit run, which is emitted once where it starts.
  const firstDigit = digitTokens[0]?.i ?? -1;
  const lastDigit = digitTokens[digitTokens.length - 1]?.i ?? -1;
  let out = '';
  tokens.forEach((t, i) => {
    if (i === firstDigit) out += body;
    if (i >= firstDigit && i <= lastDigit && firstDigit >= 0) {
      if ('lit' in t) out += t.lit;
      return;
    }
    if (eToken >= 0 && i >= eToken && 'code' in t) return;
    if ('lit' in t) out += t.lit;
    else if (t.code === '%') out += '%';
    else if (!isDigit(t.code) && t.code !== '.' && t.code !== ',')
      out += t.code;
  });
  if (firstDigit < 0) out = body + out;
  return out;
}

function formatDigits(
  abs: number,
  minInt: number,
  fracMin: number,
  fracMax: number,
  grouping: boolean,
  point: boolean,
): string {
  const fixed = abs.toFixed(fracMax);
  let [int, frac = ''] = fixed.split('.');
  if (int === '0' && minInt === 0) int = '';
  int = int.padStart(minInt, '0');
  if (grouping) int = int.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  while (frac.length > fracMin && frac.endsWith('0')) frac = frac.slice(0, -1);
  // A pattern with a decimal point keeps it even with nothing after ("0.##"
  // shows 3 as "3.").
  return frac || point ? `${int}.${frac}` : int;
}

// ── dates ─────────────────────────────────────────────────────────────────

const MONTHS = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
];
const DAYS = [
  'Sunday',
  'Monday',
  'Tuesday',
  'Wednesday',
  'Thursday',
  'Friday',
  'Saturday',
];

/** Excel 1900 serial → UTC date parts (Excel's phantom 1900-02-29 kept). */
function serialDate(serial: number) {
  const days = Math.floor(serial);
  const ms = Math.round((serial - days) * 86400000);
  // Serial 60 is the non-existent 29 Feb 1900; dates after it are one off.
  const base =
    Date.UTC(1899, 11, 31) + (days > 59 ? days - 1 : days) * 86400000;
  const d = new Date(base + ms);
  return {
    y: d.getUTCFullYear(),
    m: d.getUTCMonth(),
    d: d.getUTCDate(),
    w: d.getUTCDay(),
    h: d.getUTCHours(),
    min: d.getUTCMinutes(),
    s: d.getUTCSeconds(),
  };
}

function dateSection(v: number, tokens: Token[]): string {
  const p = serialDate(v);
  const ampm = tokens.some((t) => 'code' in t && /[aA]/.test(t.code));
  let out = '';
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if ('lit' in t) {
      out += t.lit;
      continue;
    }
    const c = t.code.toLowerCase();
    let run = 1;
    while (i + run < tokens.length) {
      const n = tokens[i + run];
      if (!('code' in n) || n.code.toLowerCase() !== c) break;
      run++;
    }
    // "m" right after an hour (or before a second) means minutes.
    const minuteContext =
      c === 'm' &&
      (/h\W*$/i.test(
        tokens
          .slice(0, i)
          .map((x) => ('code' in x ? x.code : x.lit))
          .join(''),
      ) ||
        /^\W*s/i.test(
          tokens
            .slice(i + run)
            .map((x) => ('code' in x ? x.code : x.lit))
            .join(''),
        ));
    switch (c) {
      case 'y':
        out += run <= 2 ? String(p.y % 100).padStart(2, '0') : String(p.y);
        break;
      case 'm':
        if (minuteContext)
          out += run >= 2 ? String(p.min).padStart(2, '0') : String(p.min);
        else if (run >= 5) out += MONTHS[p.m][0];
        else if (run === 4) out += MONTHS[p.m];
        else if (run === 3) out += MONTHS[p.m].slice(0, 3);
        else
          out += run === 2 ? String(p.m + 1).padStart(2, '0') : String(p.m + 1);
        break;
      case 'd':
        if (run >= 4) out += DAYS[p.w];
        else if (run === 3) out += DAYS[p.w].slice(0, 3);
        else out += run === 2 ? String(p.d).padStart(2, '0') : String(p.d);
        break;
      case 'h': {
        const h = ampm ? ((p.h + 11) % 12) + 1 : p.h;
        out += run >= 2 ? String(h).padStart(2, '0') : String(h);
        break;
      }
      case 's':
        out += run >= 2 ? String(p.s).padStart(2, '0') : String(p.s);
        break;
      case 'a': {
        // AM/PM or A/P.
        const rest = tokens
          .slice(i, i + 5)
          .map((x) => ('code' in x ? x.code : ''))
          .join('');
        if (/^am\/pm/i.test(rest)) {
          out += p.h < 12 ? 'AM' : 'PM';
          run = 5;
        } else if (/^a\/p/i.test(rest)) {
          out += p.h < 12 ? 'A' : 'P';
          run = 3;
        } else out += t.code;
        break;
      }
      default:
        out += t.code.repeat(run);
    }
    i += run - 1;
  }
  return out;
}

/** Format `value` with an Excel number format code. */
export function formatNumber(value: number, code: string | undefined): string {
  if (!Number.isFinite(value)) return '';
  const fmt = (code ?? '').trim();
  if (!fmt || /^general$/i.test(fmt)) return general(value);
  const secs = sections(fmt);
  let section = secs[0];
  let v = value;
  let explicitSign = false;
  if (value < 0 && secs.length >= 2 && secs[1] !== '') {
    section = secs[1];
    v = Math.abs(value);
    explicitSign = true;
  } else if (value === 0 && secs.length >= 3 && secs[2] !== '') {
    section = secs[2];
  }
  if (/^general$/i.test(section.trim())) return general(v);
  const tokens = tokenize(section);
  const codes = tokens.map((t) => ('code' in t ? t.code : '')).join('');
  if (DATE_CHARS.test(codes) && !/[0#?]/.test(codes))
    return dateSection(v, tokens);
  const body = numberSection(Math.abs(v), tokens);
  const negative = v < 0 && !explicitSign && Math.abs(v) >= 1e-12;
  // A negative that rounds to zero prints without its sign.
  return negative && /[1-9]/.test(body) ? `-${body}` : body;
}
