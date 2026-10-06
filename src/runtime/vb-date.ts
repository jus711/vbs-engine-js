// VBScript dates are OLE Automation dates: a day count from 30 December 1899, with the time of day as
// the fraction. A Date here carries the same instant as a local wall-clock time.

const MS_PER_DAY = 86400000;
const OLE_EPOCH_UTC = Date.UTC(1899, 11, 30);

/** The OLE Automation number of a date: whole days since 30 December 1899, time of day as the fraction. */
export function dateToSerial(date: Date): number {
  const wallClock = Date.UTC(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    date.getHours(),
    date.getMinutes(),
    date.getSeconds(),
    date.getMilliseconds()
  );
  return (wallClock - OLE_EPOCH_UTC) / MS_PER_DAY;
}

/** The date an OLE Automation number stands for, as local wall-clock time. */
export function serialToDate(serial: number): Date {
  const wallClock = new Date(OLE_EPOCH_UTC + Math.round(serial * MS_PER_DAY));
  return new Date(
    wallClock.getUTCFullYear(),
    wallClock.getUTCMonth(),
    wallClock.getUTCDate(),
    wallClock.getUTCHours(),
    wallClock.getUTCMinutes(),
    wallClock.getUTCSeconds(),
    wallClock.getUTCMilliseconds()
  );
}

function isDayZero(date: Date): boolean {
  return date.getFullYear() === 1899 && date.getMonth() === 11 && date.getDate() === 30;
}

function isMidnight(date: Date): boolean {
  return date.getHours() === 0 && date.getMinutes() === 0 && date.getSeconds() === 0;
}

// Windows' default short date keeps day and month at two digits in nearly every locale; US English is the
// common exception (M/d/yyyy).
function unpaddedDayMonth(locale: string): boolean {
  return locale.toLowerCase() === 'en-us' || locale.toLowerCase() === 'en';
}

/**
 * CStr of a date, as VBScript writes it: the short date and the long time of the locale, without the date
 * when it is day 0 (30 December 1899) and without the time at midnight.
 */
export function formatVbDate(date: Date, locale: string): string {
  const padding = unpaddedDayMonth(locale) ? 'numeric' : '2-digit';
  const datePart = new Intl.DateTimeFormat(locale, { year: 'numeric', month: padding, day: padding }).format(date);
  const hour12 = new Intl.DateTimeFormat(locale, { hour: 'numeric' }).resolvedOptions().hour12 === true;
  const timePart = new Intl.DateTimeFormat(locale, {
    hour: hour12 ? 'numeric' : '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12,
  }).format(date);

  if (isDayZero(date)) return timePart;
  if (isMidnight(date)) return datePart;
  return `${datePart} ${timePart}`;
}

type DateOrder = 'DMY' | 'MDY' | 'YMD';

function dateOrderOf(locale: string): { order: DateOrder; separator: string } {
  const parts = new Intl.DateTimeFormat(locale, { year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(
    new Date(2026, 8, 29)
  );
  const order = parts
    .filter(part => part.type === 'day' || part.type === 'month' || part.type === 'year')
    .map(part => part.type[0].toUpperCase())
    .join('') as DateOrder;
  const separator = parts.find(part => part.type === 'literal')?.value.trim() || '/';
  return { order: order === 'MDY' || order === 'YMD' ? order : 'DMY', separator };
}

function toFullYear(year: number, digits: number): number {
  if (digits > 2) return year;
  return year < 30 ? 2000 + year : 1900 + year;
}

function validDate(year: number, month: number, day: number): Date | null {
  const date = new Date(year, month - 1, day);
  date.setFullYear(year);
  return date.getMonth() === month - 1 && date.getDate() === day ? date : null;
}

function parseDatePart(text: string, locale: string): Date | null {
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text);
  if (iso) return validDate(Number(iso[1]), Number(iso[2]), Number(iso[3]));

  const { order, separator } = dateOrderOf(locale);
  const separators = new Set([separator, '/', '-']);
  const match = /^(\d{1,4})([./-])(\d{1,2})\2(\d{1,4})$/.exec(text);
  if (!match || !separators.has(match[2])) return null;

  const [first, second, third] = [match[1], match[3], match[4]];
  if (order === 'YMD') return validDate(toFullYear(Number(first), first.length), Number(second), Number(third));

  const year = toFullYear(Number(third), third.length);
  const [day, month] = order === 'DMY' ? [Number(first), Number(second)] : [Number(second), Number(first)];
  // VBScript falls back to the other order when the locale's does not make a date (9/29/2026 in German).
  return validDate(year, month, day) ?? validDate(year, day, month);
}

function parseTimePart(text: string): [number, number, number] | null {
  const match = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s*([AaPp][Mm])?$/.exec(text);
  if (!match) return null;
  let hours = Number(match[1]);
  const minutes = Number(match[2]);
  const seconds = Number(match[3] ?? 0);
  const meridiem = match[4]?.toUpperCase();
  if (meridiem) {
    if (hours < 1 || hours > 12) return null;
    hours = (hours % 12) + (meridiem === 'PM' ? 12 : 0);
  }
  return hours < 24 && minutes < 60 && seconds < 60 ? [hours, minutes, seconds] : null;
}

/**
 * A date or a time as CDate and IsDate read it: the locale's date order and separator (or ISO), a 24-hour
 * or AM/PM time, or both. A time alone is on day 0; a date alone is at midnight. `null` when it is neither.
 */
export function parseVbDate(text: string, locale: string): Date | null {
  const trimmed = text.trim();
  const time = parseTimePart(trimmed);
  if (time) return new Date(1899, 11, 30, ...time);

  const firstSpace = trimmed.search(/[\sT]/);
  const datePart = firstSpace < 0 ? trimmed : trimmed.slice(0, firstSpace);
  const timePart = firstSpace < 0 ? '' : trimmed.slice(firstSpace + 1).trim();

  const date = parseDatePart(datePart, locale);
  if (!date) return null;
  if (timePart === '') return date;

  const timeOfDay = parseTimePart(timePart);
  if (!timeOfDay) return null;
  date.setHours(...timeOfDay);
  return date;
}
