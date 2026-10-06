import type { VbValue } from '../runtime/index.ts';
import { serialToDate, toNumber, toString, toVbDate } from '../runtime/index.ts';

export const dateFunctions = {
  Now: (): VbValue => {
    return { type: 'Date', value: new Date() };
  },

  Date: (): VbValue => {
    const now = new Date();
    return { type: 'Date', value: new Date(now.getFullYear(), now.getMonth(), now.getDate()) };
  },

  Time: (): VbValue => {
    const now = new Date();
    return {
      type: 'Date',
      value: new Date(1899, 11, 30, now.getHours(), now.getMinutes(), now.getSeconds()),
    };
  },

  Year: (date: VbValue): VbValue => {
    const d = date.type === 'Date' ? (date.value as Date) : serialToDate(toNumber(date));
    return { type: 'Integer', value: d.getFullYear() };
  },

  Month: (date: VbValue): VbValue => {
    const d = date.type === 'Date' ? (date.value as Date) : serialToDate(toNumber(date));
    return { type: 'Integer', value: d.getMonth() + 1 };
  },

  Day: (date: VbValue): VbValue => {
    const d = date.type === 'Date' ? (date.value as Date) : serialToDate(toNumber(date));
    return { type: 'Integer', value: d.getDate() };
  },

  Weekday: (date: VbValue, firstDayOfWeek?: VbValue): VbValue => {
    const d = date.type === 'Date' ? (date.value as Date) : serialToDate(toNumber(date));
    const firstDay = firstDayOfWeek ? Math.floor(toNumber(firstDayOfWeek)) : 1;
    let day = d.getDay() + 1;
    day = day - firstDay + 1;
    if (day < 1) day += 7;
    return { type: 'Integer', value: day };
  },

  Hour: (time: VbValue): VbValue => {
    const d = time.type === 'Date' ? (time.value as Date) : serialToDate(toNumber(time));
    return { type: 'Integer', value: d.getHours() };
  },

  Minute: (time: VbValue): VbValue => {
    const d = time.type === 'Date' ? (time.value as Date) : serialToDate(toNumber(time));
    return { type: 'Integer', value: d.getMinutes() };
  },

  Second: (time: VbValue): VbValue => {
    const d = time.type === 'Date' ? (time.value as Date) : serialToDate(toNumber(time));
    return { type: 'Integer', value: d.getSeconds() };
  },

  DateAdd: (interval: VbValue, number: VbValue, date: VbValue): VbValue => {
    const intv = toString(interval).toLowerCase();
    const num = toNumber(number);
    const d = date.type === 'Date' ? new Date(date.value as Date) : serialToDate(toNumber(date));

    switch (intv) {
      case 'yyyy':
        d.setFullYear(d.getFullYear() + num);
        break;
      case 'q':
        d.setMonth(d.getMonth() + num * 3);
        break;
      case 'm':
        d.setMonth(d.getMonth() + num);
        break;
      case 'y':
      case 'd':
      case 'w':
        d.setDate(d.getDate() + num);
        break;
      case 'ww':
        d.setDate(d.getDate() + num * 7);
        break;
      case 'h':
        d.setHours(d.getHours() + num);
        break;
      case 'n':
        d.setMinutes(d.getMinutes() + num);
        break;
      case 's':
        d.setSeconds(d.getSeconds() + num);
        break;
    }

    return { type: 'Date', value: d };
  },

  DateDiff: (
    interval: VbValue,
    date1: VbValue,
    date2: VbValue,
    _firstDayOfWeek?: VbValue,
    _firstWeekOfYear?: VbValue
  ): VbValue => {
    void _firstDayOfWeek; // Intentionally unused - matches VBScript signature
    void _firstWeekOfYear; // Intentionally unused - matches VBScript signature
    const intv = toString(interval).toLowerCase();
    const d1 = date1.type === 'Date' ? (date1.value as Date) : serialToDate(toNumber(date1));
    const d2 = date2.type === 'Date' ? (date2.value as Date) : serialToDate(toNumber(date2));
    const diff = d2.getTime() - d1.getTime();

    let result: number;
    switch (intv) {
      case 'yyyy':
        result = d2.getFullYear() - d1.getFullYear();
        break;
      case 'q':
        result = Math.floor(
          (d2.getFullYear() - d1.getFullYear()) * 4 + (d2.getMonth() - d1.getMonth()) / 3
        );
        break;
      case 'm':
        result = (d2.getFullYear() - d1.getFullYear()) * 12 + (d2.getMonth() - d1.getMonth());
        break;
      case 'y':
      case 'd':
        result = Math.floor(diff / 86400000);
        break;
      case 'w':
        result = Math.floor(diff / 86400000);
        break;
      case 'ww':
        result = Math.floor(diff / (86400000 * 7));
        break;
      case 'h':
        result = Math.floor(diff / 3600000);
        break;
      case 'n':
        result = Math.floor(diff / 60000);
        break;
      case 's':
        result = Math.floor(diff / 1000);
        break;
      default:
        result = 0;
    }

    return { type: 'Long', value: result };
  },

  DatePart: (
    interval: VbValue,
    date: VbValue,
    _firstDayOfWeek?: VbValue,
    _firstWeekOfYear?: VbValue
  ): VbValue => {
    void _firstDayOfWeek; // Intentionally unused - matches VBScript signature
    void _firstWeekOfYear; // Intentionally unused - matches VBScript signature
    const intv = toString(interval).toLowerCase();
    const d = date.type === 'Date' ? (date.value as Date) : serialToDate(toNumber(date));

    let result: number;
    switch (intv) {
      case 'yyyy':
        result = d.getFullYear();
        break;
      case 'q':
        result = Math.floor(d.getMonth() / 3) + 1;
        break;
      case 'm':
        result = d.getMonth() + 1;
        break;
      case 'y': {
        const startOfYear = new Date(d.getFullYear(), 0, 1);
        result = Math.floor((d.getTime() - startOfYear.getTime()) / 86400000) + 1;
        break;
      }
      case 'd':
        result = d.getDate();
        break;
      case 'w':
        result = d.getDay() + 1;
        break;
      case 'ww': {
        const startOfYear2 = new Date(d.getFullYear(), 0, 1);
        result = Math.floor((d.getTime() - startOfYear2.getTime()) / (86400000 * 7)) + 1;
        break;
      }
      case 'h':
        result = d.getHours();
        break;
      case 'n':
        result = d.getMinutes();
        break;
      case 's':
        result = d.getSeconds();
        break;
      default:
        result = 0;
    }

    return { type: 'Integer', value: result };
  },

  DateSerial: (year: VbValue, month: VbValue, day: VbValue): VbValue => {
    const y = Math.floor(toNumber(year));
    const m = Math.floor(toNumber(month));
    const d = Math.floor(toNumber(day));
    return { type: 'Date', value: new Date(y, m - 1, d) };
  },

  TimeSerial: (hour: VbValue, minute: VbValue, second: VbValue): VbValue => {
    const h = Math.floor(toNumber(hour));
    const m = Math.floor(toNumber(minute));
    const s = Math.floor(toNumber(second));
    return { type: 'Date', value: new Date(1899, 11, 30, h, m, s) };
  },

  DateValue: (date: VbValue): VbValue => {
    const d = toVbDate(date);
    return { type: 'Date', value: new Date(d.getFullYear(), d.getMonth(), d.getDate()) };
  },

  TimeValue: (time: VbValue): VbValue => {
    const d = toVbDate(time);
    return { type: 'Date', value: new Date(1899, 11, 30, d.getHours(), d.getMinutes(), d.getSeconds()) };
  },

  MonthName: (month: VbValue, abbreviate?: VbValue): VbValue => {
    const m = Math.floor(toNumber(month));
    const names = [
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
    const abbrev = abbreviate ? toNumber(abbreviate) !== 0 : false;
    const name = names[m - 1] || '';
    return { type: 'String', value: abbrev ? name.substring(0, 3) : name };
  },

  WeekdayName: (weekday: VbValue, abbreviate?: VbValue, _firstDayOfWeek?: VbValue): VbValue => {
    void _firstDayOfWeek; // Intentionally unused - matches VBScript signature
    const w = Math.floor(toNumber(weekday));
    const names = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const abbrev = abbreviate ? toNumber(abbreviate) !== 0 : false;
    const name = names[w - 1] || '';
    return { type: 'String', value: abbrev ? name.substring(0, 3) : name };
  },

  Timer: (): VbValue => {
    const now = new Date();
    return {
      type: 'Single',
      value:
        now.getHours() * 3600 +
        now.getMinutes() * 60 +
        now.getSeconds() +
        now.getMilliseconds() / 1000,
    };
  },
};
