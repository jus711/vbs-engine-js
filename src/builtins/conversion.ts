import type { VbValue } from '../runtime/index.ts';
import { toNumber, toString, toBoolean, isEmpty, isNull, isNumeric } from '../runtime/index.ts';
import { getCurrentCurrency, getCurrentBCP47Locale } from './locale.ts';

export const conversionFunctions = {
  CBool: (expression: VbValue): VbValue => {
    return { type: 'Boolean', value: toBoolean(expression) };
  },

  CByte: (expression: VbValue): VbValue => {
    const num = Math.floor(toNumber(expression));
    if (num < 0 || num > 255) {
      throw new Error('Overflow: CByte');
    }
    return { type: 'Byte', value: num };
  },

  CCur: (expression: VbValue): VbValue => {
    return { type: 'Currency', value: toNumber(expression) };
  },

  CDate: (expression: VbValue): VbValue => {
    if (expression.type === 'Date') {
      return expression;
    }
    const str = toString(expression);
    const d = new Date(str);
    if (isNaN(d.getTime())) {
      throw new Error('Type mismatch: CDate');
    }
    return { type: 'Date', value: d };
  },

  CDbl: (expression: VbValue): VbValue => {
    return { type: 'Double', value: toNumber(expression) };
  },

  CInt: (expression: VbValue): VbValue => {
    const num = Math.round(toNumber(expression));
    if (num < -32768 || num > 32767) {
      throw new Error('Overflow: CInt');
    }
    return { type: 'Integer', value: num };
  },

  CLng: (expression: VbValue): VbValue => {
    const num = Math.round(toNumber(expression));
    return { type: 'Long', value: num };
  },

  CLngLng: (expression: VbValue): VbValue => {
    if (expression.type === 'LongLong') return expression;
    if (expression.type === 'String') {
      const str = (expression.value as string).trim();
      try {
        return { type: 'LongLong', value: BigInt(str) };
      } catch {
        const num = parseFloat(str);
        if (isNaN(num)) throw new Error('Type mismatch: CLngLng');
        return { type: 'LongLong', value: BigInt(Math.trunc(num)) };
      }
    }
    const num = toNumber(expression);
    return { type: 'LongLong', value: BigInt(Math.trunc(num)) };
  },

  CSng: (expression: VbValue): VbValue => {
    return { type: 'Single', value: toNumber(expression) };
  },

  CStr: (expression: VbValue): VbValue => {
    return { type: 'String', value: toString(expression) };
  },

  CVar: (expression: VbValue): VbValue => {
    return expression;
  },

  CVErr: (errorNumber: VbValue): VbValue => {
    return { type: 'Error', value: toNumber(errorNumber) };
  },

  Val: (string: VbValue): VbValue => {
    const str = toString(string).trim();
    const match = str.match(/^[-+]?[0-9]*\.?[0-9]+([eE][-+]?[0-9]+)?/);
    if (match) {
      return { type: 'Double', value: parseFloat(match[0]) };
    }
    return { type: 'Double', value: 0 };
  },

  Str: (number: VbValue): VbValue => {
    const num = toNumber(number);
    const str = String(num);
    if (num >= 0) {
      return { type: 'String', value: ' ' + str };
    }
    return { type: 'String', value: str };
  },

  FormatNumber: (
    expression: VbValue,
    numDigitsAfterDecimal?: VbValue,
    includeLeadingDigit?: VbValue,
    _useParensForNegativeNumbers?: VbValue,
    groupDigits?: VbValue
  ): VbValue => {
    const num = toNumber(expression);
    const decimals = numDigitsAfterDecimal ? Math.floor(toNumber(numDigitsAfterDecimal)) : -1;
    const leadingDigit = includeLeadingDigit ? toNumber(includeLeadingDigit) : -2;
    const grouping = groupDigits ? toNumber(groupDigits) : -2;

    const options: Intl.NumberFormatOptions = {
      minimumFractionDigits: decimals >= 0 ? decimals : undefined,
      maximumFractionDigits: decimals >= 0 ? decimals : undefined,
      useGrouping: grouping === 0 ? false : grouping === -1 ? true : undefined,
    };

    if (leadingDigit === -1) {
      options.minimumIntegerDigits = 1;
    } else if (leadingDigit === 0) {
      options.minimumIntegerDigits = 2;
    }

    const locale = getCurrentBCP47Locale();
    return { type: 'String', value: num.toLocaleString(locale, options) };
  },

  FormatCurrency: (
    expression: VbValue,
    numDigitsAfterDecimal?: VbValue,
    includeLeadingDigit?: VbValue,
    _useParensForNegativeNumbers?: VbValue,
    groupDigits?: VbValue
  ): VbValue => {
    const num = toNumber(expression);
    const decimals = numDigitsAfterDecimal ? Math.floor(toNumber(numDigitsAfterDecimal)) : -1;
    const leadingDigit = includeLeadingDigit ? toNumber(includeLeadingDigit) : -2;
    const grouping = groupDigits ? toNumber(groupDigits) : -2;

    const options: Intl.NumberFormatOptions = {
      style: 'currency',
      currency: getCurrentCurrency(),
      minimumFractionDigits: decimals >= 0 ? decimals : 2,
      maximumFractionDigits: decimals >= 0 ? decimals : 2,
      useGrouping: grouping === 0 ? false : grouping === -1 ? true : undefined,
    };

    if (leadingDigit === -1) {
      options.minimumIntegerDigits = 1;
    } else if (leadingDigit === 0) {
      options.minimumIntegerDigits = 2;
    }

    const locale = getCurrentBCP47Locale();
    return { type: 'String', value: num.toLocaleString(locale, options) };
  },

  FormatPercent: (
    expression: VbValue,
    numDigitsAfterDecimal?: VbValue,
    includeLeadingDigit?: VbValue,
    _useParensForNegativeNumbers?: VbValue,
    groupDigits?: VbValue
  ): VbValue => {
    const num = toNumber(expression);
    const decimals = numDigitsAfterDecimal ? Math.floor(toNumber(numDigitsAfterDecimal)) : -1;
    const leadingDigit = includeLeadingDigit ? toNumber(includeLeadingDigit) : -2;
    const grouping = groupDigits ? toNumber(groupDigits) : -2;

    const options: Intl.NumberFormatOptions = {
      style: 'percent',
      minimumFractionDigits: decimals >= 0 ? decimals : undefined,
      maximumFractionDigits: decimals >= 0 ? decimals : undefined,
      useGrouping: grouping === 0 ? false : grouping === -1 ? true : undefined,
    };

    if (leadingDigit === -1) {
      options.minimumIntegerDigits = 1;
    } else if (leadingDigit === 0) {
      options.minimumIntegerDigits = 2;
    }

    const locale = getCurrentBCP47Locale();
    return { type: 'String', value: num.toLocaleString(locale, options) };
  },

  FormatDateTime: (date: VbValue, namedFormat?: VbValue): VbValue => {
    const d = date.type === 'Date' ? (date.value as Date) : new Date(toString(date));
    const format = namedFormat ? Math.floor(toNumber(namedFormat)) : 0;
    const locale = getCurrentBCP47Locale();

    let result: string;
    switch (format) {
      case 0:
        result = d.toLocaleString(locale);
        break;
      case 1:
        result = d.toLocaleDateString(locale, {
          weekday: 'long',
          year: 'numeric',
          month: 'long',
          day: 'numeric',
        });
        break;
      case 2:
        result = d.toLocaleDateString(locale);
        break;
      case 3:
        result = d.toLocaleTimeString(locale);
        break;
      case 4:
        result = d.toLocaleTimeString(locale, { hour: '2-digit', minute: '2-digit' });
        break;
      default:
        result = d.toLocaleString();
    }

    return { type: 'String', value: result };
  },

  Hex: (number: VbValue): VbValue => {
    const num = Math.floor(toNumber(number));
    return { type: 'String', value: num.toString(16).toUpperCase() };
  },

  Oct: (number: VbValue): VbValue => {
    const num = Math.floor(toNumber(number));
    return { type: 'String', value: num.toString(8) };
  },
};

export const inspectionFunctions = {
  IsArray: (expression: VbValue): VbValue => {
    return { type: 'Boolean', value: expression.type === 'Array' };
  },

  IsDate: (expression: VbValue): VbValue => {
    if (expression.type === 'Date') {
      return { type: 'Boolean', value: true };
    }
    if (expression.type === 'String') {
      const d = new Date(expression.value as string);
      return { type: 'Boolean', value: !isNaN(d.getTime()) };
    }
    return { type: 'Boolean', value: false };
  },

  IsEmpty: (expression: VbValue): VbValue => {
    return { type: 'Boolean', value: isEmpty(expression) };
  },

  IsNull: (expression: VbValue): VbValue => {
    return { type: 'Boolean', value: isNull(expression) };
  },

  IsNumeric: (expression: VbValue): VbValue => {
    return { type: 'Boolean', value: isNumeric(expression) };
  },

  IsObject: (expression: VbValue): VbValue => {
    return { type: 'Boolean', value: expression.type === 'Object' };
  },

  VarType: (expression: VbValue): VbValue => {
    const typeMap: Record<string, number> = {
      Empty: 0,
      Null: 1,
      Integer: 2,
      Long: 3,
      Single: 4,
      Double: 5,
      Currency: 6,
      Date: 7,
      String: 8,
      Object: 9,
      Error: 10,
      Boolean: 11,
      Variant: 12,
      Byte: 17,
      LongLong: 20,
      Array: 8192,
    };
    return { type: 'Integer', value: typeMap[expression.type] ?? 12 };
  },

  TypeName: (expression: VbValue): VbValue => {
    if (expression.type === 'Object' && expression.value === null) return { type: 'String', value: 'Nothing' };
    const typeNames: Record<string, string> = {
      Empty: 'Empty',
      Null: 'Null',
      Integer: 'Integer',
      Long: 'Long',
      LongLong: 'LongLong',
      Single: 'Single',
      Double: 'Double',
      Currency: 'Currency',
      Date: 'Date',
      String: 'String',
      Object: 'Object',
      Error: 'Error',
      Boolean: 'Boolean',
      Variant: 'Variant',
      Byte: 'Byte',
      Array: 'Variant()',
    };
    return { type: 'String', value: typeNames[expression.type] ?? 'Variant' };
  },
};
