import { createVbError, VbErrorCodes } from './errors.ts';
import { dateToSerial, formatVbDate, parseVbDate, serialToDate } from './vb-date.ts';
import { getCurrentBCP47Locale } from '../builtins/locale.ts';

/**
 * Represents all possible VBScript value types.
 * These correspond to the Variant subtypes in VBScript.
 */
export type VbValueType =
  | 'Empty'
  | 'Null'
  | 'Boolean'
  | 'Integer'
  | 'Long'
  | 'LongLong'
  | 'Single'
  | 'Double'
  | 'Currency'
  | 'Date'
  | 'String'
  | 'Object'
  | 'Error'
  | 'Variant'
  | 'Byte'
  | 'Array';

/** Represents an uninitialized variable in VBScript */
export interface VbEmptyValue {
  type: 'Empty';
  value: undefined;
}

/** Represents a Null value in VBScript (no valid data) */
export interface VbNullValue {
  type: 'Null';
  value: null;
}

/** Represents a Boolean value (True or False) */
export interface VbBooleanValue {
  type: 'Boolean';
  value: boolean;
}

/** Represents a 16-bit signed integer (-32,768 to 32,767) */
export interface VbIntegerValue {
  type: 'Integer';
  value: number;
}

/** Represents a 32-bit signed integer (-2,147,483,648 to 2,147,483,647) */
export interface VbLongValue {
  type: 'Long';
  value: number;
}

/** Represents a 64-bit signed integer (-9,223,372,036,854,775,808 to 9,223,372,036,854,775,807) */
export interface VbLongLongValue {
  type: 'LongLong';
  value: bigint;
}

/** Represents a single-precision floating-point number */
export interface VbSingleValue {
  type: 'Single';
  value: number;
}

/** Represents a double-precision floating-point number */
export interface VbDoubleValue {
  type: 'Double';
  value: number;
}

/** Represents a currency value (64-bit scaled integer) */
export interface VbCurrencyValue {
  type: 'Currency';
  value: number;
}

/** Represents a date/time value */
export interface VbDateValue {
  type: 'Date';
  value: Date;
}

/** Represents a string value */
export interface VbStringValue {
  type: 'String';
  value: string;
}

/** Represents an unsigned 8-bit integer (0 to 255) */
export interface VbByteValue {
  type: 'Byte';
  value: number;
}

export interface VbArrayValue {
  type: 'Array';
  value: unknown;
}

/** Represents an object reference */
export interface VbObjectValue {
  type: 'Object';
  value: unknown;
}

/** Represents an error value */
export interface VbErrorValue {
  type: 'Error';
  value: number;
}

/** Represents a Variant that can hold any type */
export interface VbVariantValue {
  type: 'Variant';
  value: unknown;
}

/**
 * A discriminated union of all VBScript value types.
 * Use the `type` property to determine the actual value type.
 */
export type VbValue =
  | VbEmptyValue
  | VbNullValue
  | VbBooleanValue
  | VbIntegerValue
  | VbLongValue
  | VbLongLongValue
  | VbSingleValue
  | VbDoubleValue
  | VbCurrencyValue
  | VbDateValue
  | VbStringValue
  | VbByteValue
  | VbArrayValue
  | VbObjectValue
  | VbErrorValue
  | VbVariantValue;

/**
 * Represents the data stored in a VBScript object.
 * This can be a class instance, a COM object wrapper, or a JavaScript object proxy.
 */
export interface VbObjectValueData {
  type?: string;
  classInfo?: { name: string };
  properties?: Map<string, VbValue>;
  getProperty?: (name: string) => VbValue;
  setProperty?: (name: string, value: VbValue, isSet?: boolean) => void;
  hasMethod?: (name: string) => boolean;
  getMethod?: (name: string) => { func: (...args: VbValue[]) => VbValue };
  call?: (...args: VbValue[]) => VbValue;
  [key: string]: unknown;
}

/** A singleton representing an Empty value in VBScript */
export const VbEmpty: VbEmptyValue = { type: 'Empty', value: undefined };

/** A singleton representing a Null value in VBScript */
export const VbNull: VbNullValue = { type: 'Null', value: null };

/** A singleton representing a Nothing object reference in VBScript */
export const VbNothing: VbObjectValue = { type: 'Object', value: null };

/**
 * Creates a VbValue from a JavaScript value.
 * Automatically determines the appropriate VBScript type based on the input.
 *
 * @param value - The JavaScript value to convert
 * @returns A VbValue representation
 */
export function createVbValue(value: unknown): VbValue {
  if (value === undefined) return VbEmpty;
  if (value === null) return VbNull;
  if (value === Symbol.for('Nothing')) return VbNothing;
  if (typeof value === 'boolean') return { type: 'Boolean', value };
  if (typeof value === 'number') {
    if (Number.isInteger(value)) {
      if (value >= -32768 && value <= 32767) {
        return { type: 'Integer', value };
      }
      return { type: 'Long', value };
    }
    return { type: 'Double', value };
  }
  if (typeof value === 'string') return { type: 'String', value };
  if (value instanceof Date) return { type: 'Date', value };
  if (Array.isArray(value)) return { type: 'Array', value };
  if (typeof value === 'object') return { type: 'Object', value: value as VbObjectValueData };
  return { type: 'Variant', value };
}

/**
 * Converts a VbValue to a JavaScript boolean.
 * Follows VBScript type coercion rules.
 *
 * @param value - The VbValue to convert
 * @returns The boolean representation
 * @throws Error if the value cannot be converted (e.g., Null)
 */
export function toBoolean(value: VbValue): boolean {
  if (value.type === 'Boolean') return value.value;
  if (value.type === 'Empty') return false;
  // Real VBScript's error 94 ("Invalid use of Null") for exactly this case -
  // must be a real, catchable VbError (not a plain Error) so `On Error
  // Resume Next` around a construct that puts Null in a boolean context
  // (`If Null Then`, `While Null`, etc.) can actually suppress it, matching
  // real VBScript behavior. Found via Wine's own vbscript.dll conformance
  // suite, dlls/vbscript/tests/lang.vbs.
  if (value.type === 'Null') throw createVbError(VbErrorCodes.InvalidUseOfNull, 'Invalid use of Null');
  if (value.type === 'String') {
    const str = value.value;
    if (str === '') return false;
    const num = parseFloat(str);
    if (!isNaN(num)) return num !== 0;
    return str.toLowerCase() !== 'false';
  }
  if (
    value.type === 'Integer' ||
    value.type === 'Long' ||
    value.type === 'Double' ||
    value.type === 'Single' ||
    value.type === 'Byte'
  ) {
    return value.value !== 0;
  }
  if (value.type === 'LongLong') {
    return value.value !== BigInt(0);
  }
  return true;
}

const HEX_STRING_RE = /^&[hH]([0-9a-fA-F]+)&?$/;
const OCTAL_STRING_RE = /^&[oO]?([0-7]+)&?$/;
const DECIMAL_STRING_RE = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/;

/**
 * Strictly parses a trimmed string as a VBScript numeric literal (decimal,
 * hex `&hFF`/`&HFF`, or octal `&o17`/`&17`), requiring the ENTIRE string to
 * be consumed (unlike JS's `parseFloat`, which happily accepts a numeric
 * PREFIX of a longer string like `"5abc"` -> 5 - real VBScript's own
 * string-to-number coercion requires the whole string to be numeric).
 * Returns null if the string doesn't parse as a real VBScript number at
 * all, letting callers decide what to do (throw, fall back to string
 * comparison, etc.) rather than baking in one behavior here.
 */
export function parseVbNumericString(trimmed: string): number | null {
  if (trimmed === '') return null;
  const hexMatch = HEX_STRING_RE.exec(trimmed);
  if (hexMatch) return parseInt(hexMatch[1]!, 16);
  const octalMatch = OCTAL_STRING_RE.exec(trimmed);
  if (octalMatch) return parseInt(octalMatch[1]!, 8);
  if (!DECIMAL_STRING_RE.test(trimmed)) return null;
  const num = parseFloat(trimmed);
  return isNaN(num) ? null : num;
}

/**
 * Converts a VbValue to a JavaScript number.
 * Follows VBScript type coercion rules.
 *
 * @param value - The VbValue to convert
 * @returns The number representation
 * @throws Error if the value cannot be converted
 */
export function toNumber(value: VbValue): number {
  if (value.type === 'Empty') return 0;
  if (value.type === 'Null') throw createVbError(VbErrorCodes.InvalidUseOfNull, 'Invalid use of Null');
  if (value.type === 'Boolean') return value.value ? -1 : 0;
  if (
    value.type === 'Integer' ||
    value.type === 'Long' ||
    value.type === 'Double' ||
    value.type === 'Single' ||
    value.type === 'Byte' ||
    value.type === 'Currency'
  ) {
    return value.value;
  }
  if (value.type === 'LongLong') {
    return Number(value.value);
  }
  if (value.type === 'String') {
    const str = value.value.trim();
    if (str === '') return 0;
    const num = parseVbNumericString(str);
    if (num === null) throw createVbError(VbErrorCodes.TypeMismatch, `Type mismatch: "${str}" cannot be converted to Number`);
    return num;
  }
  if (value.type === 'Date') {
    return dateToSerial(value.value);
  }
  throw createVbError(VbErrorCodes.TypeMismatch, `Type mismatch: ${value.type} cannot be converted to Number`);
}

/**
 * Converts a VbValue to a JavaScript string.
 * Follows VBScript type coercion rules.
 *
 * @param value - The VbValue to convert
 * @returns The string representation
 */
export function toString(value: VbValue): string {
  if (value.type === 'Empty') return '';
  if (value.type === 'Null') return 'Null';
  if (value.type === 'Boolean') return value.value ? 'True' : 'False';
  if (value.type === 'String') return value.value;
  if (value.type === 'Integer' || value.type === 'Long' || value.type === 'Byte') {
    return String(Math.floor(value.value));
  }
  if (value.type === 'LongLong') {
    return String(value.value);
  }
  if (value.type === 'Double' || value.type === 'Single' || value.type === 'Currency') {
    return String(value.value);
  }
  if (value.type === 'Date') {
    return formatVbDate(value.value, getCurrentBCP47Locale());
  }
  if (value.type === 'Object') {
    return value.value === null ? 'Nothing' : '[object]';
  }
  if (value.type === 'Array') {
    return '[array]';
  }
  return String(value.value);
}

/**
 * Converts a VbValue to a JavaScript Date object.
 * Handles VBScript date serial numbers and string date formats.
 *
 * @param value - The VbValue to convert
 * @returns The Date representation
 * @throws Error if the value cannot be converted to a date
 */
export function toVbDate(value: VbValue): Date {
  if (value.type === 'Date') return value.value;
  if (value.type === 'String') {
    const str = value.value;
    const date = parseVbDate(str, getCurrentBCP47Locale()) ?? new Date(str);
    if (isNaN(date.getTime())) {
      throw createVbError(VbErrorCodes.TypeMismatch, `Type mismatch: "${str}" cannot be converted to Date`);
    }
    return date;
  }
  if (isNumeric(value)) {
    return serialToDate(toNumber(value));
  }
  throw createVbError(VbErrorCodes.TypeMismatch, `Type mismatch: ${value.type} cannot be converted to Date`);
}

/**
 * Checks if a VbValue can be interpreted as a numeric value.
 * Equivalent to VBScript's IsNumeric function.
 *
 * @param value - The VbValue to check
 * @returns True if the value is numeric or can be converted to a number
 */
export function isNumeric(value: VbValue): boolean {
  if (
    value.type === 'Integer' ||
    value.type === 'Long' ||
    value.type === 'LongLong' ||
    value.type === 'Double' ||
    value.type === 'Single' ||
    value.type === 'Currency' ||
    value.type === 'Byte'
  ) {
    return true;
  }
  if (value.type === 'String') {
    const str = value.value.trim();
    if (str === '') return false;
    return !isNaN(parseFloat(str));
  }
  if (value.type === 'Boolean') return true;
  if (value.type === 'Empty') return true;
  return false;
}

/** Checks if a VbValue is Empty (uninitialized) */
export function isEmpty(value: VbValue): boolean {
  return value.type === 'Empty';
}

/** Checks if a VbValue is Null */
export function isNull(value: VbValue): boolean {
  return value.type === 'Null';
}

/** Checks if a VbValue is an Object */
export function isObject(value: VbValue): boolean {
  return value.type === 'Object';
}

/** Checks if a VbValue is an Array */
export function isArray(value: VbValue): boolean {
  return value.type === 'Array';
}

/**
 * Gets the numeric value from a VbValue, converting if necessary.
 * @param value - The VbValue to get the numeric value from
 * @returns The numeric value
 */
export function getNumericValue(value: VbValue): number {
  switch (value.type) {
    case 'Integer':
    case 'Long':
    case 'Double':
    case 'Single':
    case 'Byte':
    case 'Currency':
      return value.value;
    case 'LongLong':
      return Number(value.value);
    default:
      return toNumber(value);
  }
}

/**
 * Gets the string value from a VbValue, converting if necessary.
 * @param value - The VbValue to get the string value from
 * @returns The string value
 */
export function getStringValue(value: VbValue): string {
  if (value.type === 'String') return value.value;
  return toString(value);
}

/**
 * Gets the boolean value from a VbValue, converting if necessary.
 * @param value - The VbValue to get the boolean value from
 * @returns The boolean value
 */
export function getBooleanValue(value: VbValue): boolean {
  if (value.type === 'Boolean') return value.value;
  return toBoolean(value);
}
