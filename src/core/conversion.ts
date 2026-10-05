import type { VbValue } from '../runtime/index.ts';
import { VbNothing } from '../runtime/index.ts';

/**
 * Converts a VbValue to its JavaScript equivalent.
 *
 * @param value - The VbValue to convert
 * @returns The JavaScript value (undefined, null, boolean, number, string, Date, array, or object)
 *
 * @example
 * ```typescript
 * const vbString = { type: 'String', value: 'hello' };
 * const jsString = vbToJs(vbString); // 'hello'
 * ```
 */
export function vbToJs(value: VbValue): unknown {
  switch (value.type) {
    case 'Empty':
      return undefined;
    case 'Null':
      return null;
    case 'Boolean':
    case 'Long':
    case 'Double':
    case 'Integer':
    case 'String':
    case 'Date':
      return value.value;
    case 'Array': {
      const arr = value.value as { toArray(): VbValue[] };
      return arr.toArray().map(vbToJs);
    }
    case 'Object': {
      const obj = value.value as Record<string, unknown> | null;
      if (
        obj &&
        typeof obj === 'object' &&
        obj['type'] === 'vbref' &&
        typeof obj['call'] === 'function'
      ) {
        const callFn = obj['call'] as (...args: unknown[]) => unknown;
        return (...jsArgs: unknown[]): unknown => {
          const vbArgs = jsArgs.map(a => jsToVb(a));
          return vbToJs(callFn(...vbArgs) as import('../runtime/index.ts').VbValue);
        };
      }
      return value.value;
    }
    default:
      return value.value;
  }
}

/**
 * Converts a JavaScript value to a VbValue.
 * Automatically determines the appropriate VBScript type based on the input.
 *
 * @param value - The JavaScript value to convert
 * @param thisArg - Optional this context for function bindings
 * @returns A VbValue representation
 *
 * @example
 * ```typescript
 * const vbNum = jsToVb(42); // { type: 'Long', value: 42 }
 * const vbStr = jsToVb('hello'); // { type: 'String', value: 'hello' }
 * ```
 */
export function jsToVb(value: unknown, thisArg?: unknown): VbValue {
  if (value === undefined) return { type: 'Empty', value: undefined };
  if (value === null) return { type: 'Null', value: null };
  if (value === Symbol.for('Nothing')) return VbNothing;
  if (typeof value === 'boolean') return { type: 'Boolean', value };
  if (typeof value === 'number') {
    if (Number.isInteger(value) && value >= -2147483648 && value <= 2147483647) {
      return { type: 'Long', value };
    }
    return { type: 'Double', value };
  }
  if (typeof value === 'string') return { type: 'String', value };
  if (value instanceof Date) return { type: 'Date', value };
  if (Array.isArray(value)) {
    return { type: 'Array', value };
  }
  if (typeof value === 'function') {
    return {
      type: 'Object',
      value: {
        type: 'jsfunction',
        func: value as (...args: unknown[]) => unknown,
        thisArg: thisArg ?? null,
      },
    };
  }
  if (typeof value === 'object') {
    return { type: 'Object', value };
  }
  return { type: 'String', value: String(value) };
}
