import type {
  Expression,
  Program,
  Identifier,
  Literal,
  MemberExpression,
  BinaryExpression,
  UnaryExpression,
  LogicalExpression,
  AssignmentExpression,
  ConditionalExpression,
  ThisExpression,
  NewExpression,
  VbWithObjectExpression,
} from '../ast/index.ts';
import type {
  VbValue,
  VbObjectValue,
  VbObjectValueData,
  VbArrayValue,
  VbLongValue,
} from '../runtime/index.ts';
import { VbContext, VbObjectInstance } from '../runtime/index.ts';
import {
  createVbValue,
  VbEmpty,
  VbNull,
  VbNothing,
  toBoolean,
  toNumber,
  toString,
  createVbError,
  VbErrorCodes,
} from '../runtime/index.ts';

interface VbMethodObject {
  type: 'method';
  object: VbObjectValueData & {
    getMethod: (name: string) => { func: (...args: VbValue[]) => VbValue };
  };
  method: string;
}

interface VbJsFunctionObject {
  type: 'jsfunction';
  func: (...args: unknown[]) => unknown;
  thisArg: unknown;
}

function isVbMethodObject(obj: VbObjectValueData): obj is VbObjectValueData & VbMethodObject {
  return obj.type === 'method' && 'object' in obj && 'method' in obj;
}

function isVbJsFunctionObject(
  obj: VbObjectValueData
): obj is VbObjectValueData & VbJsFunctionObject {
  return obj.type === 'jsfunction' && 'func' in obj;
}

/**
 * "Invalid procedure call" on its own fits every call in a ninety-line procedure. Both throw
 * sites that raise it already hold the name they failed on, so it goes into the message.
 */
function invalidProcedureCall(name: string) {
  return createVbError(
    VbErrorCodes.InvalidProcedureCall,
    name ? `Invalid procedure call '${name}'` : 'Invalid procedure call'
  );
}

/** `o.M` for the common shape; the empty string when neither half is a plain name. */
function memberExpressionName(node: MemberExpression): string {
  const objectName =
    node.object.type === 'Identifier' ? `${(node.object as Identifier).name}.` : '';
  const propertyName = node.property.type === 'Identifier' ? (node.property as Identifier).name : '';
  return objectName + propertyName;
}

// Real VBScript's And/Or/Xor/Eqv/Imp/Not all ultimately call the real OLE
// Automation VarAnd/VarOr/VarXor/VarEqv/VarImp/VarNot - genuine BITWISE
// operations, not JS-style boolean logic - with two VBScript-specific
// pre-coercions those native functions don't do themselves: Empty becomes
// Long(0), and a String becomes a number if it parses as one, else a
// Boolean (True/False, via the same leniency toBoolean() already applies to
// strings). Confirmed directly from Wine's real C source,
// dlls/vbscript/interp.c's coerce_empty_to_i4()/coerce_str_to_num_or_bool(),
// and verified against many real assertions in dlls/vbscript/tests/lang.vbs
// (`("1" Or "2") = 3`, `(Empty And Empty) = 0`, etc.) - never call this on a
// Null operand, which always needs its own three-valued-logic handling
// first (see evaluateLogical below).
function toBitwiseLong(value: VbValue): number {
  if (value.type === 'Empty') return 0;
  if (value.type === 'Boolean') return value.value ? -1 : 0;
  if (value.type === 'String') {
    const parsed = Number(value.value.trim());
    if (value.value.trim() !== '' && !isNaN(parsed)) return parsed | 0;
    return toBoolean(value) ? -1 : 0;
  }
  // An Object operand (e.g. `Not MyObject`, `MyObject And True`) doesn't
  // have a numeric form `toNumber()` can produce - this engine doesn't yet
  // implement real default-property invocation (a genuinely separate,
  // bigger feature from this bitwise-math fix; see the still-open
  // `obj.publicFunction` bug for the other half of that gap). Route through
  // toBoolean() instead of throwing, matching what this engine's own
  // pre-existing (if simplistic - toBoolean() itself just defaults to
  // `true` for any Object, not real default-property resolution) behavior
  // already did for `Not`/`And`/`Or` on objects before this fix - avoids a
  // regression on that front while still fixing the actually-reported
  // bitwise-math bug for numeric/string operands.
  if (value.type === 'Object') return toBoolean(value) ? -1 : 0;
  return toNumber(value) | 0;
}

// Boolean+Boolean stays Boolean-typed (matches VarAnd's own promotion rule -
// bitwise math on -1/0 happens to equal boolean logic anyway, so this only
// affects the RESULT's reported type, e.g. TypeName()/getVT(), not its
// truthiness). Every other combination (Empty, String, numeric, mixed)
// promotes to Long, confirmed via many `getVT(...) = "VT_I4"` assertions in
// Wine's own vbscript.dll conformance suite, dlls/vbscript/tests/lang.vbs.
function bitwiseResultType(left: VbValue, right: VbValue): 'Boolean' | 'Long' {
  return left.type === 'Boolean' && right.type === 'Boolean' ? 'Boolean' : 'Long';
}

function makeBitwiseResult(value: number, left: VbValue, right: VbValue): VbValue {
  const type = bitwiseResultType(left, right);
  return type === 'Boolean' ? { type: 'Boolean', value: value !== 0 } : { type: 'Long', value };
}

export class ExpressionEvaluator {
  constructor(private context: VbContext) {}

  evaluateProgram(program: Program): VbValue {
    let result: VbValue = VbEmpty;
    for (const stmt of program.body) {
      if (stmt.type === 'ExpressionStatement') {
        result = this.evaluate((stmt as { expression: Expression }).expression);
      } else {
        result = this.evaluate(stmt as unknown as Expression);
      }
    }
    return result;
  }

  evaluate(node: Expression): VbValue {
    switch (node.type) {
      case 'Identifier':
        return this.evaluateIdentifier(node);
      case 'Literal':
        return this.evaluateLiteral(node);
      case 'VbEmptyLiteral':
        return VbEmpty;
      case 'NewExpression':
        return this.evaluateNew(node);
      case 'ThisExpression':
        return this.evaluateMe(node);
      case 'VbWithObject':
        return this.evaluateWithObject(node);
      case 'MemberExpression':
        return this.evaluateMemberAsValue(node);
      case 'CallExpression':
        return this.evaluateCallInternal(node.callee as Expression, node.arguments);
      case 'BinaryExpression':
        return this.evaluateBinary(node);
      case 'UnaryExpression':
        return this.evaluateUnary(node);
      case 'LogicalExpression':
        return this.evaluateLogical(node);
      case 'AssignmentExpression':
        return this.evaluateAssignment(node);
      case 'ConditionalExpression':
        return this.evaluateConditional(node);
      default:
        return VbEmpty;
    }
  }

  evaluateCall(callee: Expression, args: Expression[]): VbValue {
    return this.evaluateCallInternal(callee, args);
  }

  private evaluateIdentifier(node: Identifier): VbValue {
    const name = node.name;

    // A local variable/parameter (or class instance property) must shadow a
    // same-named built-in - real VBScript allows `Function trackEval(val,
    // label)` to use `val` as an ordinary parameter without it colliding
    // with the built-in Val() function. This used to check the function
    // registry FIRST unconditionally, so reading a parameter/variable that
    // happened to share a built-in's name (found via Wine's own
    // vbscript.dll conformance suite, dlls/vbscript/tests/lang.vbs: several
    // helper functions use `val` as a parameter name) silently called the
    // built-in with zero arguments instead - Val() with no argument then
    // crashed converting `undefined` to a string.
    if (this.context.hasDeclaredVariable(name)) {
      return this.context.getVariable(name);
    }

    if (this.context.functionRegistry.has(name)) {
      return this.context.functionRegistry.call(name, []);
    }

    return this.context.getVariable(name);
  }

  private evaluateLiteral(node: Literal): VbValue {
    return createVbValue(node.value);
  }

  private evaluateNew(node: NewExpression): VbObjectValue {
    // Simple identifier: try VBScript class registry first, then globalThis fallback.
    if (node.callee.type === 'Identifier') {
      const className = node.callee.name;
      const ctorArgs = node.arguments.map(arg => this.evaluate(arg));
      try {
        const instance = this.context.classRegistry.createInstance(className, ctorArgs);
        return { type: 'Object', value: instance };
      } catch {
        // Not a VBScript class — fall through to globalThis lookup below.
      }
      const ctor = (globalThis as Record<string, unknown>)[className];
      if (typeof ctor === 'function') {
        const args = ctorArgs.map(v => this.vbToJs(v));
        return {
          type: 'Object',
          value: Reflect.construct(ctor as unknown as new (...a: unknown[]) => unknown, args),
        };
      }
      throw createVbError(
        VbErrorCodes.InvalidProcedureCall,
        `Unknown class: ${className}`
      );
    }

    // Dotted path (e.g. New Forms.Form, New NS.Sub.Class): evaluate the
    // member expression to obtain the JS constructor, then call it with new.
    const ctorValue = this.evaluate(node.callee);
    let ctor: ((...args: unknown[]) => unknown) | undefined;

    if (ctorValue.type === 'Object') {
      const obj = ctorValue.value as Record<string, unknown> | null;
      if (obj && obj.type === 'jsfunction') {
        ctor = (obj as { func: (...args: unknown[]) => unknown }).func;
      } else if (typeof obj === 'function') {
        ctor = obj as (...args: unknown[]) => unknown;
      } else if (typeof obj === 'object' && obj !== null) {
        // Raw JS Proxy (e.g. node-ps1-dotnet namespace constructor)
        ctor = obj as unknown as (...args: unknown[]) => unknown;
      }
    }

    if (typeof ctor !== 'function' && !(ctor && typeof ctor === 'object')) {
      throw createVbError(VbErrorCodes.InvalidProcedureCall, 'Not a constructor');
    }

    const args = node.arguments.map(arg => this.vbToJs(this.evaluate(arg)));
    return {
      type: 'Object',
      value: Reflect.construct(ctor as unknown as new (...a: unknown[]) => unknown, args),
    };
  }

  private evaluateMe(_node: ThisExpression): VbObjectValue {
    void _node; // Intentionally unused
    if (this.context.currentInstance) {
      return { type: 'Object', value: this.context.currentInstance };
    }
    throw new Error('Me keyword not supported outside of class context');
  }

  private evaluateWithObject(_node: VbWithObjectExpression): VbValue {
    void _node; // Intentionally unused
    const withObject = this.context.getCurrentWith();
    if (!withObject) {
      throw new Error('With object not available - must be inside a With statement');
    }
    return withObject;
  }

  /**
   * A member read written WITHOUT parentheses, reached where a value is wanted. In VBScript such
   * a read of a zero-argument Sub or Function IS the call: `r = o.F` yields F's result, and a bare
   * `o.M` on its own line runs M.
   *
   * Only this dispatch is affected. `o.F(args)` reaches evaluateCallInternal(), which calls
   * evaluateMember() directly and still sees the unresolved method reference it needs, so real
   * arguments continue to reach the real call.
   */
  private evaluateMemberAsValue(node: MemberExpression): VbValue {
    const value = this.evaluateMember(node);

    if (value.type === 'Object' && value.value !== null) {
      const obj = value.value as VbObjectValueData;
      if (isVbMethodObject(obj) || isVbJsFunctionObject(obj)) {
        // Goes through the ordinary call path with no arguments, so a class method and a host
        // function exposed via addObject() are invoked exactly as they are when written with
        // parentheses -- including the jsToVb conversion a host function's result needs.
        return this.callObjectMethod(value as VbObjectValue, []);
      }
    }

    return value;
  }

  private evaluateMember(node: MemberExpression): VbValue {
    let object = this.evaluate(node.object as Expression);

    if (node.object.type === 'VbWithObject') {
      object = this.evaluateWithObject(node.object as VbWithObjectExpression);
    }

    let propertyName: string;
    if (node.computed) {
      const propValue = this.evaluate(node.property as Expression);
      propertyName = toString(propValue);
    } else {
      propertyName = (node.property as Identifier).name;
    }

    if (object.type === 'Array') {
      return this.getArrayElement(object, node.property as Expression);
    }

    if (object.type === 'Object') {
      return this.getObjectProperty(object, propertyName);
    }

    throw createVbError(VbErrorCodes.ObjectRequired, 'Object required');
  }

  private getArrayElement(array: VbArrayValue, indexExpr: Expression): VbValue {
    const arr = array.value as unknown as { get: (indices: number[]) => VbValue };
    const index = toNumber(this.evaluate(indexExpr));
    return arr.get([Math.floor(index)]);
  }

  private getObjectProperty(objValue: VbObjectValue, propertyName: string): VbValue {
    const obj = objValue.value as VbObjectValueData | null;

    if (obj === null) {
      throw createVbError(VbErrorCodes.ObjectRequired, 'Object required');
    }

    // If this is a JS function wrapper produced by a previous property access,
    // allow further member access to flow through to the underlying JS function/
    // Proxy object. This is required for chains like form.Controls.Add(...), where
    // `Controls` may itself be represented by a callable function-proxy.
    if (isVbJsFunctionObject(obj) && !['type', 'func', 'thisArg'].includes(propertyName)) {
      const jsValue = (obj.func as unknown as Record<string, unknown>)[propertyName];
      if (jsValue === undefined) {
        return { type: 'Empty', value: undefined };
      }
      if (typeof jsValue === 'function') {
        return {
          type: 'Object',
          value: {
            type: 'jsfunction',
            func: jsValue as (...args: unknown[]) => unknown,
            thisArg: obj.func,
          },
        };
      }
      return this.jsToVb(jsValue);
    }

    if (
      (Object.prototype.hasOwnProperty.call(obj, 'getProperty') ||
        Object.prototype.hasOwnProperty.call(obj, 'classInfo')) &&
      typeof obj.getProperty === 'function'
    ) {
      if (obj.hasMethod?.(propertyName)) {
        return { type: 'Object', value: { type: 'method', object: obj, method: propertyName } };
      }
      return obj.getProperty(propertyName);
    }

    const jsValue = obj[propertyName];
    if (jsValue === undefined) {
      return { type: 'Empty', value: undefined };
    }
    if (typeof jsValue === 'function') {
      return {
        type: 'Object',
        value: {
          type: 'jsfunction',
          func: jsValue as (...args: unknown[]) => unknown,
          thisArg: obj,
        },
      };
    }
    return this.jsToVb(jsValue);
  }

  private jsToVb(value: unknown): VbValue {
    if (value === undefined) return { type: 'Empty', value: undefined };
    if (value === null) return { type: 'Null', value: null };
    if (value === Symbol.for('Nothing')) return VbNothing;
    if (typeof value === 'boolean') return { type: 'Boolean', value };
    if (typeof value === 'bigint') return { type: 'LongLong', value };
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
    if (typeof value === 'object') {
      return { type: 'Object', value: value as VbObjectValueData };
    }
    return { type: 'String', value: String(value) };
  }

  private evaluateCallInternal(calleeExpr: Expression, callArgs: Expression[]): VbValue {
    if (calleeExpr.type === 'Identifier') {
      const name = calleeExpr.name;

      if (this.context.functionRegistry.has(name)) {
        const funcInfo = this.context.functionRegistry.get(name)!;
        const hasByRefParams = funcInfo.params.some(p => p.byRef);

        if (hasByRefParams) {
          const argRefs = callArgs.map(arg => {
            if (arg.type === 'Identifier') {
              const varName = arg.name;
              const variable = this.context.currentScope.get(varName);
              return {
                value: variable ? variable.value : this.context.getVariable(varName),
                variableName: varName,
                setValue: (val: VbValue) => {
                  if (this.context.currentScope.has(varName)) {
                    this.context.currentScope.set(varName, val);
                  } else {
                    this.context.setVariable(varName, val);
                  }
                },
              };
            }
            return { value: this.evaluate(arg) };
          });
          return this.context.functionRegistry.callWithRefs(name, argRefs);
        }

        const args = callArgs.map(arg => this.evaluate(arg));
        return this.context.functionRegistry.call(name, args);
      }

      const variable = this.context.currentScope.get(name);
      if (variable) {
        if (variable.value.type === 'Array') {
          const indices = callArgs.map(arg => Math.floor(toNumber(this.evaluate(arg))));
          const arr = variable.value.value as unknown as { get: (indices: number[]) => VbValue };
          return arr.get(indices);
        }
        if (variable.value.type === 'Object' && variable.value.value !== null) {
          const val = variable.value.value as VbObjectValueData;
          if (val.hasMethod?.('default') && val.getMethod) {
            const method = val.getMethod('default');
            const args = callArgs.map(arg => this.evaluate(arg));
            return method.func(...args);
          }
        }
      }

      const callee = this.context.getVariable(name);
      if (callee.type === 'Object' && callee.value !== null) {
        return this.callObjectMethod(callee, callArgs);
      }
      throw invalidProcedureCall(name);
    }

    if (calleeExpr.type === 'MemberExpression') {
      const memberResult = this.evaluateMember(calleeExpr);

      // Chained array read, e.g. `obj.Items(0)` or `arr(0)(4)` where the inner
      // `arr(0)` was itself parsed as a MemberExpression (see the Identifier
      // branch above, which already handles the non-chained `arr(4)` case the
      // same way). Missing here until 2026-09-03: this branch only ever tried
      // callObjectMethod and threw for anything else, so a nested array read
      // could never succeed - previously unreachable in practice because
      // nested-array WRITES were separately broken (see parseStatementAssignment
      // above), so `arr(0)` never actually held a populated array to expose this.
      if (memberResult.type === 'Array') {
        const indices = callArgs.map(arg => Math.floor(toNumber(this.evaluate(arg))));
        const arr = memberResult.value as unknown as { get: (indices: number[]) => VbValue };
        return arr.get(indices);
      }
      if (memberResult.type === 'Object' && memberResult.value !== null) {
        return this.callObjectMethod(memberResult, callArgs);
      }
      throw invalidProcedureCall(memberExpressionName(calleeExpr));
    }

    const callee = this.evaluate(calleeExpr);
    // Same chained-array-read gap as above, for callees more deeply nested than
    // one level (e.g. a CallExpression callee from `arr(0)(4)(1)`).
    if (callee.type === 'Array') {
      const indices = callArgs.map(arg => Math.floor(toNumber(this.evaluate(arg))));
      const arr = callee.value as unknown as { get: (indices: number[]) => VbValue };
      return arr.get(indices);
    }
    if (callee.type === 'Object' && callee.value !== null) {
      return this.callObjectMethod(callee, callArgs);
    }
    throw createVbError(VbErrorCodes.InvalidProcedureCall, 'Invalid procedure call');
  }

  private callObjectMethod(objValue: VbObjectValue, callArgs: Expression[]): VbValue {
    const obj = objValue.value as VbObjectValueData | null;
    if (obj === null) {
      throw createVbError(VbErrorCodes.ObjectRequired, 'Object required');
    }

    if (isVbMethodObject(obj)) {
      const args = callArgs.map(arg => this.evaluate(arg));
      return obj.object.getMethod(obj.method).func.call(obj.object, ...args);
    }

    if (isVbJsFunctionObject(obj)) {
      const args = callArgs.map(arg => {
        const vbVal = this.evaluate(arg);
        return this.vbToJs(vbVal);
      });
      const result = obj.func.call(obj.thisArg, ...args);
      return this.jsToVb(result);
    }

    if (typeof obj.call === 'function') {
      const args = callArgs.map(arg => this.evaluate(arg));
      return obj.call(...args);
    }

    throw createVbError(VbErrorCodes.InvalidProcedureCall, 'Invalid procedure call');
  }

  private vbToJs(value: VbValue): unknown {
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
        return value.value;
      case 'LongLong':
        return value.value; // return as bigint
      case 'Date':
        return value.value instanceof Date ? value.value : new Date(value.value);
      case 'Array':
        return value.value;
      case 'Object': {
        const obj = value.value as VbObjectValueData | null;
        if (obj && typeof obj === 'object' && (obj as Record<string, unknown>).type === 'vbref') {
          const callFn = (obj as Record<string, unknown>).call as
            ((...args: VbValue[]) => VbValue) | undefined;
          if (typeof callFn === 'function') {
            return (...jsArgs: unknown[]): unknown => {
              const vbArgs = jsArgs.map(a => this.jsToVb(a));
              return this.vbToJs(callFn(...vbArgs));
            };
          }
        }
        return value.value;
      }
      default:
        return value.value;
    }
  }

  private evaluateBinary(node: BinaryExpression): VbValue {
    const left = this.evaluate(node.left);
    const right = this.evaluate(node.right);

    switch (node.operator) {
      case '+':
        return this.add(left, right);
      case '-':
        return this.subtract(left, right);
      case '*':
        return this.multiply(left, right);
      case '/':
        return this.divide(left, right);
      case '\\':
        if (left.type === 'LongLong' || right.type === 'LongLong') {
          return this.longLongArithmetic(left, right, '\\');
        }
        return this.integerDivide(left, right);
      case '%':
      case 'Mod':
        return this.modulo(left, right);
      case '**':
      case '^':
        return this.power(left, right);
      case '==':
      case '=':
        return this.equals(left, right);
      case '!=':
      case '<>':
        return this.notEquals(left, right);
      case '<':
        return this.lessThan(left, right);
      case '<=':
        return this.lessThanOrEqual(left, right);
      case '>':
        return this.greaterThan(left, right);
      case '>=':
        return this.greaterThanOrEqual(left, right);
      case '&':
        return { type: 'String', value: toString(left) + toString(right) };
      case 'Is':
        return {
          type: 'Boolean',
          value: left.type === 'Object' && right.type === 'Object' && left.value === right.value,
        };
      default:
        return VbEmpty;
    }
  }

  private add(left: VbValue, right: VbValue): VbValue {
    if (left.type === 'Null' || right.type === 'Null') {
      return VbNull;
    }
    if (left.type === 'String' || right.type === 'String') {
      return { type: 'String', value: toString(left) + toString(right) };
    }
    if (left.type === 'LongLong' || right.type === 'LongLong') {
      return this.longLongArithmetic(left, right, '+');
    }
    const leftNum = toNumber(left);
    const rightNum = toNumber(right);
    return createVbValue(leftNum + rightNum);
  }

  private subtract(left: VbValue, right: VbValue): VbValue {
    if (left.type === 'Null' || right.type === 'Null') {
      return VbNull;
    }
    if (left.type === 'LongLong' || right.type === 'LongLong') {
      return this.longLongArithmetic(left, right, '-');
    }
    const leftNum = toNumber(left);
    const rightNum = toNumber(right);
    return createVbValue(leftNum - rightNum);
  }

  private multiply(left: VbValue, right: VbValue): VbValue {
    if (left.type === 'Null' || right.type === 'Null') {
      return VbNull;
    }
    if (left.type === 'LongLong' || right.type === 'LongLong') {
      return this.longLongArithmetic(left, right, '*');
    }
    const leftNum = toNumber(left);
    const rightNum = toNumber(right);
    return createVbValue(leftNum * rightNum);
  }

  private divide(left: VbValue, right: VbValue): VbValue {
    if (left.type === 'Null' || right.type === 'Null') {
      return VbNull;
    }
    const leftNum = toNumber(left);
    const rightNum = toNumber(right);
    if (rightNum === 0) {
      throw createVbError(VbErrorCodes.DivisionByZero, 'Division by zero');
    }
    return { type: 'Double', value: leftNum / rightNum };
  }

  private integerDivide(left: VbValue, right: VbValue): VbLongValue {
    if (left.type === 'Null' || right.type === 'Null') {
      return VbNull as unknown as VbLongValue;
    }
    const leftNum = toNumber(left);
    const rightNum = toNumber(right);
    if (rightNum === 0) {
      throw createVbError(VbErrorCodes.DivisionByZero, 'Division by zero');
    }
    return { type: 'Long', value: Math.floor(leftNum / rightNum) };
  }

  private modulo(left: VbValue, right: VbValue): VbValue {
    if (left.type === 'Null' || right.type === 'Null') {
      return VbNull;
    }
    const leftNum = toNumber(left);
    const rightNum = toNumber(right);
    if (rightNum === 0) {
      throw createVbError(VbErrorCodes.DivisionByZero, 'Division by zero');
    }
    return createVbValue(leftNum % rightNum);
  }

  private power(left: VbValue, right: VbValue): VbValue {
    if (left.type === 'Null' || right.type === 'Null') {
      return VbNull;
    }
    const leftNum = toNumber(left);
    const rightNum = toNumber(right);
    return { type: 'Double', value: Math.pow(leftNum, rightNum) };
  }

  // Real VBScript propagates Null through every comparison operator (=, <>,
  // <, <=, >, >=) - the result is Null, not a definite True/False - the same
  // "Null infects the result" rule this engine already applied to arithmetic
  // (add/subtract/etc. above). `equals` previously hardcoded `false` for a
  // Null operand instead of propagating, and lessThan/lessThanOrEqual/
  // greaterThan/greaterThanOrEqual had no Null handling at all, so comparing
  // anything against Null (`x < Null`) crashed with "Type mismatch: Null
  // cannot be converted to Number" instead of yielding Null. Found via
  // Wine's own vbscript.dll conformance suite, dlls/vbscript/tests/lang.vbs.
  //
  // Two more real VBScript comparison quirks, both confirmed against Wine's
  // lang.vbs comments/assertions rather than guessed: (1) an Array operand
  // on EITHER side of any comparison always raises type mismatch (error 13)
  // - this wins even over Null propagation and Empty/Object special-casing,
  // so it must be checked first. (2) Boolean compared against String never
  // does numeric coercion or Null-style handling: it's always a plain
  // string comparison against CStr(bool) ("True"/"False"), case-sensitive,
  // no trimming, never errors - so it must be checked before the generic
  // String branch below (which lowercases for case-insensitive equality,
  // wrong for this specific combination).
  private checkComparisonOperands(left: VbValue, right: VbValue): void {
    if (left.type === 'Array' || right.type === 'Array') {
      throw createVbError(VbErrorCodes.TypeMismatch, 'Type mismatch');
    }
  }

  private compareBooleanString(left: VbValue, right: VbValue): number | undefined {
    if (
      (left.type === 'Boolean' && right.type === 'String') ||
      (left.type === 'String' && right.type === 'Boolean')
    ) {
      const ls = toString(left);
      const rs = toString(right);
      return ls < rs ? -1 : ls > rs ? 1 : 0;
    }
    return undefined;
  }

  private equals(left: VbValue, right: VbValue): VbValue {
    this.checkComparisonOperands(left, right);
    if (left.type === 'Empty' && right.type === 'Empty') {
      return { type: 'Boolean', value: true };
    }
    if (left.type === 'Null' || right.type === 'Null') {
      return VbNull;
    }
    if (left.type === 'Object' && right.type === 'Object') {
      return { type: 'Boolean', value: left.value === right.value };
    }
    const boolStr = this.compareBooleanString(left, right);
    if (boolStr !== undefined) {
      return { type: 'Boolean', value: boolStr === 0 };
    }
    if (left.type === 'String' || right.type === 'String') {
      return {
        type: 'Boolean',
        value: toString(left).toLowerCase() === toString(right).toLowerCase(),
      };
    }
    return { type: 'Boolean', value: toNumber(left) === toNumber(right) };
  }

  private notEquals(left: VbValue, right: VbValue): VbValue {
    const eq = this.equals(left, right);
    if (eq.type === 'Null') return VbNull;
    return { type: 'Boolean', value: !toBoolean(eq) };
  }

  private lessThan(left: VbValue, right: VbValue): VbValue {
    this.checkComparisonOperands(left, right);
    if (left.type === 'Null' || right.type === 'Null') {
      return VbNull;
    }
    const boolStr = this.compareBooleanString(left, right);
    if (boolStr !== undefined) {
      return { type: 'Boolean', value: boolStr < 0 };
    }
    if (left.type === 'String' || right.type === 'String') {
      return { type: 'Boolean', value: toString(left) < toString(right) };
    }
    return { type: 'Boolean', value: toNumber(left) < toNumber(right) };
  }

  private lessThanOrEqual(left: VbValue, right: VbValue): VbValue {
    this.checkComparisonOperands(left, right);
    if (left.type === 'Null' || right.type === 'Null') {
      return VbNull;
    }
    const boolStr = this.compareBooleanString(left, right);
    if (boolStr !== undefined) {
      return { type: 'Boolean', value: boolStr <= 0 };
    }
    if (left.type === 'String' || right.type === 'String') {
      return { type: 'Boolean', value: toString(left) <= toString(right) };
    }
    return { type: 'Boolean', value: toNumber(left) <= toNumber(right) };
  }

  private greaterThan(left: VbValue, right: VbValue): VbValue {
    this.checkComparisonOperands(left, right);
    if (left.type === 'Null' || right.type === 'Null') {
      return VbNull;
    }
    const boolStr = this.compareBooleanString(left, right);
    if (boolStr !== undefined) {
      return { type: 'Boolean', value: boolStr > 0 };
    }
    if (left.type === 'String' || right.type === 'String') {
      return { type: 'Boolean', value: toString(left) > toString(right) };
    }
    return { type: 'Boolean', value: toNumber(left) > toNumber(right) };
  }

  private greaterThanOrEqual(left: VbValue, right: VbValue): VbValue {
    this.checkComparisonOperands(left, right);
    if (left.type === 'Null' || right.type === 'Null') {
      return VbNull;
    }
    const boolStr = this.compareBooleanString(left, right);
    if (boolStr !== undefined) {
      return { type: 'Boolean', value: boolStr >= 0 };
    }
    if (left.type === 'String' || right.type === 'String') {
      return { type: 'Boolean', value: toString(left) >= toString(right) };
    }
    return { type: 'Boolean', value: toNumber(left) >= toNumber(right) };
  }

  private evaluateUnary(node: UnaryExpression): VbValue {
    const argument = this.evaluate(node.argument);

    switch (node.operator) {
      case '-':
        // Null propagates through unary minus too (`-Null` is Null, not a
        // crash) - same reasoning as the binary arithmetic/comparison Null
        // handling elsewhere in this file. Found via Wine's own vbscript.dll
        // conformance suite, dlls/vbscript/tests/lang.vbs: `getVT(-null)`.
        if (argument.type === 'Null') return VbNull;
        return createVbValue(-toNumber(argument));
      case '+':
        if (argument.type === 'Null') return VbNull;
        return createVbValue(toNumber(argument));
      case '!':
      case 'Not':
        // Null propagates through Not too (`Not Null` is Null).
        if (argument.type === 'Null') return VbNull;
        // Real VBScript's Not is VarNot - a genuine BITWISE complement, not
        // a boolean negation, except when the operand is already strictly
        // Boolean-typed (where bitwise complement of -1/0 happens to equal
        // boolean negation anyway, so this is just a type-preservation
        // distinction, not a value one). `Not 5` is the Long -6, not a
        // Boolean. Confirmed via Wine's own vbscript.dll conformance suite,
        // dlls/vbscript/tests/lang.vbs: `(Not Empty) = -1,
        // getVT(Not Empty) = VT_I4`.
        if (argument.type === 'Boolean') {
          return { type: 'Boolean', value: !argument.value };
        }
        return { type: 'Long', value: ~toBitwiseLong(argument) };
      default:
        return VbEmpty;
    }
  }

  // Real VBScript's And/Or/Xor/Eqv/Imp are three-valued (True/False/Null),
  // not JS's two-valued short-circuit logic - `Null` propagates through all
  // five, EXCEPT And/Or's documented "absorbing value" special case (`False
  // And Null = False`, `True Or Null = True`, since the result is already
  // determined regardless of what the Null side would have been). This also
  // means And/Or can no longer short-circuit: `Null Or True` must actually
  // see the True on the right before it can resolve to True rather than
  // Null. Found via Wine's own vbscript.dll conformance suite, dlls/
  // vbscript/tests/lang.vbs - `Null Or True` previously crashed outright
  // (toBoolean(Null) throws) since the old short-circuit form evaluated
  // toBoolean(left) before ever looking at the right operand.
  private evaluateLogical(node: LogicalExpression): VbValue {
    switch (node.operator) {
      case '&&':
      case 'And':
        return this.applyAnd(this.evaluate(node.left), this.evaluate(node.right));
      case '||':
      case 'Or':
        return this.applyOr(this.evaluate(node.left), this.evaluate(node.right));
      case 'Xor': {
        const left = this.evaluate(node.left);
        const right = this.evaluate(node.right);
        if (left.type === 'Null' || right.type === 'Null') return VbNull;
        return makeBitwiseResult(toBitwiseLong(left) ^ toBitwiseLong(right), left, right);
      }
      case 'Eqv': {
        const left = this.evaluate(node.left);
        const right = this.evaluate(node.right);
        if (left.type === 'Null' || right.type === 'Null') return VbNull;
        return makeBitwiseResult(~(toBitwiseLong(left) ^ toBitwiseLong(right)), left, right);
      }
      case 'Imp': {
        // A Imp B == (Not A) Or B - real VBScript's own documented truth
        // table for Imp matches this composition exactly, including how
        // Null propagates through it (e.g. `False Imp Null = True`: False
        // negates to True, which is Or's own absorbing value). Delegating to
        // the already-fixed applyNot()/applyOr() rather than hand-duplicating
        // their logic (absorption, bitwise math, result typing) a third time.
        const left = this.evaluate(node.left);
        const right = this.evaluate(node.right);
        return this.applyOr(this.applyNot(left), right);
      }
      default:
        return VbEmpty;
    }
  }

  private applyNot(argument: VbValue): VbValue {
    if (argument.type === 'Null') return VbNull;
    if (argument.type === 'Boolean') return { type: 'Boolean', value: !argument.value };
    return { type: 'Long', value: ~toBitwiseLong(argument) };
  }

  // Real VBScript's And is three-valued (True/False/Null) AND genuinely
  // bitwise for non-Boolean operands - `5 And 3` is `1`, not a JS-style
  // "pick a value" result. Null only ever propagates as Null UNLESS the
  // OTHER side is already the absorbing value (a definite falsy 0/False),
  // in which case the whole result is decided regardless of the Null side -
  // and that absorption returns the ACTUAL absorbing operand itself
  // (preserving its real type, e.g. `CInt(0) And Null` stays Integer-typed
  // 0), not a synthetic Boolean. Confirmed via Wine's own vbscript.dll
  // conformance suite, dlls/vbscript/tests/lang.vbs, and Wine's real C
  // source, dlls/vbscript/interp.c's interp_and()/coerce_empty_to_i4()/
  // coerce_str_to_num_or_bool() (which is where the Empty->Long(0) and
  // String->number-or-boolean pre-coercions toBitwiseLong() replicates come
  // from - native VarAnd/VarOr/VarImp don't do those themselves).
  private applyAnd(left: VbValue, right: VbValue): VbValue {
    if (left.type === 'Null' || right.type === 'Null') {
      if (left.type !== 'Null' && !toBoolean(left)) return left;
      if (right.type !== 'Null' && !toBoolean(right)) return right;
      return VbNull;
    }
    return makeBitwiseResult(toBitwiseLong(left) & toBitwiseLong(right), left, right);
  }

  // Mirror of applyAnd() above for Or - True is the absorbing value instead
  // of False, same reasoning and same source (Wine's interp_or()).
  private applyOr(left: VbValue, right: VbValue): VbValue {
    if (left.type === 'Null' || right.type === 'Null') {
      if (left.type !== 'Null' && toBoolean(left)) return left;
      if (right.type !== 'Null' && toBoolean(right)) return right;
      return VbNull;
    }
    return makeBitwiseResult(toBitwiseLong(left) | toBitwiseLong(right), left, right);
  }

  private evaluateAssignment(node: AssignmentExpression): VbValue {
    const value = this.evaluate(node.right);
    const isSet = node.isSet ?? false;

    if (node.left.type === 'Identifier') {
      if (isSet) {
        const oldValue = this.context.getVariable(node.left.name);
        this.callTerminateIfNeeded(oldValue, value);
      }
      this.context.setVariable(node.left.name, value);
      return value;
    }

    if (node.left.type === 'MemberExpression') {
      this.assignToMember(node.left, value, isSet);
      return value;
    }

    throw new Error(`Invalid assignment target: ${node.left.type}`);
  }

  private callTerminateIfNeeded(oldValue: VbValue, newValue: VbValue): void {
    if (oldValue.type === 'Object' && oldValue.value instanceof VbObjectInstance) {
      if (newValue.type !== 'Object' || newValue.value !== oldValue.value) {
        const instance = oldValue.value as VbObjectInstance;
        const terminateProp = instance.classInfo.properties.get('class_terminate');
        if (terminateProp && terminateProp.get) {
          terminateProp.get.call(instance);
        }
      }
    }
  }

  private assignToMember(node: MemberExpression, value: VbValue, isSet: boolean): void {
    let object = this.evaluate(node.object as Expression);

    if (node.object.type === 'VbWithObject') {
      object = this.evaluateWithObject(node.object as VbWithObjectExpression);
    }

    let propertyName: string;
    if (node.computed) {
      const propValue = this.evaluate(node.property as Expression);
      propertyName = toString(propValue);
    } else {
      propertyName = (node.property as Identifier).name;
    }

    if (object.type === 'Array') {
      const arr = object.value as unknown as { set: (indices: number[], v: VbValue) => void };
      const subscripts = node.vbIndices ?? [node.property as Expression];
      arr.set(
        subscripts.map((subscript) => Math.floor(toNumber(this.evaluate(subscript)))),
        value
      );
    } else if (object.type === 'Object') {
      const obj = object.value as VbObjectValueData | null;
      if (obj === null || (typeof obj !== 'object' && typeof obj !== 'function')) {
        throw createVbError(VbErrorCodes.ObjectRequired, 'Object required');
      }

      if (
        (Object.prototype.hasOwnProperty.call(obj, 'setProperty') ||
          Object.prototype.hasOwnProperty.call(obj, 'classInfo')) &&
        typeof obj.setProperty === 'function'
      ) {
        obj.setProperty(propertyName, value, isSet);
      } else {
        const jsValue = this.vbToJs(value);
        obj[propertyName] = jsValue;
      }
    } else {
      throw createVbError(VbErrorCodes.ObjectRequired, 'Object required');
    }
  }

  private evaluateConditional(node: ConditionalExpression): VbValue {
    const test = this.evaluate(node.test);
    if (toBoolean(test)) {
      return this.evaluate(node.consequent);
    }
    return this.evaluate(node.alternate);
  }

  private toLongLongBigInt(v: VbValue): bigint {
    if (v.type === 'LongLong') return v.value;
    if (v.type === 'Integer' || v.type === 'Long' || v.type === 'Byte') return BigInt(v.value);
    if (v.type === 'Boolean') return v.value ? BigInt(-1) : BigInt(0);
    if (v.type === 'Empty') return BigInt(0);
    // Float types demote to Double for mixed arithmetic
    return BigInt(Math.trunc(toNumber(v)));
  }

  private longLongArithmetic(left: VbValue, right: VbValue, op: '+' | '-' | '*' | '\\'): VbValue {
    // If either side is floating-point, promote to Double
    const floatTypes = ['Single', 'Double', 'Currency'];
    if (floatTypes.includes(left.type) || floatTypes.includes(right.type)) {
      const l = toNumber(left);
      const r = toNumber(right);
      if (op === '+') return { type: 'Double', value: l + r };
      if (op === '-') return { type: 'Double', value: l - r };
      if (op === '*') return { type: 'Double', value: l * r };
      if (r === 0) throw createVbError(VbErrorCodes.DivisionByZero, 'Division by zero');
      return { type: 'Long', value: Math.floor(l / r) };
    }
    const l = this.toLongLongBigInt(left);
    const r = this.toLongLongBigInt(right);
    if (op === '+') return { type: 'LongLong', value: l + r };
    if (op === '-') return { type: 'LongLong', value: l - r };
    if (op === '*') return { type: 'LongLong', value: l * r };
    if (r === BigInt(0))
      throw createVbError(VbErrorCodes.DivisionByZero, 'Division by zero');
    return { type: 'LongLong', value: l / r };
  }
}
