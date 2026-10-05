import { describe, it, expect } from 'vitest';
import { VbsEngine } from '../index.ts';

// ---------------------------------------------------------------------------
// VB6 Type Syntax (As Type in declarations and parameters)
// ---------------------------------------------------------------------------
describe('VB6 Type Syntax', () => {
  it('Dim with As Integer initializes to 0', () => {
    const engine = new VbsEngine();
    engine.executeStatement('Dim x As Integer');
    expect(engine._getVariable('x').type).toBe('Integer');
    expect(engine._getVariable('x').value).toBe(0);
  });

  it('Dim with As Long initializes to 0', () => {
    const engine = new VbsEngine();
    engine.executeStatement('Dim x As Long');
    expect(engine._getVariable('x').type).toBe('Long');
    expect(engine._getVariable('x').value).toBe(0);
  });

  it('Dim with As String initializes to empty string', () => {
    const engine = new VbsEngine();
    engine.executeStatement('Dim s As String');
    expect(engine._getVariable('s').type).toBe('String');
    expect(engine._getVariable('s').value).toBe('');
  });

  it('Dim with As Boolean initializes to False', () => {
    const engine = new VbsEngine();
    engine.executeStatement('Dim b As Boolean');
    expect(engine._getVariable('b').type).toBe('Boolean');
    expect(engine._getVariable('b').value).toBe(false);
  });

  it('Dim with As Double initializes to 0', () => {
    const engine = new VbsEngine();
    engine.executeStatement('Dim d As Double');
    expect(engine._getVariable('d').type).toBe('Double');
    expect(engine._getVariable('d').value).toBe(0);
  });

  it('Dim with As Byte initializes to 0', () => {
    const engine = new VbsEngine();
    engine.executeStatement('Dim b As Byte');
    expect(engine._getVariable('b').type).toBe('Byte');
    expect(engine._getVariable('b').value).toBe(0);
  });

  it('Dim with As Object initializes to Nothing', () => {
    const engine = new VbsEngine();
    engine.executeStatement('Set obj = Nothing');
    const v = engine._getVariable('obj');
    expect(v.type).toBe('Object');
    expect(v.value).toBeNull();
  });

  it('Function with typed parameter and return type parses correctly', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Function Add(ByVal a As Integer, ByVal b As Integer) As Long
        Add = a + b
      End Function
      result = Add(3, 4)
    `);
    expect(engine._getVariable('result').value).toBe(7);
  });

  it('Sub with typed parameters works', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Sub Greet(ByVal name As String)
        greeting = "Hello " & name
      End Sub
      Greet("World")
    `);
    expect(engine._getVariable('greeting').value).toBe('Hello World');
  });

  it('Property Get with As Type parses correctly', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Class Counter
        Private m_count As Integer
        Property Get Count() As Integer
          Count = m_count
        End Property
        Property Let Count(ByVal v As Integer)
          m_count = v
        End Property
      End Class
      Set c = New Counter
      c.Count = 5
      result = c.Count
    `);
    expect(engine._getVariable('result').value).toBe(5);
  });
});

// ---------------------------------------------------------------------------
// VB6 Enums
// ---------------------------------------------------------------------------
describe('VB6 Enums', () => {
  it('Enum members are declared as constants with auto-increment from 0', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Enum Color
        Red
        Green
        Blue
      End Enum
    `);
    expect(engine._getVariable('Red').value).toBe(0);
    expect(engine._getVariable('Green').value).toBe(1);
    expect(engine._getVariable('Blue').value).toBe(2);
  });

  it('Enum members with explicit values', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Enum Status
        Pending = 10
        Active = 20
        Inactive = 30
      End Enum
    `);
    expect(engine._getVariable('Pending').value).toBe(10);
    expect(engine._getVariable('Active').value).toBe(20);
    expect(engine._getVariable('Inactive').value).toBe(30);
  });

  it('Enum auto-increment after explicit value', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Enum Priority
        Low
        Medium = 10
        High
        Critical
      End Enum
    `);
    expect(engine._getVariable('Low').value).toBe(0);
    expect(engine._getVariable('Medium').value).toBe(10);
    expect(engine._getVariable('High').value).toBe(11);
    expect(engine._getVariable('Critical').value).toBe(12);
  });

  it('Enum values can be used in expressions', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Enum Direction
        North = 0
        East = 90
        South = 180
        West = 270
      End Enum
      heading = East + 45
    `);
    expect(engine._getVariable('heading').value).toBe(135);
  });

  it('Enum values can be used in Select Case', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Enum Fruit
        Apple
        Banana
        Cherry
      End Enum
      pick = Banana
      Select Case pick
        Case Apple
          result = "apple"
        Case Banana
          result = "banana"
        Case Cherry
          result = "cherry"
      End Select
    `);
    expect(engine._getVariable('result').value).toBe('banana');
  });

  it('Public Enum is accessible globally', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Public Enum Weekday2
        Mon = 1
        Tue = 2
        Wed = 3
      End Enum
      day = Wed
    `);
    expect(engine._getVariable('day').value).toBe(3);
  });
});

// ---------------------------------------------------------------------------
// VB6 Collection
// ---------------------------------------------------------------------------
describe('VB6 Collection', () => {
  it('New Collection creates an empty collection', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Set col = New Collection
      result = col.Count
    `);
    expect(engine._getVariable('result').value).toBe(0);
  });

  it('Collection.Add increases Count', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Set col = New Collection
      col.Add "Alpha"
      col.Add "Beta"
      result = col.Count
    `);
    expect(engine._getVariable('result').value).toBe(2);
  });

  it('Collection.Item with 1-based index', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Set col = New Collection
      col.Add "First"
      col.Add "Second"
      col.Add "Third"
      r1 = col.Item(1)
      r2 = col.Item(2)
      r3 = col.Item(3)
    `);
    expect(engine._getVariable('r1').value).toBe('First');
    expect(engine._getVariable('r2').value).toBe('Second');
    expect(engine._getVariable('r3').value).toBe('Third');
  });

  it('Collection.Add with string key allows Item lookup by key', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Set col = New Collection
      col.Add "Paris", "FR"
      col.Add "London", "UK"
      col.Add "Berlin", "DE"
      r1 = col.Item("FR")
      r2 = col.Item("UK")
      r3 = col.Item("DE")
    `);
    expect(engine._getVariable('r1').value).toBe('Paris');
    expect(engine._getVariable('r2').value).toBe('London');
    expect(engine._getVariable('r3').value).toBe('Berlin');
  });

  it('Collection.Remove by index', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Set col = New Collection
      col.Add "A"
      col.Add "B"
      col.Add "C"
      col.Remove 2
      result = col.Count
      r1 = col.Item(1)
      r2 = col.Item(2)
    `);
    expect(engine._getVariable('result').value).toBe(2);
    expect(engine._getVariable('r1').value).toBe('A');
    expect(engine._getVariable('r2').value).toBe('C');
  });

  it('Collection.Remove by key', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Set col = New Collection
      col.Add "Paris", "FR"
      col.Add "London", "UK"
      col.Remove "FR"
      result = col.Count
      r1 = col.Item(1)
    `);
    expect(engine._getVariable('result').value).toBe(1);
    expect(engine._getVariable('r1').value).toBe('London');
  });

  it('For Each iterates all Collection items', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Set col = New Collection
      col.Add 10
      col.Add 20
      col.Add 30
      total = 0
      For Each item In col
        total = total + item
      Next
    `);
    expect(engine._getVariable('total').value).toBe(60);
  });

  it('Collection stores mixed types', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Set col = New Collection
      col.Add 42
      col.Add "hello"
      col.Add True
      n = col.Item(1)
      s = col.Item(2)
      b = col.Item(3)
    `);
    expect(engine._getVariable('n').value).toBe(42);
    expect(engine._getVariable('s').value).toBe('hello');
    expect(engine._getVariable('b').value).toBe(true);
  });

  it('TypeName of Collection is "Collection"', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Set col = New Collection
      result = TypeName(col)
    `);
    expect(engine._getVariable('result').value).toBe('Collection');
  });
});

// ---------------------------------------------------------------------------
// VBA7 LongLong type and 64-bit integer support
// ---------------------------------------------------------------------------
describe('VBA7 LongLong type', () => {
  it('Dim with As LongLong initializes to BigInt(0)', () => {
    const engine = new VbsEngine();
    engine.executeStatement('Dim x As LongLong');
    const v = engine._getVariable('x');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt(0));
  });

  it('CLngLng converts integer to LongLong', () => {
    const engine = new VbsEngine();
    engine.executeStatement('x = CLngLng(42)');
    const v = engine._getVariable('x');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt(42));
  });

  it('CLngLng converts string to LongLong', () => {
    const engine = new VbsEngine();
    engine.executeStatement('x = CLngLng("9223372036854775807")');
    const v = engine._getVariable('x');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt('9223372036854775807'));
  });

  it('CLngLng converts negative value', () => {
    const engine = new VbsEngine();
    engine.executeStatement('x = CLngLng(-100)');
    const v = engine._getVariable('x');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt(-100));
  });

  it('LongLong + LongLong = LongLong', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      a = CLngLng(1000000000)
      b = CLngLng(2000000000)
      c = a + b
    `);
    const v = engine._getVariable('c');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt(3000000000));
  });

  it('LongLong - LongLong = LongLong', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      a = CLngLng(5000000000)
      b = CLngLng(3000000000)
      c = a - b
    `);
    const v = engine._getVariable('c');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt(2000000000));
  });

  it('LongLong * LongLong = LongLong', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      a = CLngLng(1000000)
      b = CLngLng(1000000)
      c = a * b
    `);
    const v = engine._getVariable('c');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt(1000000000000));
  });

  it('LongLong + Integer = LongLong', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      a = CLngLng(9000000000)
      c = a + 1
    `);
    const v = engine._getVariable('c');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt(9000000001));
  });

  it('LongLong + Double promotes to Double', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      a = CLngLng(100)
      c = a + 1.5
    `);
    const v = engine._getVariable('c');
    expect(v.type).toBe('Double');
    expect(v.value).toBeCloseTo(101.5);
  });

  it('LongLong comparison works', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      a = CLngLng(9000000000)
      b = CLngLng(9000000000)
      If a = b Then eq = True Else eq = False
      If CLngLng(1) < CLngLng(2) Then lt = True Else lt = False
    `);
    expect(engine._getVariable('eq').value).toBe(true);
    expect(engine._getVariable('lt').value).toBe(true);
  });

  it('VarType of LongLong is 20', () => {
    const engine = new VbsEngine();
    engine.executeStatement('x = VarType(CLngLng(0))');
    expect(engine._getVariable('x').value).toBe(20);
  });

  it('TypeName of LongLong is "LongLong"', () => {
    const engine = new VbsEngine();
    engine.executeStatement('x = TypeName(CLngLng(0))');
    expect(engine._getVariable('x').value).toBe('LongLong');
  });

  it('CStr converts LongLong to string', () => {
    const engine = new VbsEngine();
    engine.executeStatement('x = CStr(CLngLng("9223372036854775807"))');
    expect(engine._getVariable('x').value).toBe('9223372036854775807');
  });

  it('LongLong integer division (\\)', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      a = CLngLng(10000000000)
      b = CLngLng(3)
      c = a \\ b
    `);
    const v = engine._getVariable('c');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt(3333333333));
  });

  it('64-bit range: max LongLong via string', () => {
    const engine = new VbsEngine();
    engine.executeStatement('x = CLngLng("9223372036854775807")');
    const v = engine._getVariable('x');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt('9223372036854775807'));
  });
});

// ---------------------------------------------------------------------------
// VB6 User-Defined Types (Type...End Type)
// ---------------------------------------------------------------------------
describe('VB6 User-Defined Types', () => {
  it('Type declaration and Dim creates instance with default-initialized fields', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Type Point
        X As Integer
        Y As Integer
      End Type
      Dim p As Point
    `);
    const p = engine._getVariable('p');
    expect(p.type).toBe('Object');
    const instance = p.value as import('../runtime/class-registry.ts').VbObjectInstance;
    expect(instance.getProperty('X').type).toBe('Integer');
    expect(instance.getProperty('X').value).toBe(0);
    expect(instance.getProperty('Y').type).toBe('Integer');
    expect(instance.getProperty('Y').value).toBe(0);
  });

  it('Field access and assignment', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Type Point
        X As Integer
        Y As Integer
      End Type
      Dim p As Point
      p.X = 10
      p.Y = 20
      result = p.X + p.Y
    `);
    expect(engine._getVariable('result').value).toBe(30);
  });

  it('String and Boolean fields', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Type Person
        Name As String
        Active As Boolean
        Age As Integer
      End Type
      Dim per As Person
      per.Name = "Alice"
      per.Active = True
      per.Age = 30
    `);
    const per = engine._getVariable('per');
    const instance = per.value as import('../runtime/class-registry.ts').VbObjectInstance;
    expect(instance.getProperty('Name').value).toBe('Alice');
    expect(instance.getProperty('Active').value).toBe(true);
    expect(instance.getProperty('Age').value).toBe(30);
  });

  it('Multiple independent instances', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Type Point
        X As Integer
        Y As Integer
      End Type
      Dim a As Point
      Dim b As Point
      a.X = 1
      b.X = 99
    `);
    const a = engine._getVariable('a')
      .value as import('../runtime/class-registry.ts').VbObjectInstance;
    const b = engine._getVariable('b')
      .value as import('../runtime/class-registry.ts').VbObjectInstance;
    expect(a.getProperty('X').value).toBe(1);
    expect(b.getProperty('X').value).toBe(99);
  });

  it('Fields default to type-appropriate zero values', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Type Defaults
        N As Long
        S As String
        B As Boolean
        D As Double
      End Type
      Dim d As Defaults
    `);
    const instance = engine._getVariable('d')
      .value as import('../runtime/class-registry.ts').VbObjectInstance;
    expect(instance.getProperty('N').value).toBe(0);
    expect(instance.getProperty('S').value).toBe('');
    expect(instance.getProperty('B').value).toBe(false);
    expect(instance.getProperty('D').value).toBe(0);
  });

  it('UDT used in Sub parameter (ByRef)', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Type Vec
        X As Double
        Y As Double
      End Type
      Sub Scale(v As Vec, factor As Double)
        v.X = v.X * factor
        v.Y = v.Y * factor
      End Sub
      Dim v As Vec
      v.X = 3
      v.Y = 4
      Scale v, 2
      rx = v.X
      ry = v.Y
    `);
    expect(engine._getVariable('rx').value).toBe(6);
    expect(engine._getVariable('ry').value).toBe(8);
  });
});

// ---------------------------------------------------------------------------
// Keywords usable as ordinary variable names
// ---------------------------------------------------------------------------
// These assert the STORED VALUE rather than that the line parses. Parsing alone is not the
// property that matters here: a statement led by a keyword that is never recognised as an
// assignment still parses - as a comparison - and then discards its value, leaving the variable
// unwritten with nothing raised. A parse-only assertion cannot tell that apart from a real fix.
describe('Keyword-named variables', () => {
  it.each<[string, number]>([
    ['Step', 4],
    ['Error', 7],
    ['Object', 9],
  ])('assigns to %s and reads the value back', (word, value) => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Dim ${word}
      ${word} = ${value}
    `);
    expect(engine._getVariable(word).value).toBe(value);
  });

  it('takes the branch the stored value selects', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Dim Step
      Step = 4
      If Step = 4 Then
        taken = "yes"
      Else
        taken = "no"
      End If
    `);
    expect(engine._getVariable('taken').value).toBe('yes');
  });
});

// ---------------------------------------------------------------------------
// Assigning to a property written with empty parentheses
// ---------------------------------------------------------------------------
// `o.P() = v` is the parenthesised spelling of `o.P = v` and the real engine accepts it. A bare
// `a() = v` is NOT the same thing - there it raises a runtime Type mismatch - so the negative
// case below is as much a part of this fix as the positive ones: the point is to admit the
// member form without also admitting the identifier form.
describe('Empty-parenthesis assignment targets', () => {
  it('assigns to a public field written with empty parentheses', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Class C
        Public P
      End Class
      Dim o
      Set o = New C
      o.P() = 5
      r = o.P
    `);
    expect(engine._getVariable('r').value).toBe(5);
  });

  it('assigns through a Property Let written with empty parentheses', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Class C
        Private m
        Public Property Let P(v)
          m = v * 2
        End Property
        Public Property Get P
          P = m
        End Property
      End Class
      Dim o
      Set o = New C
      o.P() = 5
      r = o.P
    `);
    expect(engine._getVariable('r').value).toBe(10);
  });

  it('still refuses empty parentheses on a plain variable', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Dim a
      a() = 5
    `);
    // The error itself is not pinned: the real engine says Type mismatch here and this one does
    // not, and pinning the wording would pin that divergence rather than the behaviour.
    expect(engine.error).not.toBeNull();
  });
});

// ---------------------------------------------------------------------------
// Assigning to a multi-dimensional array element
// ---------------------------------------------------------------------------
// The transposed pair is the test that matters. An implementation that collapses the subscripts
// into one - or that keeps only the first, as this one did - still passes a single-cell check,
// because writing and reading agree with each other on the wrong cell. Two cells that differ only
// in the order of their subscripts must hold different values.
describe('Multi-dimensional array assignment', () => {
  it('writes and reads a two-dimensional element', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Dim a(3,3)
      a(1,2) = 7
      r = a(1,2)
    `);
    expect(engine._getVariable('r').value).toBe(7);
  });

  it('keeps transposed cells apart', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Dim a(3,3)
      a(1,2) = 7
      a(2,1) = 9
      r1 = a(1,2)
      r2 = a(2,1)
    `);
    expect(engine._getVariable('r1').value).toBe(7);
    expect(engine._getVariable('r2').value).toBe(9);
  });

  it('writes and reads a three-dimensional element', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Dim b(2,2,2)
      b(1,0,1) = 5
      r = b(1,0,1)
    `);
    expect(engine._getVariable('r').value).toBe(5);
  });

  it('leaves the single-subscript write alone', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Dim a(3)
      a(2) = 4
      r = a(2)
    `);
    expect(engine._getVariable('r').value).toBe(4);
  });
});

// ---------------------------------------------------------------------------
// Line continuation with trailing whitespace
// ---------------------------------------------------------------------------
// The significant character here is a space or tab AFTER the underscore, so these sources are
// built by joining explicit strings rather than written as template literals: a trailing space at
// the end of a line in this file would not survive the formatter, and the test would quietly stop
// testing anything. The failure being pinned is silent - the statement simply ends at the
// underscore and the continued half is dropped, with nothing raised.
describe('Line continuation', () => {
  it.each<[string, string]>([
    ['no trailing whitespace', 'a = "x" & _'],
    ['a trailing space', 'a = "x" & _ '],
    ['a trailing tab', 'a = "x" & _\t'],
    ['trailing spaces and tabs', 'a = "x" & _ \t '],
  ])('continues the statement with %s', (_label, firstLine) => {
    const engine = new VbsEngine();
    engine.executeStatement(['Dim a', firstLine, '"1"'].join('\n'));
    expect(engine.error).toBeNull();
    expect(engine._getVariable('a').value).toBe('x1');
  });

  it('leaves an underscore inside an identifier alone', () => {
    const engine = new VbsEngine();
    engine.executeStatement(['Dim a_b', 'a_b = 5'].join('\n'));
    expect(engine._getVariable('a_b').value).toBe(5);
  });
});

// ---------------------------------------------------------------------------
// Ampersand before an identifier that starts with h or o
// ---------------------------------------------------------------------------
// The literal rows are as much a part of this as the concatenation rows. VBScript resolves the
// ambiguity in favour of the literal whenever a digit of the base follows, so widening the
// concatenation case must not cost `&hFF` - and `"x"&h1` stays a literal there even where a
// variable h1 exists, which is why no row claims otherwise.
describe('Ampersand before an h/o identifier', () => {
  it.each<[string, string, number]>([
    ['hex', 'r = &hFF', 255],
    ['octal with a prefix', 'r = &o17', 15],
    ['bare octal', 'r = &100', 64],
  ])('still reads a %s literal', (_label, source, value) => {
    const engine = new VbsEngine();
    engine.executeStatement(source);
    expect(engine._getVariable('r').value).toBe(value);
  });

  it.each<[string, string]>([
    ['hl', 'x7'],
    ['oShell', 'x6'],
  ])('concatenates with the variable %s', (name, expected) => {
    const engine = new VbsEngine();
    engine.executeStatement([`Dim ${name}`, `${name} = ${expected.slice(1)}`, `r = "x"&${name}`].join('\n'));
    expect(engine.error).toBeNull();
    expect(engine._getVariable('r').value).toBe(expected);
  });

  it('concatenates with a member call on such an identifier', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Class C
        Public Function M()
          M = 3
        End Function
      End Class
      Dim hlObj
      Set hlObj = New C
      r = "Nr: "&hlObj.M()
    `);
    expect(engine._getVariable('r').value).toBe('Nr: 3');
  });
});

// ---------------------------------------------------------------------------
// A colon in front of the token that ends a single-line If
// ---------------------------------------------------------------------------
// `: ElseIf` is NOT part of this: the real engine rejects it ("Must be first statement on the
// line"), and so does this parser, so the last case pins that the widening stopped where the
// language stops. Its wording is not asserted - the two engines disagree on the message and
// only agree on the verdict.
describe('Colon before the end of a single-line If', () => {
  it.each<[string, string, number]>([
    ['End If', 'If 1 = 1 Then : x = 1 : End If', 1],
    ['Else', 'If 1 = 2 Then : x = 1 : Else : x = 2 : End If', 2],
    ['a doubled colon', 'If 1 = 1 Then :: x = 1 :: End If', 1],
    ['an empty body', 'If 1 = 1 Then : End If', 0],
    ['no colon at all', 'If 1 = 1 Then x = 1', 1],
  ])('parses a colon before %s', (_label, source, expected) => {
    const engine = new VbsEngine();
    engine.executeStatement(['Dim x', 'x = 0', source].join('\n'));
    expect(engine.error).toBeNull();
    expect(engine._getVariable('x').value).toBe(expected);
  });

  it('still refuses ElseIf after a colon', () => {
    const engine = new VbsEngine();
    engine.executeStatement(
      ['Dim x', 'x = 0', 'If 1 = 2 Then : x = 1 : ElseIf 1 = 1 Then : x = 3 : End If'].join('\n')
    );
    expect(engine.error).not.toBeNull();
  });
});

// ---------------------------------------------------------------------------
// Do ... Loop in every form
// ---------------------------------------------------------------------------
// Only the first row fails without the fix; the other five are the forms that must not move
// while it is made. The last one is the one a careless fix breaks: a post-test loop runs its
// body once even when the condition is false from the start, so "no condition" and "condition
// false" cannot be collapsed into the same answer. Every case terminates on its own - an
// unconditional loop with no Exit Do would hang the suite rather than fail it.
describe('Do loops', () => {
  it.each<[string, string, number]>([
    ['unconditional, left by Exit Do', 'Do\n n = n + 1\n If n >= 3 Then Exit Do\nLoop', 3],
    ['Do While, pre-test', 'Do While n < 4\n n = n + 1\nLoop', 4],
    ['Do Until, pre-test', 'Do Until n >= 4\n n = n + 1\nLoop', 4],
    ['Loop While, post-test', 'Do\n n = n + 1\nLoop While n < 4', 4],
    ['Loop Until, post-test', 'Do\n n = n + 1\nLoop Until n >= 4', 4],
    ['post-test with a false condition, one pass', 'Do\n n = n + 1\nLoop While 1 = 2', 1],
  ])('runs the body the right number of times: %s', (_label, body, expected) => {
    const engine = new VbsEngine();
    engine.executeStatement(['Dim n', 'n = 0', body].join('\n'));
    expect(engine.error).toBeNull();
    expect(engine._getVariable('n').value).toBe(expected);
  });
});

// ---------------------------------------------------------------------------
// Round
// ---------------------------------------------------------------------------
// The whole matrix was read off cscript //E:vbscript on Windows rather than derived. Two
// property pinned is that a tie goes to the even neighbour. The non-tie rows are green in both
// states and keep ordinary rounding honest.
//
// The 1.005 / 1.015 / 1.025 / 8.835 rows pin the opposite direction, and they are the reason
// this comment exists. Those literals do not land on an exact tie once scaled, and it is
// tempting to "repair" that by snapping the scaled value back to its decimal reading. VBScript
// does not: it answers 1.01 for Round(1.015, 2), which is the binary value rounding down, not
// the 1.02 a decimal reading gives. These rows fail against that well-meant extra step.
describe('Round', () => {
  it.each<[string, number]>([
    ['Round(0.5)', 0],
    ['Round(1.5)', 2],
    ['Round(2.5)', 2],
    ['Round(3.5)', 4],
    ['Round(-0.5)', 0],
    ['Round(-1.5)', -2],
    ['Round(-2.5)', -2],
    ['Round(0.125, 2)', 0.12],
    ['Round(0.135, 2)', 0.14],
    ['Round(2.675, 2)', 2.68],
    ['Round(1.2345, 3)', 1.234],
    ['Round(7.7, 0)', 8],
    ['Round(1.005, 2)', 1],
    ['Round(1.015, 2)', 1.01],
    ['Round(1.025, 2)', 1.02],
    ['Round(8.835, 2)', 8.84],
    ['Round(2.4)', 2],
    ['Round(2.6)', 3],
    ['Round(-2.6)', -3],
  ])('%s', (expression, expected) => {
    const engine = new VbsEngine();
    engine.executeStatement(`r = ${expression}`);
    expect(engine._getVariable('r').value).toBe(expected);
  });
});

// ---------------------------------------------------------------------------
// A member read without parentheses is the call
// ---------------------------------------------------------------------------
// Four of the six positive cases fail silently rather than loudly: the concatenation yields
// "[object]", the bare Sub statement simply does not run, the nested read gives Empty, and the
// host member hands back the wrapper. Only the comparison raises anything, and it raises a type
// error that names nothing useful.
//
// The Sub case asserts an instance field rather than an enclosing-scope variable on purpose -
// a class method writing an outer variable is a SEPARATE defect in this engine, and pinning it
// here would make this test fail for a reason that has nothing to do with the call.
//
// The four guards are the other half: a plain property must NOT be called, and a member that
// already has parentheses - with or without arguments - must keep reaching the real call path.
describe('Parenthesis-less member reads', () => {
  const CLASSES = [
    'Class Inner',
    '  Public Function G()',
    '    G = 8',
    '  End Function',
    'End Class',
    'Class C',
    '  Public P',
    '  Public Function F()',
    '    F = 3',
    '  End Function',
    '  Public Function Add(a)',
    '    Add = a + 1',
    '  End Function',
    '  Public Sub M()',
    '    P = 99',
    '  End Sub',
    '  Public Function MakeInner()',
    '    Set MakeInner = New Inner',
    '  End Function',
    'End Class',
    'Dim o',
    'Set o = New C',
    'o.P = 5',
  ].join('\n');

  it.each<[string, string, string, string | number]>([
    ['calls a function read without parentheses', 'r = o.F', 'r', 3],
    ['calls it as a comparison operand', 'If o.F <= 0 Then\n r = "le"\nElse\n r = "gt"\nEnd If', 'r', 'gt'],
    ['calls it as a concatenation operand', 'r = "n=" & o.F', 'r', 'n=3'],
    ['runs a Sub written as a bare statement', 'o.M', 'o.P', 99],
    ['calls through a nested member read', 'r = o.MakeInner.G', 'r', 8],
    ['GUARD leaves a plain property alone', 'r = o.P', 'r', 5],
    ['GUARD keeps arguments reaching the call', 'r = o.Add(4)', 'r', 5],
    ['GUARD keeps the parenthesised form working', 'r = o.F()', 'r', 3],
  ])('%s', (_label, tail, read, expected) => {
    const engine = new VbsEngine();
    engine.executeStatement(`${CLASSES}\n${tail}`);
    expect(engine.error).toBeNull();
    expect(engine.eval(read)).toBe(expected);
  });

  it('calls a host function exposed through addObject', () => {
    const engine = new VbsEngine();
    engine.addObject('hlObj', { GetValue: () => 42, Name: 'x' }, true);
    engine.executeStatement('r = hlObj.GetValue');
    expect(engine.error).toBeNull();
    expect(engine.eval('r')).toBe(42);
  });

  it('GUARD leaves a plain host property alone', () => {
    const engine = new VbsEngine();
    engine.addObject('hlObj', { GetValue: () => 42, Name: 'x' }, true);
    engine.executeStatement('r = hlObj.Name');
    expect(engine.eval('r')).toBe('x');
  });
});

// ---------------------------------------------------------------------------
// Naming what could not be called
// ---------------------------------------------------------------------------
// Unlike the other fixes here this one is not a conformance change - real VBScript does not
// phrase it this way either. It is about a message that fits every call in a ninety-line
// procedure body being no help in finding which one failed. Both throw sites already hold the
// name, so the test asserts it reaches the description and nothing about the wording around it.
describe('Invalid procedure call names its target', () => {
  it.each<[string, string, string]>([
    ['an identifier', 'Dim v\nv = 1\nv 1, 2', "'v'"],
    ['a member', 'Class C\nEnd Class\nDim o\nSet o = New C\no.Nope 1', "'o.Nope'"],
  ])('names %s that could not be called', (_label, source, expected) => {
    const engine = new VbsEngine();
    engine.executeStatement(source);
    expect(engine.error).not.toBeNull();
    expect(engine.error!.description).toContain(expected);
  });
});

// ---------------------------------------------------------------------------
// The error number and the position of a parse failure
// ---------------------------------------------------------------------------
// VbsError declares number, line and column, and documents line as "Source line number where
// error occurred". Every error answered -1 with neither, although the engine's own VbErrorCodes
// table already held the right codes. The numbers below are the real engine's (cscript
// //E:vbscript on Windows), and so is the position: for the same script it reports (5, 7).
describe('Error number and position', () => {
  it.each<[string, string, number]>([
    ['type mismatch', 'x = CInt("abc")', 13],
    ['division by zero', 'x = 1/0', 11],
    ['object required', 'Dim o\nx = o.Foo', 424],
    ['invalid procedure call', 'Dim v\nv = 1\nv 1, 2', 5],
  ])('reports %s as its real VBScript code', (_label, source, code) => {
    const engine = new VbsEngine();
    engine.executeStatement(source);
    expect(engine.error).not.toBeNull();
    expect(engine.error!.number).toBe(code);
  });

  it('reports the line and column a parse failure stopped at', () => {
    const engine = new VbsEngine();
    engine.executeStatement(['Dim x', 'x = 1', '', 'If 1 = 1 Then', '  x = )', 'End If'].join('\n'));
    expect(engine.error).not.toBeNull();
    expect(engine.error!.line).toBe(5);
    expect(engine.error!.column).toBe(7);
  });
});

// ---------------------------------------------------------------------------
// The position of a runtime failure
// ---------------------------------------------------------------------------
// The indented case is the one that matters: it pins the COLUMN as the statement's own start,
// not the start of the line and not the sub-expression inside it. For these two scripts cscript
// //E:vbscript reports (5, 5) and (2, 1), which is what these assert. Because every enclosing
// statement catches the same error again, only the innermost frame may stamp it - a later frame
// overwriting would move the position out to the block and quietly pass a laxer test.
describe('Runtime error position', () => {
  it('reports the failing statement, not the block around it', () => {
    const engine = new VbsEngine();
    engine.executeStatement(
      ['Dim x', 'x = 1', '', 'If 1 = 1 Then', '    x = CInt("abc")', 'End If'].join('\n')
    );
    expect(engine.error).not.toBeNull();
    expect(engine.error!.line).toBe(5);
    expect(engine.error!.column).toBe(5);
  });

  it('reports a failure at the top level', () => {
    const engine = new VbsEngine();
    engine.executeStatement(['Dim x', 'x = 1/0'].join('\n'));
    expect(engine.error!.line).toBe(2);
    expect(engine.error!.column).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// The MSScriptControl-style error surface (engine.error) vs the IE/WSH-style
// thrown native Error (executeStatementThrows)
// ---------------------------------------------------------------------------
describe('Error surfaces (ScriptControl member vs thrown host Error)', () => {
  it('stamps the localized source on a runtime error', () => {
    const engine = new VbsEngine();
    engine.executeStatement('x = 1/0');
    expect(engine.error!.source).toBe('Microsoft VBScript runtime error');
    expect(engine.error!.number).toBe(11);
  });

  it('stamps the compile source on a parse error', () => {
    const engine = new VbsEngine();
    engine.executeStatement('x = )');
    expect(engine.error).not.toBeNull();
  });

  it('surfaces Err.Raise help file and context back to the script', () => {
    const engine = new VbsEngine();
    engine.executeStatement('On Error Resume Next');
    engine.executeStatement('Err.Raise 5, "Src", "Desc", "x.hlp", 3');
    engine.executeStatement('a = Err.HelpFile');
    engine.executeStatement('b = Err.HelpContext');
    expect(engine._getVariable('a').value).toBe('x.hlp');
    expect(engine._getVariable('b').value).toBe(3);
  });

  it('throws an HRESULT-bearing native Error for division by zero', () => {
    const engine = new VbsEngine();
    let caught: Error & { number: number; description: string } | null = null;
    try {
      engine.executeStatementThrows('x = 1/0');
    } catch (e) {
      caught = e as Error & { number: number; description: string };
    }
    expect(caught).not.toBeNull();
    expect(caught).toBeInstanceOf(Error);
    expect(caught!.name).toBe('Error');
    expect(caught!.number).toBe(-2146828277); // 0x800A000B
    expect(caught!.message).toBe('Division by zero');
    expect(caught!.message).toBe(caught!.description);
  });

  it('still populates engine.error even when throwing', () => {
    const engine = new VbsEngine();
    try {
      engine.executeStatementThrows('x = CInt("abc")');
    } catch {
      /* expected */
    }
    expect(engine.error!.number).toBe(13);
    expect(engine.error!.source).toBe('Microsoft VBScript runtime error');
  });
});

// ---------------------------------------------------------------------------
// A host member that answers Nothing
// ---------------------------------------------------------------------------
// A host has no other way to say "no object": null already means Null. The expected
// answers were read off vbscript.dll through a COM property that holds Nothing
// (Scripting.Dictionary.Item after d.Add "k", Nothing).
describe('A host member that answers Nothing', () => {
  const NOTHING = Symbol.for('Nothing');

  function engineWith(host: object): VbsEngine {
    const engine = new VbsEngine();
    engine.addObject('host', host, true);
    return engine;
  }

  it.each<[string, string, unknown]>([
    ['TypeName', 'r = TypeName(host.Image)', 'Nothing'],
    ['IsObject', 'r = IsObject(host.Image)', true],
    ['Is Nothing', 'r = host.Image Is Nothing', true],
    ['VarType', 'r = VarType(host.Image)', 9],
    ['If ... Is Nothing', 'If host.Image Is Nothing Then\n r = "taken"\nElse\n r = "not taken"\nEnd If', 'taken'],
    ['Set', 'Set x = host.Image\nr = x Is Nothing', true],
  ])('reads a property as Nothing: %s', (_label, code, expected) => {
    const engine = engineWith({ get Image() { return NOTHING; } });
    engine.executeStatement(code);
    expect(engine.error).toBeNull();
    expect(engine.eval('r')).toBe(expected);
  });

  it('reads a method result as Nothing', () => {
    const engine = engineWith({ Find: () => NOTHING });
    engine.executeStatement('Set x = host.Find("k")\nr = x Is Nothing');
    expect(engine.error).toBeNull();
    expect(engine.eval('r')).toBe(true);
  });

  it('GUARD keeps null meaning Null', () => {
    const engine = engineWith({ get Image() { return null; } });
    engine.executeStatement('r = TypeName(host.Image)');
    expect(engine.eval('r')).toBe('Null');
  });
});

// ---------------------------------------------------------------------------
// TypeName of Nothing
// ---------------------------------------------------------------------------
// vbscript.dll names an object reference that holds nothing "Nothing", not "Object";
// VarType stays 9 either way.
describe('TypeName of Nothing', () => {
  it.each<[string, string]>([
    ['a variable set to Nothing', 'Set x = Nothing\nr = TypeName(x)'],
    ['the literal', 'r = TypeName(Nothing)'],
  ])('names %s "Nothing"', (_label, code) => {
    const engine = new VbsEngine();
    engine.executeStatement(code);
    expect(engine.error).toBeNull();
    expect(engine.eval('r')).toBe('Nothing');
  });

  it('GUARD still names a live object by its class', () => {
    const engine = new VbsEngine();
    engine.executeStatement('Class C\nEnd Class\nSet x = New C\nr = TypeName(x)');
    expect(engine.eval('r')).toBe('C');
  });
});
