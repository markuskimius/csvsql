const { test, expect } = require('@playwright/test');
const { openApp } = require('../helpers');

test.describe('Expression language', () => {
  let page;

  test.beforeEach(async ({ page: p }) => {
    page = p;
    await openApp(page);
  });

  async function evalExpr(expr, env = {}) {
    return page.evaluate(({ expr, env }) => {
      const ast = app._test.exprCompile(expr);
      const result = app._test.exprEval(ast, {
        value: env.value ?? '',
        column: env.column ?? 'col',
        row: env.row ?? {},
        table: env.table ?? 'tbl',
      });
      return result;
    }, { expr, env });
  }

  async function evalToString(expr, env = {}) {
    return page.evaluate(({ expr, env }) => {
      const ast = app._test.exprCompile(expr);
      return app._test.exprEvalToString(ast, {
        value: env.value ?? '',
        column: env.column ?? 'col',
        row: env.row ?? {},
        table: env.table ?? 'tbl',
      });
    }, { expr, env });
  }

  test('string literals', async () => {
    expect(await evalExpr("'hello'")).toBe('hello');
    expect(await evalExpr('"world"')).toBe('world');
  });

  test('number literals', async () => {
    expect(await evalExpr('42')).toBe(42);
    expect(await evalExpr('3.14')).toBe(3.14);
  });

  test('boolean and null literals', async () => {
    expect(await evalExpr('true')).toBe(true);
    expect(await evalExpr('false')).toBe(false);
    expect(await evalExpr('null')).toBeNull();
  });

  test('variable access', async () => {
    expect(await evalExpr('value', { value: 'hi' })).toBe('hi');
    expect(await evalExpr('column', { column: 'name' })).toBe('name');
    expect(await evalExpr('table', { table: 'users' })).toBe('users');
  });

  test('row property access', async () => {
    expect(await evalExpr('row.name', { row: { name: 'Alice' } })).toBe('Alice');
    expect(await evalExpr('row.missing', { row: {} })).toBeNull();
  });

  test('blocked prototype access', async () => {
    expect(await evalExpr('row.constructor', { row: {} })).toBeNull();
    expect(await evalExpr('row.__proto__', { row: {} })).toBeNull();
  });

  test('arithmetic operators', async () => {
    expect(await evalExpr('3 + 4')).toBe(7);
    expect(await evalExpr('10 - 3')).toBe(7);
    expect(await evalExpr('2 * 5')).toBe(10);
    expect(await evalExpr('15 / 3')).toBe(5);
    expect(await evalExpr('7 % 3')).toBe(1);
  });

  test('division by zero returns null', async () => {
    expect(await evalExpr('5 / 0')).toBeNull();
    expect(await evalExpr('5 % 0')).toBeNull();
  });

  test('string concatenation with +', async () => {
    expect(await evalExpr("'hello' + ' ' + 'world'")).toBe('hello world');
    expect(await evalExpr("'$' + 42")).toBe('$42');
  });

  test('comparison operators', async () => {
    expect(await evalExpr('3 < 5')).toBe(true);
    expect(await evalExpr('5 > 3')).toBe(true);
    expect(await evalExpr('3 <= 3')).toBe(true);
    expect(await evalExpr('3 >= 4')).toBe(false);
  });

  test('equality operators', async () => {
    expect(await evalExpr("'a' == 'a'")).toBe(true);
    expect(await evalExpr("'a' != 'b'")).toBe(true);
    expect(await evalExpr("'a' == 'b'")).toBe(false);
  });

  test('logical operators with short-circuit', async () => {
    expect(await evalExpr("'hello' || 'default'")).toBe('hello');
    expect(await evalExpr("'' || 'default'")).toBe('default');
    expect(await evalExpr("'hello' && 'world'")).toBe('world');
    expect(await evalExpr("'' && 'world'")).toBe('');
  });

  test('ternary operator', async () => {
    expect(await evalExpr("true ? 'yes' : 'no'")).toBe('yes');
    expect(await evalExpr("false ? 'yes' : 'no'")).toBe('no');
  });

  test('unary operators', async () => {
    expect(await evalExpr('!true')).toBe(false);
    expect(await evalExpr('!false')).toBe(true);
    expect(await evalExpr('-5')).toBe(-5);
  });

  test('operator precedence', async () => {
    expect(await evalExpr('2 + 3 * 4')).toBe(14);
    expect(await evalExpr('(2 + 3) * 4')).toBe(20);
  });

  test('string functions', async () => {
    expect(await evalExpr("upper('hello')")).toBe('HELLO');
    expect(await evalExpr("lower('HELLO')")).toBe('hello');
    expect(await evalExpr("trim('  hi  ')")).toBe('hi');
    expect(await evalExpr("len('abc')")).toBe(3);
    expect(await evalExpr("substr('hello', 1, 3)")).toBe('ell');
    expect(await evalExpr("replace('a-b-c', '-', '/')")).toBe('a/b-c');
    expect(await evalExpr("replaceAll('a-b-c', '-', '/')")).toBe('a/b/c');
    expect(await evalExpr("startsWith('hello', 'he')")).toBe(true);
    expect(await evalExpr("endsWith('hello', 'lo')")).toBe(true);
    expect(await evalExpr("contains('hello', 'ell')")).toBe(true);
    expect(await evalExpr("padLeft('42', 5, '0')")).toBe('00042');
    expect(await evalExpr("padRight('hi', 5, '.')")).toBe('hi...');
    expect(await evalExpr("concat('a', 'b', 'c')")).toBe('abc');
    expect(await evalExpr("repeat('*', 3)")).toBe('***');
  });

  test('number functions', async () => {
    expect(await evalExpr("num('42.5')")).toBe(42.5);
    expect(await evalExpr("num('abc')")).toBeNull();
    expect(await evalToString("fixed(3.1, 2)")).toBe('3.10');
    expect(await evalExpr("round(3.456, 2)")).toBe(3.46);
    expect(await evalExpr("floor(3.9)")).toBe(3);
    expect(await evalExpr("ceil(3.1)")).toBe(4);
    expect(await evalExpr("abs(-5)")).toBe(5);
    expect(await evalExpr("min(3, 7)")).toBe(3);
    expect(await evalExpr("max(3, 7)")).toBe(7);
  });

  test('date function', async () => {
    expect(await evalExpr("date('2024-01-15', 'iso')")).toBe('2024-01-15');
    expect(await evalExpr("date('not-a-date', 'locale')")).toBe('not-a-date');
    expect(await evalExpr("date('', 'locale')")).toBe('');
    expect(await evalExpr("date('2024-03-05T12:00:00', 'YYYY/MM/DD')")).toBe('2024/03/05');
  });

  test('date function with decimal seconds timestamp', async () => {
    const result = await evalToString("date('1780862826.123456789', 'full')");
    expect(result).toMatch(/^2026-06-0[78] \d{2}:\d{2}:\d{2}\.123456789$/);
  });

  test('date function with decimal seconds pads fractional to 9 digits', async () => {
    const result = await evalToString("date('1780862826.12', 'full')");
    expect(result).toMatch(/^2026-06-0[78] \d{2}:\d{2}:\d{2}\.120000000$/);
  });

  test('date function with decimal seconds truncates beyond 9 digits', async () => {
    const result = await evalToString("date('1780862826.1234567891111', 'full')");
    expect(result).toMatch(/^2026-06-0[78] \d{2}:\d{2}:\d{2}\.123456789$/);
  });

  test('date function with decimal seconds and non-full format', async () => {
    const result = await evalExpr("date('1780862826.5', 'iso')");
    expect(result).toBe('2026-06-07');
  });

  test('date function with integer seconds timestamp string', async () => {
    const result = await evalExpr("date('1704067200', 'iso')");
    expect(result).toBe('2024-01-01');
  });

  test('if function', async () => {
    expect(await evalExpr("if(true, 'yes', 'no')")).toBe('yes');
    expect(await evalExpr("if(false, 'yes', 'no')")).toBe('no');
  });

  test('choose function with match', async () => {
    expect(await evalExpr("choose('A', 'A', 'Active', 'I', 'Inactive')")).toBe('Active');
    expect(await evalExpr("choose('I', 'A', 'Active', 'I', 'Inactive')")).toBe('Inactive');
  });

  test('choose function with no match returns value', async () => {
    expect(await evalExpr("choose('X', 'A', 'Active', 'I', 'Inactive')")).toBe('X');
  });

  test('choose function with default', async () => {
    expect(await evalExpr("choose('X', 'A', 'Active', 'I', 'Inactive', 'Unknown')")).toBe('Unknown');
  });

  test('coalesce function', async () => {
    expect(await evalExpr("coalesce('', null, 'fallback')")).toBe('fallback');
    expect(await evalExpr("coalesce('first', 'second')")).toBe('first');
  });

  test('isEmpty and isNum', async () => {
    expect(await evalExpr("isEmpty('')")).toBe(true);
    expect(await evalExpr("isEmpty('x')")).toBe(false);
    expect(await evalExpr("isEmpty(null)")).toBe(true);
    expect(await evalExpr("isNum('42')")).toBe(true);
    expect(await evalExpr("isNum('abc')")).toBe(false);
    expect(await evalExpr("isNum('')")).toBe(false);
  });

  test('nested function calls', async () => {
    expect(await evalExpr("upper(trim('  hello  '))")).toBe('HELLO');
  });

  test('complex real-world expression', async () => {
    const result = await evalToString(
      "isEmpty(value) ? '' : date(value, 'locale')",
      { value: '' }
    );
    expect(result).toBe('');
  });

  test('evalToString coerces null to empty string', async () => {
    expect(await evalToString('null')).toBe('');
  });

  test('evalToString coerces booleans', async () => {
    expect(await evalToString('true')).toBe('true');
    expect(await evalToString('false')).toBe('false');
  });

  test('parse error on invalid syntax', async () => {
    const error = await page.evaluate(() => {
      try { app._test.exprCompile('upper('); return null; }
      catch (e) { return e.message; }
    });
    expect(error).toBeTruthy();
  });

  test('runtime error falls back to raw value', async () => {
    const result = await page.evaluate(() => {
      const ast = app._test.exprCompile("unknownFunc(value)");
      return app._test.exprEvalToString(ast, { value: 'raw', column: 'c', row: {}, table: 't' });
    });
    expect(result).toBe('raw');
  });
});
