const { test, expect } = require('@playwright/test');
const { openApp } = require('../helpers');

async function tokenize(page, text) {
  return page.evaluate(t => app._test.sqlTokenize(t), text);
}

async function highlight(page, text) {
  return page.evaluate(t => app._test.sqlHighlightHTML(t), text);
}

test.describe('sqlTokenize', () => {
  test('classifies keywords, words, numbers, and operators', async ({ page }) => {
    await openApp(page);
    const tokens = await tokenize(page, 'SELECT name, age FROM people WHERE age > 30');
    const types = tokens.filter(t => t.type !== 'ws').map(t => [t.type, t.text]);
    expect(types).toEqual([
      ['keyword', 'SELECT'],
      ['word', 'name'],
      ['op', ','],
      ['word', 'age'],
      ['keyword', 'FROM'],
      ['word', 'people'],
      ['keyword', 'WHERE'],
      ['word', 'age'],
      ['op', '>'],
      ['number', '30'],
    ]);
  });

  test('keyword detection is case-insensitive', async ({ page }) => {
    await openApp(page);
    const tokens = await tokenize(page, 'select From wHeRe');
    expect(tokens.filter(t => t.type !== 'ws').map(t => t.type)).toEqual([
      'keyword', 'keyword', 'keyword',
    ]);
  });

  test('token offsets reconstruct the input exactly', async ({ page }) => {
    await openApp(page);
    const inputs = [
      "SELECT * FROM [my table] WHERE name LIKE '%a''b%' -- trailing\n/* block */ 3.14 <= x",
      '',
      '   \t\n  ',
      "'unterminated string",
      '[unterminated bracket',
      '/* unterminated comment',
      '--',
      'a.b.c(1,2)',
    ];
    for (const input of inputs) {
      const tokens = await tokenize(page, input);
      expect(tokens.map(t => t.text).join('')).toBe(input);
      let pos = 0;
      for (const t of tokens) {
        expect(t.start).toBe(pos);
        expect(t.end).toBeGreaterThan(t.start);
        expect(t.text).toBe(input.slice(t.start, t.end));
        pos = t.end;
      }
      expect(pos).toBe(input.length);
    }
  });

  test('groups consecutive whitespace into one token', async ({ page }) => {
    await openApp(page);
    const tokens = await tokenize(page, 'a  \t\n  b');
    expect(tokens).toEqual([
      { type: 'word', start: 0, end: 1, text: 'a' },
      { type: 'ws', start: 1, end: 7, text: '  \t\n  ' },
      { type: 'word', start: 7, end: 8, text: 'b' },
    ]);
  });

  test('single-quoted strings with escaped quotes', async ({ page }) => {
    await openApp(page);
    const tokens = await tokenize(page, "x = 'it''s' AND y");
    const str = tokens.find(t => t.type === 'string');
    expect(str.text).toBe("'it''s'");
  });

  test('unterminated string runs to end of input', async ({ page }) => {
    await openApp(page);
    const tokens = await tokenize(page, "WHERE name = 'abc");
    const str = tokens.find(t => t.type === 'string');
    expect(str.text).toBe("'abc");
    expect(str.end).toBe(17);
  });

  test('line comments end at newline, block comments at */', async ({ page }) => {
    await openApp(page);
    const tokens = await tokenize(page, '-- one\nSELECT /* two */ 1');
    const comments = tokens.filter(t => t.type === 'comment').map(t => t.text);
    expect(comments).toEqual(['-- one', '/* two */']);
  });

  test('unterminated block comment runs to end of input', async ({ page }) => {
    await openApp(page);
    const tokens = await tokenize(page, 'SELECT /* oops');
    expect(tokens[tokens.length - 1]).toMatchObject({ type: 'comment', text: '/* oops' });
  });

  test('bracket-quoted identifiers, including spaces and unterminated', async ({ page }) => {
    await openApp(page);
    let tokens = await tokenize(page, 'SELECT [my col] FROM [my table (1)]');
    expect(tokens.filter(t => t.type === 'bracket').map(t => t.text))
      .toEqual(['[my col]', '[my table (1)]']);
    tokens = await tokenize(page, 'FROM [oops');
    expect(tokens[tokens.length - 1]).toMatchObject({ type: 'bracket', text: '[oops' });
  });

  test('numbers with decimals and leading dot', async ({ page }) => {
    await openApp(page);
    const tokens = await tokenize(page, 'x > 3.14 AND y < .5 OR z = 42');
    expect(tokens.filter(t => t.type === 'number').map(t => t.text))
      .toEqual(['3.14', '.5', '42']);
  });

  test('dot after identifier is an operator, not a number', async ({ page }) => {
    await openApp(page);
    const tokens = await tokenize(page, 'people.name');
    expect(tokens.map(t => [t.type, t.text])).toEqual([
      ['word', 'people'],
      ['op', '.'],
      ['word', 'name'],
    ]);
  });

  test('minus not followed by minus is an operator', async ({ page }) => {
    await openApp(page);
    const tokens = await tokenize(page, 'a - 1');
    expect(tokens.filter(t => t.type === 'op').map(t => t.text)).toEqual(['-']);
  });
});

test.describe('sqlHighlightHTML', () => {
  test('renders keyword, string, number, bracket, and comment spans', async ({ page }) => {
    await openApp(page);
    expect(await highlight(page, "SELECT x FROM [my table] WHERE n = 'a' -- c"))
      .toBe('<span class="sql-kw">SELECT</span> x <span class="sql-kw">FROM</span> ' +
        '<span class="sql-brk">[my table]</span> <span class="sql-kw">WHERE</span> n = ' +
        '<span class="sql-str">\'a\'</span> <span class="sql-cmt">-- c</span>');
    expect(await highlight(page, 'x > 3.14'))
      .toBe('x &gt; <span class="sql-num">3.14</span>');
  });

  test('escapes HTML special characters in all token types', async ({ page }) => {
    await openApp(page);
    expect(await highlight(page, "'<b>' <> [a<b]"))
      .toBe('<span class="sql-str">\'&lt;b&gt;\'</span> &lt;&gt; ' +
        '<span class="sql-brk">[a&lt;b]</span>');
  });

  test('preserves whitespace verbatim', async ({ page }) => {
    await openApp(page);
    expect(await highlight(page, 'a  \n\tb')).toBe('a  \n\tb');
  });

  test('empty input renders empty string', async ({ page }) => {
    await openApp(page);
    expect(await highlight(page, '')).toBe('');
  });
});
