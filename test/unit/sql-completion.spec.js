const { test, expect } = require('@playwright/test');
const { openApp } = require('../helpers');

const SCHEMA = {
  people: ['name', 'age', 'city'],
  orders: ['id', 'person_id', 'total'],
  'my table (1)': ['col a', 'order'],
};

// '|' marks the caret and is removed before calling. Avoid '||' in inputs.
async function ctx(page, textWithCaret, opts) {
  const caret = textWithCaret.indexOf('|');
  const text = textWithCaret.replace('|', '');
  return page.evaluate(
    ({ text, caret, opts }) => app._test.sqlCompletionContext(text, caret, opts),
    { text, caret, opts: Object.assign({ schema: SCHEMA }, opts) }
  );
}

const labels = (r, kind) => r.items.filter(i => !kind || i.kind === kind).map(i => i.label);
const kinds = r => [...new Set(r.items.map(i => i.kind))];

test.describe('sqlCompletionContext — table position', () => {
  test('after FROM suggests tables only, sorted, quoted as needed', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, 'SELECT * FROM |');
    expect(kinds(r)).toEqual(['table']);
    expect(labels(r)).toEqual(['my table (1)', 'orders', 'people']);
    expect(r.items[0].insert).toBe('[my table (1)]');
    expect(r.items[1].insert).toBe('orders');
  });

  test('prefix after FROM filters tables and sets replace range', async ({ page }) => {
    await openApp(page);
    const text = 'SELECT * FROM peo';
    const r = await ctx(page, 'SELECT * FROM peo|');
    expect(labels(r)).toEqual(['people']);
    expect(r.prefix).toBe('peo');
    expect(r.start).toBe(text.indexOf('peo'));
    expect(r.end).toBe(text.length);
  });

  test('caret mid-word replaces the whole word but matches the left part', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, 'SELECT * FROM pe|ox');
    expect(r.prefix).toBe('pe');
    expect(r.end - r.start).toBe('peox'.length);
    expect(labels(r)).toEqual(['people']);
  });

  test('JOIN, UPDATE, and TABLE also give table position', async ({ page }) => {
    await openApp(page);
    for (const sql of ['SELECT * FROM people JOIN or|', 'UPDATE or|', 'DROP TABLE or|']) {
      const r = await ctx(page, sql);
      expect(kinds(r)).toEqual(['table']);
      expect(labels(r)).toEqual(['orders']);
    }
  });

  test('comma continues a FROM list', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, 'SELECT * FROM people, or|');
    expect(kinds(r)).toEqual(['table']);
    expect(labels(r)).toEqual(['orders']);
  });

  test('comma in a SELECT list is not table position', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, 'SELECT name, ag| FROM people');
    expect(labels(r, 'column')).toEqual(['age']);
  });

  test('INSERT INTO suggests tables; SELECT INTO suggests nothing (new name)', async ({ page }) => {
    await openApp(page);
    const ins = await ctx(page, 'INSERT INTO pe|');
    expect(labels(ins)).toEqual(['people']);
    const sel = await ctx(page, 'SELECT * INTO | FROM people');
    expect(sel.items).toEqual([]);
  });
});

test.describe('sqlCompletionContext — qualified columns', () => {
  test('table.prefix suggests that table\'s columns', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, 'SELECT people.| FROM people');
    expect(labels(r)).toEqual(['name', 'age', 'city']);
    expect(kinds(r)).toEqual(['column']);
  });

  test('table resolution is case-insensitive', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, 'SELECT PEOPLE.na| FROM people');
    expect(labels(r)).toEqual(['name']);
  });

  test('alias.prefix resolves through FROM/JOIN aliases, with and without AS', async ({ page }) => {
    await openApp(page);
    let r = await ctx(page, 'SELECT p.| FROM people p');
    expect(labels(r)).toEqual(['name', 'age', 'city']);
    r = await ctx(page, 'SELECT o.to| FROM people AS p JOIN orders AS o ON p.name = o.id');
    expect(labels(r)).toEqual(['total']);
  });

  test('bracket-quoted table qualifier resolves and quotes inserted columns', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, 'SELECT [my table (1)].| FROM [my table (1)]');
    expect(labels(r)).toEqual(['col a', 'order']);
    expect(r.items[0].insert).toBe('[col a]');
    expect(r.items[1].insert).toBe('[order]'); // keyword collision gets quoted too
  });

  test('unknown qualifier suggests nothing rather than something wrong', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, 'SELECT nosuch.| FROM people');
    expect(r.items).toEqual([]);
  });
});

test.describe('sqlCompletionContext — suppression', () => {
  test('returns null inside strings, comments, and numbers', async ({ page }) => {
    await openApp(page);
    expect(await ctx(page, "SELECT * FROM people WHERE name = 'ab|c'")).toBeNull();
    expect(await ctx(page, 'SELECT 1 -- from pe|ople')).toBeNull();
    expect(await ctx(page, 'SELECT 1 /* fro|m */')).toBeNull();
    expect(await ctx(page, 'SELECT 12|34')).toBeNull();
  });
});

test.describe('sqlCompletionContext — default context', () => {
  test('columns of mentioned tables come first, then tables, then keywords', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, 'SELECT | FROM people');
    const ks = r.items.map(i => i.kind);
    expect(labels(r, 'column')).toEqual(['name', 'age', 'city']);
    expect(ks.indexOf('table')).toBeGreaterThan(ks.lastIndexOf('column'));
    expect(ks.indexOf('keyword')).toBeGreaterThan(ks.lastIndexOf('table'));
  });

  test('typing a keyword prefix offers the keyword', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, 'sel|');
    expect(labels(r, 'keyword')).toEqual(['SELECT']);
  });

  test('statements are isolated at semicolons', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, 'SELECT name FROM people; SELECT |');
    expect(labels(r, 'column')).toEqual([]); // people mentioned only in the first statement
    const r2 = await ctx(page, 'SELECT name FROM people; SELECT ag| FROM people');
    expect(labels(r2, 'column')).toEqual(['age']);
  });

  test('empty input at caret 0 returns tables and keywords without throwing', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, '|');
    expect(r.items.length).toBeGreaterThan(0);
    expect(labels(r, 'table')).toEqual(['my table (1)', 'orders', 'people']);
  });

  test('completing inside an open bracket keeps bracket quoting', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, 'SELECT * FROM [my|');
    expect(r.prefix).toBe('my');
    expect(labels(r)).toEqual(['my table (1)']);
    expect(r.items[0].insert).toBe('[my table (1)]');
  });
});

test.describe('sqlCompletionContext — filter mode', () => {
  const opts = { filterTable: 'people' };

  test('own columns first, keywords after, no bare table names', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, '|', opts);
    expect(labels(r, 'column')).toEqual(['name', 'age', 'city']);
    expect(labels(r, 'table')).toEqual([]);
    expect(r.items.some(i => i.kind === 'keyword')).toBe(true);
  });

  test('prefix matches columns and keywords together', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, 'na|', opts);
    expect(r.items[0]).toMatchObject({ label: 'name', kind: 'column' });
    expect(labels(r, 'keyword')).toContain('NATURAL');
  });

  test('subquery FROM inside a filter still suggests tables', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, 'age > (SELECT avg(age) FROM |', opts);
    expect(kinds(r)).toEqual(['table']);
    expect(labels(r)).toEqual(['my table (1)', 'orders', 'people']);
  });

  test('unknown filterTable is ignored gracefully', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, 'na|', { filterTable: 'nosuch' });
    expect(r.items.length).toBeGreaterThan(0);
  });
});

test.describe('sqlCompletionContext — keyword context gating', () => {
  test('type names and NOCASE are not offered in expression context', async ({ page }) => {
    await openApp(page);
    expect(labels(await ctx(page, 'SELECT nu| FROM people'), 'keyword')).not.toContain('NUMERIC');
    expect(labels(await ctx(page, 'SELECT in| FROM people'), 'keyword')).not.toContain('INTEGER');
    expect(labels(await ctx(page, 'SELECT no| FROM people'), 'keyword')).not.toContain('NOCASE');
  });

  test('CAST(... AS offers exactly the type names', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, 'SELECT CAST(age AS |) FROM people');
    expect(kinds(r)).toEqual(['keyword']);
    expect(labels(r).sort()).toEqual(['BLOB', 'BOOLEAN', 'INTEGER', 'NUMERIC', 'REAL', 'TEXT']);
    expect(labels(await ctx(page, 'SELECT CAST(age AS nu|) FROM people'))).toEqual(['NUMERIC']);
  });

  test('CAST detection tracks nested parens', async ({ page }) => {
    await openApp(page);
    expect(labels(await ctx(page, 'SELECT CAST(round(age, 2) AS in|'))).toEqual(['INTEGER']);
  });

  test('plain alias AS is not a type position', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, 'SELECT age AS n| FROM people');
    expect(labels(r, 'keyword')).not.toContain('NUMERIC');
    expect(labels(r, 'column')).toEqual(['name']);
  });

  test('COLLATE offers only collation names', async ({ page }) => {
    await openApp(page);
    const r = await ctx(page, 'SELECT name FROM people ORDER BY name COLLATE no|');
    expect(labels(r)).toEqual(['NOCASE']);
    expect(kinds(r)).toEqual(['keyword']);
  });

  test('statement commands only at statement start', async ({ page }) => {
    await openApp(page);
    expect(labels(await ctx(page, 'va|'), 'keyword')).toContain('VACUUM');
    expect(labels(await ctx(page, 'SELECT name FROM people; va|'), 'keyword')).toContain('VACUUM');
    expect(labels(await ctx(page, 'SELECT * FROM people WHERE va|'), 'keyword')).not.toContain('VACUUM');
    expect(labels(await ctx(page, 'SELECT * FROM people WHERE pr|'), 'keyword')).not.toContain('PRAGMA');
  });

  test('SELECT mid-statement only inside parens, after compound ops, or after AS', async ({ page }) => {
    await openApp(page);
    expect(labels(await ctx(page, 'SELECT * FROM people WHERE sel|'), 'keyword')).not.toContain('SELECT');
    expect(labels(await ctx(page, 'SELECT * FROM people WHERE age > (sel|'), 'keyword')).toContain('SELECT');
    expect(labels(await ctx(page, 'SELECT name FROM people UNION sel|'), 'keyword')).toContain('SELECT');
    expect(labels(await ctx(page, 'SELECT name FROM people UNION ALL sel|'), 'keyword')).toContain('SELECT');
    expect(labels(await ctx(page, 'EXPLAIN sel|'), 'keyword')).toContain('SELECT');
    expect(labels(await ctx(page, 'CREATE TABLE t2 AS sel|'), 'keyword')).toContain('SELECT');
  });

  test('VALUES stays available mid-statement (INSERT INTO ... VALUES)', async ({ page }) => {
    await openApp(page);
    expect(labels(await ctx(page, 'INSERT INTO people val|'), 'keyword')).toContain('VALUES');
  });

  test('filter inputs never see statement commands', async ({ page }) => {
    await openApp(page);
    const opts = { filterTable: 'people' };
    expect(labels(await ctx(page, 'va|', opts), 'keyword')).not.toContain('VACUUM');
    expect(labels(await ctx(page, 'sel|', opts), 'keyword')).not.toContain('SELECT');
    expect(labels(await ctx(page, 'nu|', opts), 'keyword')).not.toContain('NUMERIC');
    expect(labels(await ctx(page, 'age > (SELECT avg(age) FROM people) AND sel|', opts), 'keyword'))
      .not.toContain('SELECT');
    expect(labels(await ctx(page, 'age > (sel|', opts), 'keyword')).toContain('SELECT');
  });
});

test.describe('sqlCompletionContext — robustness', () => {
  test('never throws: fuzzed inputs at every caret position', async ({ page }) => {
    await openApp(page);
    const failures = await page.evaluate(schema => {
      const fixed = [
        'SELECT * FROM people;;;', 'FROM', '.', '..', 'a.', '.a', '[', ']', '[]',
        'FROM [', "'", '--', '/*', ';', ',,,', 'FROM ,', 'SELECT (', 'INTO',
        'FROM people people people', 'SELECT a.b.c.d FROM x',
      ];
      let seed = 7;
      const rnd = () => (seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
      const alphabet = " '[].,;-/*\n\tSELCTfromjoinwhere_019()x";
      const corpus = fixed.slice();
      for (let n = 0; n < 300; n++) {
        let s = '';
        const L = Math.floor(rnd() * 30);
        for (let k = 0; k < L; k++) s += alphabet[Math.floor(rnd() * alphabet.length)];
        corpus.push(s);
      }
      const failures = [];
      for (const text of corpus) {
        for (let caret = 0; caret <= text.length; caret++) {
          try {
            const r = app._test.sqlCompletionContext(text, caret, { schema });
            if (r !== null && !Array.isArray(r.items)) {
              failures.push({ text, caret, problem: 'bad shape' });
            }
          } catch (e) {
            failures.push({ text, caret, problem: String(e) });
          }
        }
      }
      return failures;
    }, SCHEMA);
    expect(failures).toEqual([]);
  });
});
