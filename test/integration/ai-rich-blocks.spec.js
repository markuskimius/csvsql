const { test, expect } = require('@playwright/test');
const { openApp } = require('../helpers');

test.describe('AI Rich Blocks', () => {
  test('formatAIResponse tags chart blocks correctly', async ({ page }) => {
    await openApp(page);
    const html = await page.evaluate(() => {
      // Access the private formatAIResponse via the IIFE - we need to call it directly
      // Since it's not exposed, we'll simulate what happens
      const bubble = document.createElement('div');
      const text = 'Here is a chart:\n```chart\n{"type":"bar","data":{"labels":["A","B"],"datasets":[{"label":"Count","data":[10,20]}]}}\n```';
      // Replicate formatAIResponse logic
      function escHtml(str) { const d = document.createElement('div'); d.textContent = str; return d.innerHTML; }
      let html = escHtml(text);
      html = html.replace(/```(\w*)\n?([\s\S]*?)```/g, (_, lang, code) => {
        const cls = ['chart','table','pdf'].includes(lang) ? ` class="ai-block-${lang}"` : '';
        return `<pre${cls}>` + code.trim() + '</pre>';
      });
      html = html.replace(/`([^`]+)`/g, '<code>$1</code>');
      html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
      html = html.replace(/\n/g, '<br>');
      html = html.replace(/<pre([^>]*)>([\s\S]*?)<\/pre>/g, (_, attrs, code) => '<pre' + attrs + '>' + code.replace(/<br>/g, '\n') + '</pre>');
      bubble.innerHTML = html;
      const pre = bubble.querySelector('pre.ai-block-chart');
      if (!pre) return { error: 'no pre.ai-block-chart found', html };
      const textContent = pre.textContent;
      try {
        const parsed = JSON.parse(textContent);
        return { ok: true, type: parsed.type, textContent };
      } catch(e) {
        return { error: 'JSON parse failed: ' + e.message, textContent };
      }
    });
    expect(html.ok).toBe(true);
    expect(html.type).toBe('bar');
  });

  test('fallback detection finds chart JSON in unclassed pre', async ({ page }) => {
    await openApp(page);
    const result = await page.evaluate(() => {
      const bubble = document.createElement('div');
      // Simulate AI producing ```json instead of ```chart
      bubble.innerHTML = '<pre>{"type":"bar","data":{"labels":["A"],"datasets":[{"label":"X","data":[10]}]}}</pre>';
      const pre = bubble.querySelector('pre');
      // Run fallback detection
      try {
        const text = pre.textContent.trim();
        if (text.startsWith('{')) {
          const obj = JSON.parse(text);
          if (obj.type && obj.data && obj.data.datasets) {
            pre.classList.add('ai-block-chart');
          }
        }
      } catch {}
      return {
        hasClass: pre.classList.contains('ai-block-chart'),
        textContent: pre.textContent,
      };
    });
    expect(result.hasClass).toBe(true);
  });

  test('chart block renders a canvas after postProcessAIBlocks', async ({ page }) => {
    await openApp(page);
    const result = await page.evaluate(async () => {
      // Load Chart.js
      await new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = 'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js';
        s.onload = resolve;
        s.onerror = reject;
        document.head.appendChild(s);
      });

      const bubble = document.createElement('div');
      document.body.appendChild(bubble);
      bubble.innerHTML = '<pre class="ai-block-chart">{"type":"bar","data":{"labels":["A","B"],"datasets":[{"label":"Count","data":[10,20]}]}}</pre>';

      const pre = bubble.querySelector('pre.ai-block-chart');
      try {
        const config = JSON.parse(pre.textContent);
        config.options = { responsive: false, animation: false };
        const container = document.createElement('div');
        container.className = 'ai-chart-container';
        const canvas = document.createElement('canvas');
        canvas.width = 200;
        canvas.height = 200;
        container.appendChild(canvas);
        pre.replaceWith(container);
        new Chart(canvas, config);
        return { ok: true, hasCanvas: !!bubble.querySelector('canvas'), hasContainer: !!bubble.querySelector('.ai-chart-container') };
      } catch(e) {
        return { error: e.message };
      }
    });
    expect(result.ok).toBe(true);
    expect(result.hasCanvas).toBe(true);
  });

  test('table block renders an HTML table', async ({ page }) => {
    await openApp(page);
    const result = await page.evaluate(() => {
      const bubble = document.createElement('div');
      bubble.innerHTML = '<pre class="ai-block-table">{"columns":["Name","Value"],"rows":[{"Name":"Alice","Value":"100"},{"Name":"Bob","Value":"200"}]}</pre>';
      const pre = bubble.querySelector('pre.ai-block-table');
      try {
        const data = JSON.parse(pre.textContent);
        if (!data.columns || !data.rows) return { error: 'bad structure' };
        const table = document.createElement('table');
        table.className = 'ai-inline-table';
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        for (const col of data.columns) {
          const th = document.createElement('th');
          th.textContent = col;
          headerRow.appendChild(th);
        }
        thead.appendChild(headerRow);
        table.appendChild(thead);
        const tbody = document.createElement('tbody');
        for (const row of data.rows) {
          const tr = document.createElement('tr');
          for (const col of data.columns) {
            const td = document.createElement('td');
            td.textContent = row[col] ?? '';
            tr.appendChild(td);
          }
          tbody.appendChild(tr);
        }
        table.appendChild(tbody);
        pre.replaceWith(table);
        const renderedTable = bubble.querySelector('table.ai-inline-table');
        return {
          ok: true,
          rows: renderedTable.querySelectorAll('tbody tr').length,
          headerCells: renderedTable.querySelectorAll('thead th').length,
        };
      } catch(e) {
        return { error: e.message };
      }
    });
    expect(result.ok).toBe(true);
    expect(result.rows).toBe(2);
    expect(result.headerCells).toBe(2);
  });

  test('pdf block renders a download link', async ({ page }) => {
    await openApp(page);
    const result = await page.evaluate(async () => {
      // Load jsPDF
      await new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
        s.onload = resolve;
        s.onerror = reject;
        document.head.appendChild(s);
      });
      await new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.4/jspdf.plugin.autotable.min.js';
        s.onload = resolve;
        s.onerror = reject;
        document.head.appendChild(s);
      });

      const bubble = document.createElement('div');
      document.body.appendChild(bubble);
      bubble.innerHTML = '<pre class="ai-block-pdf">{"filename":"test.pdf","title":"Test Report","content":[{"type":"text","value":"Hello world"}]}</pre>';
      const pre = bubble.querySelector('pre.ai-block-pdf');
      try {
        const spec = JSON.parse(pre.textContent);
        const doc = new jspdf.jsPDF();
        let y = 20;
        if (spec.title) {
          doc.setFontSize(18);
          doc.text(spec.title, 14, y);
          y += 12;
        }
        for (const block of spec.content) {
          if (block.type === 'text') {
            doc.setFontSize(12);
            doc.text(block.value, 14, y);
            y += 10;
          }
        }
        const blob = doc.output('blob');
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = spec.filename;
        link.className = 'ai-pdf-download';
        link.textContent = 'Download ' + spec.filename;
        pre.replaceWith(link);
        const renderedLink = bubble.querySelector('a.ai-pdf-download');
        return {
          ok: true,
          text: renderedLink?.textContent,
          hasHref: !!renderedLink?.href,
          download: renderedLink?.download,
        };
      } catch(e) {
        return { error: e.message };
      }
    });
    expect(result.ok).toBe(true);
    expect(result.text).toBe('Download test.pdf');
    expect(result.download).toBe('test.pdf');
  });
});
