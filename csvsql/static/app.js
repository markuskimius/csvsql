// ============================================================
// CSVSQL - CSV Database Application
// ============================================================

const app = (() => {
  let windows = [];
  let nextWinId = 1;
  let nextZIndex = 100;
  let activeWinId = null;
  let tables = {};  // tableName -> { columns, rows, filename, modified }
  let db = null;    // sql.js Database instance

  // Virtual scrolling constants
  const ROW_HEIGHT = 26;
  const OVERSCAN = 10;

  // Debounced sync timers
  const syncTimers = {};

  // Sort optimization
  const collator = new Intl.Collator(undefined, { sensitivity: 'base' });

  // ---- Init ----
  async function init() {
    const SQL = await initSqlJs({
      locateFile: file => `https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.10.3/${file}`
    });
    db = new SQL.Database();
    setupConsoleResize();
    setupFileInput();
    setupDragAndDrop();
    setupKeyboard();
    setupMenuClose();
    fixShortcutLabels();
  }

  function fixShortcutLabels() {
    if (navigator.platform.includes('Mac') || navigator.userAgent.includes('Mac')) {
      document.querySelectorAll('.shortcut').forEach(el => {
        el.textContent = el.textContent.replace('Ctrl+', '\u2318').replace('Shift+', '\u21E7');
      });
    }
  }

  // ---- File Menu ----
  function setupFileInput() {
    document.getElementById('file-input').addEventListener('change', (e) => {
      for (const file of e.target.files) {
        openFileByType(file);
      }
      e.target.value = '';
    });
  }

  function setupDragAndDrop() {
    const overlay = document.getElementById('drop-overlay');
    let dragCounter = 0;

    document.addEventListener('dragenter', (e) => {
      e.preventDefault();
      dragCounter++;
      overlay.classList.add('visible');
    });

    document.addEventListener('dragleave', (e) => {
      e.preventDefault();
      dragCounter--;
      if (dragCounter <= 0) {
        dragCounter = 0;
        overlay.classList.remove('visible');
      }
    });

    document.addEventListener('dragover', (e) => {
      e.preventDefault();
    });

    document.addEventListener('drop', async (e) => {
      e.preventDefault();
      dragCounter = 0;
      overlay.classList.remove('visible');
      closeMenus();
      const entries = [...e.dataTransfer.items].map(item => ({
        file: item.getAsFile(),
        handlePromise: item.getAsFileSystemHandle ? item.getAsFileSystemHandle().catch(() => null) : Promise.resolve(null),
      }));
      for (const entry of entries) {
        const handle = await entry.handlePromise;
        if (entry.file) openFileByType(entry.file, handle);
      }
    });
  }

  async function openFile() {
    if (window.showOpenFilePicker) {
      try {
        const handles = await showOpenFilePicker({
          multiple: true,
          types: [
            { description: 'Data files', accept: { 'text/csv': ['.csv', '.tsv', '.psv', '.txt'] } },
            { description: 'Excel files', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'], 'application/vnd.ms-excel': ['.xls'] } },
            { description: 'Compressed files', accept: { 'application/gzip': ['.gz'], 'application/zip': ['.zip'] } },
          ],
        });
        for (const handle of handles) {
          const file = await handle.getFile();
          openFileByType(file, handle);
        }
      } catch (e) {
        if (e.name !== 'AbortError') setStatus(`Error opening file: ${e.message}`, 'error');
      }
    } else {
      document.getElementById('file-input').click();
    }
  }

  function openURL() {
    showPrompt('Open URL', 'Enter URL (http or https):', '', async (url) => {
      if (!url) return;
      url = url.trim();
      if (/^(ftp|sftp):\/\//i.test(url)) {
        setStatus('FTP/SFTP not supported in browser — use http or https', 'error');
        return;
      }
      if (!/^https?:\/\//i.test(url)) {
        url = 'https://' + url;
      }
      const filename = decodeURIComponent(url.split('/').pop().split('?')[0]) || 'download.csv';
      setStatus(`Fetching ${filename}...`, 'working');
      try {
        const resp = await fetch(url);
        if (!resp.ok) throw new Error(`HTTP ${resp.status} ${resp.statusText}`);
        const blob = await resp.blob();
        const file = new File([blob], filename, { type: blob.type });
        openFileByType(file, null);
      } catch (e) {
        setStatus(`Error fetching URL: ${e.message}`, 'error');
      }
    });
  }

  // Compression extensions and the data file extensions they may wrap
  const COMPRESSION_EXTS = new Set(['gz', 'zip', 'bz2', 'xz', 'rar', '7z', 'zst']);
  const DATA_EXTS = new Set(['csv', 'tsv', 'psv', 'txt', 'xlsx', 'xls']);

  function openFileByType(file, fileHandle) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (COMPRESSION_EXTS.has(ext)) {
      decompressAndOpen(file);
    } else if (ext === 'xlsx' || ext === 'xls') {
      loadExcelFile(file);
    } else {
      loadDelimitedFile(file, fileHandle);
    }
  }

  async function decompressAndOpen(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    setStatus(`Decompressing ${file.name}...`, 'working');
    try {
      if (ext === 'gz') {
        await decompressGzip(file);
      } else if (ext === 'zip') {
        await decompressZip(file);
      } else {
        setStatus(`Unsupported compression format: .${ext} — please decompress the file first and open the decompressed file`, 'error');
      }
    } catch (e) {
      setStatus(`Error decompressing ${file.name}: ${e.message}`, 'error');
    }
  }

  async function decompressGzip(file) {
    const ds = new DecompressionStream('gzip');
    const decompressed = file.stream().pipeThrough(ds);
    const blob = await new Response(decompressed).blob();
    // Inner filename: strip .gz
    const innerName = file.name.replace(/\.gz$/i, '') || 'decompressed.csv';
    const innerFile = new File([blob], innerName, { type: 'application/octet-stream' });
    openFileByType(innerFile, null);
  }

  async function decompressZip(file) {
    const zip = await JSZip.loadAsync(file);
    let found = 0;
    for (const [name, entry] of Object.entries(zip.files)) {
      if (entry.dir) continue;
      const innerExt = name.split('.').pop().toLowerCase();
      if (DATA_EXTS.has(innerExt) || COMPRESSION_EXTS.has(innerExt)) {
        const blob = await entry.async('blob');
        const innerFile = new File([blob], name, { type: 'application/octet-stream' });
        openFileByType(innerFile, null);
        found++;
      }
    }
    if (found === 0) {
      // No recognized files — try to open the first file as CSV
      const firstEntry = Object.values(zip.files).find(e => !e.dir);
      if (firstEntry) {
        const blob = await firstEntry.async('blob');
        const innerFile = new File([blob], firstEntry.name || 'data.csv', { type: 'text/csv' });
        openFileByType(innerFile, null);
      } else {
        setStatus('ZIP archive is empty', 'error');
      }
    }
  }

  function delimiterForExt(filename) {
    const ext = filename.split('.').pop().toLowerCase();
    if (ext === 'tsv') return '\t';
    if (ext === 'psv') return '|';
    return undefined; // let Papa auto-detect (handles csv, txt)
  }

  function loadDelimitedFile(file, fileHandle) {
    setStatus(`Loading ${file.name}...`, 'working');
    const t0 = performance.now();
    const delimiter = delimiterForExt(file.name);
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      dynamicTyping: false,
      delimiter,
      complete(results) {
        const detectedDelimiter = delimiter || results.meta.delimiter || ',';
        const name = sanitizeTableName(file.name.replace(/\.[^.]+$/, ''));
        const uniqueName = getUniqueTableName(name);
        const rawColumns = results.meta.fields || [];
        const columns = sanitizeColumns(rawColumns);
        const rows = results.data.map((row, i) => {
          const r = { _rownum: i + 1 };
          rawColumns.forEach((raw, j) => { r[columns[j]] = row[raw] ?? ''; });
          return r;
        });
        tables[uniqueName] = { columns, rows, filename: file.name, modified: false, fileHandle: fileHandle || null, delimiter: detectedDelimiter };
        registerTable(uniqueName);
        createTableWindow(uniqueName);
        const elapsed = performance.now() - t0;
        setStatus(`Opened ${file.name} (${rows.length} rows) in ${formatElapsed(elapsed)}`, 'success');
      },
      error(err) {
        setStatus(`Error parsing ${file.name}: ${err.message}`, 'error');
      }
    });
  }

  function loadExcelFile(file) {
    setStatus(`Loading ${file.name}...`, 'working');
    const t0 = performance.now();
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target.result, { type: 'array' });
        workbook.SheetNames.forEach(sheetName => {
          const sheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: '' });
          if (jsonData.length === 0) return;
          const baseName = file.name.replace(/\.[^.]+$/, '');
          const label = workbook.SheetNames.length > 1 ? baseName + '_' + sheetName : baseName;
          const name = sanitizeTableName(label);
          const uniqueName = getUniqueTableName(name);
          const rawColumns = Object.keys(jsonData[0]);
          const columns = sanitizeColumns(rawColumns);
          const rows = jsonData.map((row, i) => {
            const r = { _rownum: i + 1 };
            rawColumns.forEach((raw, j) => { r[columns[j]] = row[raw] != null ? String(row[raw]) : ''; });
            return r;
          });
          tables[uniqueName] = { columns, rows, filename: file.name, modified: false };
          registerTable(uniqueName);
          createTableWindow(uniqueName);
        });
        const elapsed = performance.now() - t0;
        setStatus(`Opened ${file.name} (${workbook.SheetNames.length} sheet(s)) in ${formatElapsed(elapsed)}`, 'success');
      } catch (err) {
        setStatus(`Error reading ${file.name}: ${err.message}`, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function sanitizeTableName(name) {
    return name.replace(/[^a-zA-Z0-9_]/g, '_').replace(/^[0-9]/, '_$&');
  }

  function sanitizeColumnName(name) {
    return name.replace(/[^a-zA-Z0-9_]/g, '_').replace(/^[0-9]/, '_$&');
  }

  // Dedup duplicate column names (e.g. two columns both named "Name" → "Name", "Name_2")
  function sanitizeColumns(rawColumns) {
    const seen = {};
    return rawColumns.map(raw => {
      let col = raw;
      if (seen[col]) {
        let n = seen[col] + 1;
        while (seen[col + '_' + n]) n++;
        seen[col] = n;
        col = col + '_' + n;
      }
      seen[col] = (seen[col] || 0) + 1;
      return col;
    });
  }

  function getUniqueTableName(base) {
    let name = base;
    let i = 2;
    while (tables[name]) { name = base + '_' + i; i++; }
    return name;
  }

  function registerTable(tableName) {
    const t = tables[tableName];
    if (!db) return;
    try { db.run(`DROP TABLE IF EXISTS [${tableName}]`); } catch (e) {}
    if (t.columns.length === 0) {
      db.run(`CREATE TABLE [${tableName}] (_empty TEXT)`);
    } else {
      const colDefs = t.columns.map(c => `[${c}] TEXT`).join(', ');
      db.run(`CREATE TABLE [${tableName}] (${colDefs})`);
    }
    if (t.rows.length === 0 || t.columns.length === 0) return;
    const placeholders = t.columns.map(() => '?').join(', ');
    db.run('BEGIN TRANSACTION');
    const stmt = db.prepare(`INSERT INTO [${tableName}] VALUES (${placeholders})`);
    for (const row of t.rows) {
      stmt.run(t.columns.map(c => row[c] ?? ''));
    }
    stmt.free();
    db.run('COMMIT');
  }

  async function saveActiveTable() {
    flushAllSyncs();
    const win = getActiveDataWindow();
    if (!win) return;
    const t = tables[win.tableName];
    if (!t) return;
    if (t.fileHandle) {
      await writeToHandle(win.tableName, t.fileHandle);
    } else {
      const filename = t.filename || win.tableName + '.csv';
      downloadCSV(win.tableName, filename);
    }
  }

  async function saveActiveTableAs() {
    flushAllSyncs();
    const win = getActiveDataWindow();
    if (!win) return;
    const t = tables[win.tableName];
    if (!t) return;
    const filename = t.filename || win.tableName + '.csv';
    if (window.showSaveFilePicker) {
      try {
        const handle = await showSaveFilePicker({
          suggestedName: filename,
          types: [
            { description: 'CSV files', accept: { 'text/csv': ['.csv'] } },
            { description: 'TSV files', accept: { 'text/tab-separated-values': ['.tsv'] } },
            { description: 'PSV files', accept: { 'text/plain': ['.psv'] } },
          ],
        });
        await writeToHandle(win.tableName, handle);
        t.fileHandle = handle;
      } catch (e) {
        if (e.name !== 'AbortError') setStatus(`Error saving: ${e.message}`, 'error');
      }
    } else {
      showPrompt('Save As', 'Filename:', filename, (newName) => {
        if (newName) downloadCSV(win.tableName, newName);
      });
    }
  }

  async function writeToHandle(tableName, handle) {
    const t = tables[tableName];
    if (!t) return;
    setStatus(`Saving ${t.filename || tableName}...`, 'working');
    const t0 = performance.now();
    const writable = await handle.createWritable();
    await writable.write(serializeHeader(t));
    const CHUNK = 50000;
    for (let i = 0; i < t.rows.length; i += CHUNK) {
      const chunk = serializeChunk(t, i, Math.min(i + CHUNK, t.rows.length));
      await writable.write(chunk);
    }
    await writable.close();
    const filename = handle.name;
    t.delimiter = delimiterForExt(filename) || t.delimiter || ',';
    t.modified = false;
    t.filename = filename;
    updateWindowTitle(tableName);
    const elapsed = performance.now() - t0;
    setStatus(`Saved ${filename} (${t.rows.length.toLocaleString()} rows) in ${formatElapsed(elapsed)}`, 'success');
  }

  async function downloadCSV(tableName, filename) {
    flushAllSyncs();
    const t = tables[tableName];
    if (!t) return;
    setStatus(`Saving ${filename}...`, 'working');
    // Defer so "Saving..." can paint
    await new Promise(r => setTimeout(r, 0));
    const t0 = performance.now();
    const parts = [serializeHeader(t)];
    const CHUNK = 50000;
    for (let i = 0; i < t.rows.length; i += CHUNK) {
      parts.push(serializeChunk(t, i, Math.min(i + CHUNK, t.rows.length)));
    }
    const blob = new Blob(parts, { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
    t.modified = false;
    t.filename = filename;
    updateWindowTitle(tableName);
    const elapsed = performance.now() - t0;
    setStatus(`Saved ${filename} (${t.rows.length.toLocaleString()} rows) in ${formatElapsed(elapsed)}`, 'success');
  }

  function escapeField(val, delim) {
    const s = String(val ?? '');
    if (s.includes(delim) || s.includes('"') || s.includes('\n') || s.includes('\r')) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  }

  function serializeHeader(t) {
    const d = t.delimiter || ',';
    return t.columns.map(c => escapeField(c, d)).join(d) + '\r\n';
  }

  function serializeChunk(t, start, end) {
    const cols = t.columns;
    const d = t.delimiter || ',';
    let out = '';
    for (let i = start; i < end; i++) {
      const row = t.rows[i];
      for (let c = 0; c < cols.length; c++) {
        if (c > 0) out += d;
        out += escapeField(row[cols[c]], d);
      }
      out += '\r\n';
    }
    return out;
  }

  function newTable() {
    showPrompt('New Table', 'Table name:', '', (name) => {
      if (!name) return;
      const safeName = sanitizeTableName(name);
      const uniqueName = getUniqueTableName(safeName);
      showPrompt('Columns', 'Column names (comma-separated):', 'id, name, value', (colStr) => {
        if (!colStr) return;
        const columns = colStr.split(',').map(c => c.trim()).filter(Boolean);
        tables[uniqueName] = { columns, rows: [], filename: null, modified: true };
        registerTable(uniqueName);
        createTableWindow(uniqueName);
      });
    });
  }

  // ---- Window Management ----
  function createSubwindow(title, contentFn, opts = {}) {
    const id = nextWinId++;
    const area = document.getElementById('window-area');
    const rect = area.getBoundingClientRect();
    const cascadeOffset = ((windows.length) % 8) * 30;
    const w = opts.width || Math.min(700, rect.width - 40);
    const h = opts.height || Math.min(400, rect.height - 40);
    const x = opts.x ?? Math.min(cascadeOffset + 20, rect.width - w - 10);
    const y = opts.y ?? Math.min(cascadeOffset + 20, rect.height - h - 10);

    const el = document.createElement('div');
    el.className = 'subwindow';
    el.id = `win-${id}`;
    el.style.left = x + 'px';
    el.style.top = y + 'px';
    el.style.width = w + 'px';
    el.style.height = h + 'px';
    el.style.zIndex = ++nextZIndex;

    el.innerHTML = `
      <div class="win-titlebar">
        <span class="win-title">${escHtml(title)}</span>
        <div class="win-controls">
          <button class="btn-min" title="Minimize">&#8211;</button>
          <button class="btn-max" title="Maximize">&#9633;</button>
          <button class="btn-close" title="Close">&#10005;</button>
        </div>
      </div>
      <div class="win-body"></div>
      <div class="win-statusbar"><span class="status-left"></span><span class="status-right"></span></div>
      <div class="resize-handle rh-top"></div>
      <div class="resize-handle rh-bottom"></div>
      <div class="resize-handle rh-left"></div>
      <div class="resize-handle rh-right"></div>
      <div class="resize-handle rh-tl"></div>
      <div class="resize-handle rh-tr"></div>
      <div class="resize-handle rh-bl"></div>
      <div class="resize-handle rh-br"></div>
    `;

    area.appendChild(el);

    const winObj = {
      id, el, title,
      tableName: opts.tableName || null,
      isQuery: opts.isQuery || false,
      maximized: false,
      prevBounds: null,
      sortCols: [],   // [{col, dir:'asc'|'desc'}, ...]
      filterText: '',
    };
    windows.push(winObj);

    setupWindowDrag(winObj);
    setupWindowResize(winObj);
    setupWindowButtons(winObj);

    el.addEventListener('mousedown', () => focusWindow(id));

    if (contentFn) contentFn(winObj, el.querySelector('.win-body'));
    focusWindow(id);
    updateWindowsList();
    return winObj;
  }

  function focusWindow(id) {
    activeWinId = id;
    windows.forEach(w => w.el.classList.toggle('active', w.id === id));
    const win = windows.find(w => w.id === id);
    if (win) win.el.style.zIndex = ++nextZIndex;
  }

  function closeWindow(id) {
    const idx = windows.findIndex(w => w.id === id);
    if (idx === -1) return;
    const win = windows[idx];
    // If it's a table window (not a query result), also remove from tables
    if (win.tableName && !win.isQuery && tables[win.tableName]) {
      const t = tables[win.tableName];
      if (t.modified) {
        if (!confirm(`Table "${win.tableName}" has unsaved changes. Close anyway?`)) return;
      }
      try { db.run(`DROP TABLE IF EXISTS [${win.tableName}]`); } catch (e) {}
      delete tables[win.tableName];
    }
    win.el.remove();
    windows.splice(idx, 1);
    if (activeWinId === id) {
      activeWinId = windows.length ? windows[windows.length - 1].id : null;
      if (activeWinId) focusWindow(activeWinId);
    }
    updateWindowsList();
  }

  function minimizeWindow(id) {
    const win = windows.find(w => w.id === id);
    if (win) {
      win.el.classList.add('minimized');
      updateWindowsList();
    }
  }

  function restoreWindow(id) {
    const win = windows.find(w => w.id === id);
    if (win) {
      win.el.classList.remove('minimized');
      focusWindow(id);
      updateWindowsList();
    }
  }

  function toggleMaximize(id) {
    const win = windows.find(w => w.id === id);
    if (!win) return;
    const area = document.getElementById('window-area');
    const rect = area.getBoundingClientRect();
    if (win.maximized) {
      const b = win.prevBounds;
      win.el.style.left = b.left + 'px';
      win.el.style.top = b.top + 'px';
      win.el.style.width = b.width + 'px';
      win.el.style.height = b.height + 'px';
      win.maximized = false;
    } else {
      win.prevBounds = {
        left: parseInt(win.el.style.left),
        top: parseInt(win.el.style.top),
        width: parseInt(win.el.style.width),
        height: parseInt(win.el.style.height),
      };
      win.el.style.left = '0px';
      win.el.style.top = '0px';
      win.el.style.width = rect.width + 'px';
      win.el.style.height = rect.height + 'px';
      win.maximized = true;
    }
  }

  function setupWindowDrag(win) {
    const titlebar = win.el.querySelector('.win-titlebar');
    let dragging = false, startX, startY, origX, origY;

    titlebar.addEventListener('mousedown', (e) => {
      if (e.target.tagName === 'BUTTON') return;
      dragging = true;
      startX = e.clientX;
      startY = e.clientY;
      origX = parseInt(win.el.style.left);
      origY = parseInt(win.el.style.top);
      titlebar.style.cursor = 'grabbing';
      e.preventDefault();
    });

    document.addEventListener('mousemove', (e) => {
      if (!dragging) return;
      win.el.style.left = (origX + e.clientX - startX) + 'px';
      win.el.style.top = (origY + e.clientY - startY) + 'px';
      if (win.maximized) win.maximized = false;
    });

    document.addEventListener('mouseup', () => {
      if (dragging) {
        dragging = false;
        titlebar.style.cursor = 'grab';
      }
    });
  }

  function setupWindowResize(win) {
    const handles = win.el.querySelectorAll('.resize-handle');
    handles.forEach(handle => {
      let resizing = false, startX, startY, origW, origH, origLeft, origTop;
      const cl = handle.classList;
      const resizeR = cl.contains('rh-right') || cl.contains('rh-tr') || cl.contains('rh-br');
      const resizeB = cl.contains('rh-bottom') || cl.contains('rh-bl') || cl.contains('rh-br');
      const resizeL = cl.contains('rh-left') || cl.contains('rh-tl') || cl.contains('rh-bl');
      const resizeT = cl.contains('rh-top') || cl.contains('rh-tl') || cl.contains('rh-tr');

      handle.addEventListener('mousedown', (e) => {
        resizing = true;
        startX = e.clientX;
        startY = e.clientY;
        origW = parseInt(win.el.style.width);
        origH = parseInt(win.el.style.height);
        origLeft = parseInt(win.el.style.left);
        origTop = parseInt(win.el.style.top);
        e.preventDefault();
        e.stopPropagation();
      });

      document.addEventListener('mousemove', (e) => {
        if (!resizing) return;
        const dx = e.clientX - startX;
        const dy = e.clientY - startY;
        if (resizeR) win.el.style.width = Math.max(280, origW + dx) + 'px';
        if (resizeB) win.el.style.height = Math.max(160, origH + dy) + 'px';
        if (resizeL) {
          const newW = Math.max(280, origW - dx);
          win.el.style.width = newW + 'px';
          win.el.style.left = (origLeft + origW - newW) + 'px';
        }
        if (resizeT) {
          const newH = Math.max(160, origH - dy);
          win.el.style.height = newH + 'px';
          win.el.style.top = (origTop + origH - newH) + 'px';
        }
      });

      document.addEventListener('mouseup', () => { resizing = false; });
    });
  }

  function setupWindowButtons(win) {
    win.el.querySelector('.btn-close').addEventListener('click', () => closeWindow(win.id));
    win.el.querySelector('.btn-min').addEventListener('click', () => minimizeWindow(win.id));
    win.el.querySelector('.btn-max').addEventListener('click', () => toggleMaximize(win.id));
    // Double-click title text to rename, double-click elsewhere on titlebar to maximize
    win.el.querySelector('.win-title').addEventListener('dblclick', (e) => {
      e.stopPropagation();
      if (win.tableName) startInlineRename(win);
    });
    win.el.querySelector('.win-titlebar').addEventListener('dblclick', (e) => {
      if (e.target.tagName !== 'BUTTON' && !e.target.classList.contains('win-title')) {
        toggleMaximize(win.id);
      }
    });
  }

  function startInlineRename(win) {
    const oldName = win.tableName;
    const t = tables[oldName];
    if (!t) return;
    const titleEl = win.el.querySelector('.win-title');
    const input = document.createElement('input');
    input.className = 'inline-rename';
    input.value = oldName;
    titleEl.textContent = '';
    titleEl.appendChild(input);
    input.focus();
    input.select();

    function commit() {
      const raw = input.value.trim();
      // Remove input and restore text
      if (input.parentNode) input.remove();
      if (!raw || raw === oldName) {
        updateWindowTitle(oldName);
        return;
      }
      const newName = sanitizeTableName(raw);
      if (newName === oldName) {
        updateWindowTitle(oldName);
        return;
      }
      const uniqueName = tables[newName] ? getUniqueTableName(newName) : newName;

      // Move table data
      tables[uniqueName] = t;
      delete tables[oldName];

      // Rename in SQL (instant, no data re-insertion)
      try { db.run(`ALTER TABLE [${oldName}] RENAME TO [${uniqueName}]`); } catch (_) {}

      // Update all windows referencing this table
      windows.filter(w => w.tableName === oldName).forEach(w => {
        w.tableName = uniqueName;
        w.title = uniqueName;
      });
      updateWindowTitle(uniqueName);
      updateWindowsList();
      setStatus(`Renamed "${oldName}" to "${uniqueName}"`, 'success');
    }

    let done = false;
    function finish() {
      if (done) return;
      done = true;
      commit();
    }

    input.addEventListener('blur', finish);
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); input.blur(); }
      if (e.key === 'Escape') { e.preventDefault(); input.value = oldName; input.blur(); }
      e.stopPropagation();
    });
    input.addEventListener('mousedown', (e) => e.stopPropagation());
  }

  function getActiveDataWindow() {
    return windows.find(w => w.id === activeWinId) || null;
  }

  function updateWindowTitle(tableName) {
    windows.filter(w => w.tableName === tableName).forEach(w => {
      const t = tables[tableName];
      const mod = t && t.modified ? ' *' : '';
      w.el.querySelector('.win-title').textContent = tableName + mod;
    });
  }

  function updateWindowsList() {
    const list = document.getElementById('windows-list');
    list.innerHTML = '';
    if (windows.length === 0) {
      list.innerHTML = '<button disabled style="color:var(--text-dim)">No windows</button>';
      return;
    }
    windows.forEach(w => {
      const btn = document.createElement('button');
      const minimized = w.el.classList.contains('minimized');
      btn.textContent = (minimized ? '[_] ' : '') + w.title;
      btn.addEventListener('click', () => {
        if (minimized) restoreWindow(w.id);
        else focusWindow(w.id);
      });
      list.appendChild(btn);
    });
  }

  // ---- Layout ----
  function getVisibleWindows() {
    return windows.filter(w => !w.el.classList.contains('minimized'));
  }

  function layoutTileH() {
    const vw = getVisibleWindows();
    if (!vw.length) return;
    const area = document.getElementById('window-area').getBoundingClientRect();
    const h = area.height / vw.length;
    vw.forEach((w, i) => {
      w.el.style.left = '0px';
      w.el.style.top = Math.round(i * h) + 'px';
      w.el.style.width = area.width + 'px';
      w.el.style.height = Math.round(h) + 'px';
      w.maximized = false;
    });
  }

  function layoutTileV() {
    const vw = getVisibleWindows();
    if (!vw.length) return;
    const area = document.getElementById('window-area').getBoundingClientRect();
    const w = area.width / vw.length;
    vw.forEach((win, i) => {
      win.el.style.left = Math.round(i * w) + 'px';
      win.el.style.top = '0px';
      win.el.style.width = Math.round(w) + 'px';
      win.el.style.height = area.height + 'px';
      win.maximized = false;
    });
  }

  function layoutGrid() {
    const vw = getVisibleWindows();
    if (!vw.length) return;
    const area = document.getElementById('window-area').getBoundingClientRect();
    const cols = Math.ceil(Math.sqrt(vw.length));
    const rows = Math.ceil(vw.length / cols);
    const cellW = area.width / cols;
    const cellH = area.height / rows;
    vw.forEach((win, i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      win.el.style.left = Math.round(col * cellW) + 'px';
      win.el.style.top = Math.round(row * cellH) + 'px';
      win.el.style.width = Math.round(cellW) + 'px';
      win.el.style.height = Math.round(cellH) + 'px';
      win.maximized = false;
    });
  }

  function layoutCascade() {
    const vw = getVisibleWindows();
    if (!vw.length) return;
    const area = document.getElementById('window-area').getBoundingClientRect();
    const w = Math.min(600, area.width * 0.7);
    const h = Math.min(400, area.height * 0.7);
    vw.forEach((win, i) => {
      const offset = (i % 10) * 30;
      win.el.style.left = (20 + offset) + 'px';
      win.el.style.top = (20 + offset) + 'px';
      win.el.style.width = w + 'px';
      win.el.style.height = h + 'px';
      win.maximized = false;
      win.el.style.zIndex = ++nextZIndex;
    });
  }

  function minimizeAll() {
    windows.forEach(w => w.el.classList.add('minimized'));
    updateWindowsList();
  }

  function restoreAll() {
    windows.forEach(w => w.el.classList.remove('minimized'));
    updateWindowsList();
  }

  // ---- Table Window ----
  function createTableWindow(tableName) {
    const t = tables[tableName];
    createSubwindow(tableName, (win, body) => {
      win.tableName = tableName;
      renderTableView(win, body, t);
    }, { tableName });
  }

  function renderTableView(win, body, tableData) {
    body.innerHTML = '';

    // Toolbar
    const toolbar = document.createElement('div');
    toolbar.className = 'win-toolbar';
    toolbar.innerHTML = `
      <label>Filter:</label>
      <input type="text" class="filter-input" placeholder="Filter: text  col>5  col~regex  col=val" value="${escHtml(win.filterText)}">
      <button class="btn-add-row">+ Row</button>
      <button class="btn-add-col">+ Col</button>
    `;
    body.appendChild(toolbar);

    const filterInput = toolbar.querySelector('.filter-input');
    let filterTimeout;
    filterInput.addEventListener('input', () => {
      clearTimeout(filterTimeout);
      filterTimeout = setTimeout(() => {
        win.filterText = filterInput.value;
        rebuildTable(win);
      }, 200);
    });

    toolbar.querySelector('.btn-add-row').addEventListener('click', () => {
      addRow(win.tableName);
      rebuildTable(win);
      // Scroll to bottom to show new row
      const container = win.el.querySelector('.table-container');
      if (container) container.scrollTop = container.scrollHeight;
    });

    toolbar.querySelector('.btn-add-col').addEventListener('click', () => {
      showPrompt('Add Column', 'Column name:', '', (colName) => {
        if (!colName) return;
        addColumn(win.tableName, colName);
        rebuildTable(win);
      });
    });

    // Table container
    const container = document.createElement('div');
    container.className = 'table-container';
    body.appendChild(container);

    buildTableHTML(win, container, tableData);
  }

  // Parse filter text into predicates: "col>5 col2~regex col3=foo bar"
  // Tokens: "col op value" for column filters, bare words for global search
  function parseFilters(filterText, columns) {
    if (!filterText) return null;
    const predicates = [];
    const colSet = new Set(columns);
    // Match: colName operator value (operator is >=, <=, !=, >, <, =, ~)
    const tokenRe = /(\S+?)(>=|<=|!=|>|<|=|~)(\S+)|(\S+)/g;
    let m;
    while ((m = tokenRe.exec(filterText)) !== null) {
      if (m[1] && colSet.has(m[1])) {
        const col = m[1], op = m[2], val = m[3];
        if (op === '~') {
          try {
            const re = new RegExp(val, 'i');
            predicates.push(row => re.test(String(row[col] ?? '')));
          } catch (_) {
            // Invalid regex — treat as literal
            const lower = val.toLowerCase();
            predicates.push(row => String(row[col] ?? '').toLowerCase().includes(lower));
          }
        } else {
          const numVal = Number(val);
          const isNum = !isNaN(numVal) && val !== '';
          predicates.push(row => {
            const cell = row[col] ?? '';
            const cellNum = Number(cell);
            const useNum = isNum && !isNaN(cellNum) && cell !== '';
            if (useNum) {
              if (op === '>') return cellNum > numVal;
              if (op === '<') return cellNum < numVal;
              if (op === '>=') return cellNum >= numVal;
              if (op === '<=') return cellNum <= numVal;
              if (op === '=') return cellNum === numVal;
              if (op === '!=') return cellNum !== numVal;
            }
            const sv = String(cell), sv2 = val;
            if (op === '=') return sv === sv2;
            if (op === '!=') return sv !== sv2;
            const cmp = collator.compare(sv, sv2);
            if (op === '>') return cmp > 0;
            if (op === '<') return cmp < 0;
            if (op === '>=') return cmp >= 0;
            if (op === '<=') return cmp <= 0;
            return false;
          });
        }
      } else {
        // Global search token
        const q = (m[4] || m[0]).toLowerCase();
        predicates.push(row => columns.some(c => String(row[c] ?? '').toLowerCase().includes(q)));
      }
    }
    return predicates.length > 0 ? predicates : null;
  }

  function buildTableHTML(win, container, tableData) {
    const { columns, rows } = tableData;
    let displayRows = [...rows];

    // Filter
    const predicates = parseFilters(win.filterText, columns);
    if (predicates) {
      displayRows = displayRows.filter(row => predicates.every(p => p(row)));
    }

    // Multi-column sort
    if (win.sortCols.length > 0) {
      displayRows.sort((a, b) => {
        for (const { col, dir } of win.sortCols) {
          const m = dir === 'asc' ? 1 : -1;
          const va = a[col] ?? '', vb = b[col] ?? '';
          const na = Number(va), nb = Number(vb);
          if (!isNaN(na) && !isNaN(nb) && va !== '' && vb !== '') {
            if (na !== nb) return (na - nb) * m;
          } else {
            const cmp = collator.compare(String(va), String(vb));
            if (cmp !== 0) return cmp * m;
          }
        }
        return 0;
      });
    }

    // Store display rows for virtual scrolling
    win._displayRows = displayRows;
    win._columns = columns;
    win._container = container;

    const table = document.createElement('table');
    table.className = 'data-table';

    // Header
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    const rowNumTh = document.createElement('th');
    rowNumTh.className = 'row-num-header';
    rowNumTh.textContent = '#';
    headerRow.appendChild(rowNumTh);

    columns.forEach(col => {
      const th = document.createElement('th');
      th.textContent = col;
      const sortIdx = win.sortCols.findIndex(s => s.col === col);
      if (sortIdx !== -1) {
        const arrow = document.createElement('span');
        arrow.className = 'sort-arrow';
        const dir = win.sortCols[sortIdx].dir;
        arrow.textContent = dir === 'asc' ? '\u25B2' : '\u25BC';
        if (win.sortCols.length > 1) arrow.textContent += (sortIdx + 1);
        th.appendChild(arrow);
      }
      // Single click: sort; double-click: rename — use timer to distinguish
      let clickTimer = null;
      th.addEventListener('click', (e) => {
        if (clickTimer) { clearTimeout(clickTimer); clickTimer = null; return; }
        clickTimer = setTimeout(() => {
          clickTimer = null;
          const existing = win.sortCols.findIndex(s => s.col === col);
          if (e.shiftKey) {
            if (existing !== -1) {
              if (win.sortCols[existing].dir === 'asc') {
                win.sortCols[existing].dir = 'desc';
              } else {
                win.sortCols.splice(existing, 1);
              }
            } else {
              win.sortCols.push({ col, dir: 'asc' });
            }
          } else {
            if (existing !== -1 && win.sortCols.length === 1) {
              if (win.sortCols[0].dir === 'asc') {
                win.sortCols[0].dir = 'desc';
              } else {
                win.sortCols = [];
              }
            } else {
              win.sortCols = [{ col, dir: 'asc' }];
            }
          }
          rebuildTable(win);
        }, 250);
      });
      th.addEventListener('dblclick', (e) => {
        e.stopPropagation();
        startColumnRename(win, th, col);
      });
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    // Body — virtual scrolling renders only visible rows
    const tbody = document.createElement('tbody');
    table.appendChild(tbody);
    win._tbody = tbody;
    win._table = table;

    // Event delegation on table — replaces per-cell listeners
    table.addEventListener('blur', (e) => {
      const td = e.target;
      if (td.tagName !== 'TD' || !td.getAttribute('contenteditable')) return;
      const tr = td.parentElement;
      const displayIdx = parseInt(tr.dataset.displayIdx, 10);
      const colIdx = parseInt(td.dataset.colIdx, 10);
      if (isNaN(displayIdx) || isNaN(colIdx)) return;
      const row = win._displayRows[displayIdx];
      const col = win._columns[colIdx];
      if (!row || col == null) return;
      const newVal = td.textContent;
      if (newVal !== String(row[col] ?? '')) {
        row[col] = newVal;
        td.classList.add('modified');
        markModified(win.tableName);
        const t = tables[win.tableName];
        if (t) {
          if (!t._dirtyCells) t._dirtyCells = [];
          t._dirtyCells.push({ rownum: row._rownum, col, value: newVal });
        }
        debouncedSync(win.tableName);
      }
    }, true); // capture phase for blur

    table.addEventListener('keydown', (e) => {
      const td = e.target;
      if (td.tagName !== 'TD' || !td.getAttribute('contenteditable')) return;
      const tr = td.parentElement;
      if (e.key === 'Tab') {
        e.preventDefault();
        const cells = [...tr.querySelectorAll('td[contenteditable]')];
        const idx = cells.indexOf(td);
        const next = e.shiftKey ? cells[idx - 1] : cells[idx + 1];
        if (next) next.focus();
      } else if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        td.blur();
        const nextTr = tr.nextElementSibling;
        if (nextTr && !nextTr.classList.contains('virtual-pad')) {
          const colIdx = [...tr.children].indexOf(td);
          const nextTd = nextTr.children[colIdx];
          if (nextTd && nextTd.getAttribute('contenteditable')) nextTd.focus();
        }
      } else if (e.key === 'Escape') {
        const displayIdx = parseInt(tr.dataset.displayIdx, 10);
        const colIdx = parseInt(td.dataset.colIdx, 10);
        if (!isNaN(displayIdx) && !isNaN(colIdx)) {
          const row = win._displayRows[displayIdx];
          const col = win._columns[colIdx];
          if (row && col != null) td.textContent = row[col] ?? '';
        }
        td.blur();
      }
    });

    table.addEventListener('contextmenu', (e) => {
      const td = e.target.closest('td.row-num');
      if (!td) return;
      e.preventDefault();
      const tr = td.parentElement;
      const displayIdx = parseInt(tr.dataset.displayIdx, 10);
      if (isNaN(displayIdx)) return;
      const row = win._displayRows[displayIdx];
      if (row) showRowContextMenu(e.clientX, e.clientY, win.tableName, row._rownum, win);
    });

    container.innerHTML = '';
    container.appendChild(table);

    // Reset render range tracking
    win._renderStart = -1;
    win._renderEnd = -1;

    // Initial render of visible rows
    renderVisibleRows(win);

    // Scroll listener for virtual scrolling
    let scrollRaf = 0;
    container.addEventListener('scroll', () => {
      if (scrollRaf) return;
      scrollRaf = requestAnimationFrame(() => {
        scrollRaf = 0;
        renderVisibleRows(win);
      });
    });

    // ResizeObserver to re-render when window resizes
    if (win._resizeObserver) win._resizeObserver.disconnect();
    win._resizeObserver = new ResizeObserver(() => {
      renderVisibleRows(win);
    });
    win._resizeObserver.observe(container);

    // Update statusbar
    const statusLeft = win.el.querySelector('.status-left');
    const statusRight = win.el.querySelector('.status-right');
    statusLeft.textContent = `${displayRows.length} of ${rows.length} rows`;
    statusRight.textContent = `${columns.length} columns`;
  }

  function renderVisibleRows(win) {
    const container = win._container;
    const tbody = win._tbody;
    const displayRows = win._displayRows;
    const columns = win._columns;
    if (!container || !tbody || !displayRows) return;

    const totalRows = displayRows.length;
    const scrollTop = container.scrollTop;
    const clientHeight = container.clientHeight;

    // Account for thead height
    const theadHeight = win._table.querySelector('thead')?.offsetHeight || 0;
    const adjustedScrollTop = Math.max(0, scrollTop - theadHeight);

    let startIdx = Math.floor(adjustedScrollTop / ROW_HEIGHT) - OVERSCAN;
    let endIdx = Math.ceil((adjustedScrollTop + clientHeight) / ROW_HEIGHT) + OVERSCAN;
    startIdx = Math.max(0, startIdx);
    endIdx = Math.min(totalRows, endIdx);

    // Early return if range unchanged
    if (startIdx === win._renderStart && endIdx === win._renderEnd) return;

    // Blur active cell if it's inside this tbody
    const active = document.activeElement;
    if (active && tbody.contains(active)) active.blur();

    win._renderStart = startIdx;
    win._renderEnd = endIdx;

    const colCount = columns.length + 1; // +1 for row number column

    // Build new tbody content
    const fragment = document.createDocumentFragment();

    // Top padding row
    if (startIdx > 0) {
      const padTr = document.createElement('tr');
      padTr.className = 'virtual-pad';
      const padTd = document.createElement('td');
      padTd.setAttribute('colspan', colCount);
      padTd.style.height = (startIdx * ROW_HEIGHT) + 'px';
      padTr.appendChild(padTd);
      fragment.appendChild(padTr);
    }

    // Visible rows
    for (let i = startIdx; i < endIdx; i++) {
      const row = displayRows[i];
      const tr = document.createElement('tr');
      tr.dataset.displayIdx = i;

      const numTd = document.createElement('td');
      numTd.className = 'row-num';
      numTd.textContent = row._rownum;
      tr.appendChild(numTd);

      for (let c = 0; c < columns.length; c++) {
        const td = document.createElement('td');
        td.textContent = row[columns[c]] ?? '';
        td.setAttribute('contenteditable', 'true');
        td.dataset.colIdx = c;
        tr.appendChild(td);
      }
      fragment.appendChild(tr);
    }

    // Bottom padding row
    if (endIdx < totalRows) {
      const padTr = document.createElement('tr');
      padTr.className = 'virtual-pad';
      const padTd = document.createElement('td');
      padTd.setAttribute('colspan', colCount);
      padTd.style.height = ((totalRows - endIdx) * ROW_HEIGHT) + 'px';
      padTr.appendChild(padTd);
      fragment.appendChild(padTr);
    }

    tbody.innerHTML = '';
    tbody.appendChild(fragment);
  }

  function rebuildTable(win) {
    const t = tables[win.tableName];
    if (!t) return;
    const container = win.el.querySelector('.table-container');
    if (!container) return;

    const oldFilter = win._lastFilterText;
    const filterChanged = oldFilter !== undefined && oldFilter !== win.filterText;
    win._lastFilterText = win.filterText;

    buildTableHTML(win, container, t);

    if (filterChanged) {
      container.scrollTop = 0;
    }
  }

  function addRow(tableName) {
    const t = tables[tableName];
    const newRow = { _rownum: t.rows.length + 1 };
    t.columns.forEach(c => { newRow[c] = ''; });
    t.rows.push(newRow);
    markModified(tableName);
    debouncedSync(tableName);
  }

  function deleteRow(tableName, rownum, win) {
    const t = tables[tableName];
    t.rows = t.rows.filter(r => r._rownum !== rownum);
    // Renumber
    t.rows.forEach((r, i) => { r._rownum = i + 1; });
    markModified(tableName);
    debouncedSync(tableName);
    rebuildTable(win);
  }

  function addColumn(tableName, colName) {
    const t = tables[tableName];
    if (t.columns.includes(colName)) return;
    t.columns.push(colName);
    t.rows.forEach(r => { r[colName] = ''; });
    markModified(tableName);
    registerTable(tableName);
  }

  function renameColumn(tableName, oldCol, newCol, win) {
    const t = tables[tableName];
    if (!t || oldCol === newCol) return;
    if (t.columns.includes(newCol)) return; // duplicate
    const idx = t.columns.indexOf(oldCol);
    if (idx === -1) return;
    t.columns[idx] = newCol;
    for (const row of t.rows) {
      row[newCol] = row[oldCol];
      delete row[oldCol];
    }
    // Update sort references
    for (const s of win.sortCols) {
      if (s.col === oldCol) s.col = newCol;
    }
    markModified(tableName);
    try { db.run(`ALTER TABLE [${tableName}] RENAME COLUMN [${oldCol}] TO [${newCol}]`); } catch (_) {}
    rebuildTable(win);
  }

  function startColumnRename(win, th, oldCol) {
    const input = document.createElement('input');
    input.className = 'inline-rename';
    input.value = oldCol;
    th.textContent = '';
    th.appendChild(input);
    input.focus();
    input.select();

    let done = false;
    function commit() {
      if (done) return;
      done = true;
      const raw = input.value.trim();
      if (input.parentNode) input.remove();
      if (!raw || raw === oldCol) {
        rebuildTable(win);
        return;
      }
      renameColumn(win.tableName, oldCol, raw, win);
    }
    input.addEventListener('blur', commit);
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); commit(); }
      if (e.key === 'Escape') { done = true; input.remove(); rebuildTable(win); }
    });
  }

  function markModified(tableName) {
    const t = tables[tableName];
    if (t) t.modified = true;
    updateWindowTitle(tableName);
  }

  function syncToSQL(tableName) {
    const t = tables[tableName];
    if (!t || !db || t.columns.length === 0) return;

    // If we have targeted dirty cells, apply only those changes
    if (t._dirtyCells && t._dirtyCells.length > 0) {
      try {
        db.run('BEGIN TRANSACTION');
        for (const { rownum, col, value } of t._dirtyCells) {
          db.run(`UPDATE [${tableName}] SET [${col}] = ? WHERE rowid = ?`, [value, rownum]);
        }
        db.run('COMMIT');
      } catch (e) {
        try { db.run('ROLLBACK'); } catch (_) {}
      }
      t._dirtyCells = [];
      return;
    }

    // Full resync fallback (for add/delete row, add column, etc.)
    try {
      db.run('BEGIN TRANSACTION');
      db.run(`DELETE FROM [${tableName}]`);
      const placeholders = t.columns.map(() => '?').join(', ');
      const stmt = db.prepare(`INSERT INTO [${tableName}] VALUES (${placeholders})`);
      for (const row of t.rows) {
        stmt.run(t.columns.map(c => row[c] ?? ''));
      }
      stmt.free();
      db.run('COMMIT');
    } catch (e) {
      try { db.run('ROLLBACK'); } catch (_) {}
    }
  }

  function debouncedSync(tableName, delay = 300) {
    clearTimeout(syncTimers[tableName]);
    syncTimers[tableName] = setTimeout(() => {
      delete syncTimers[tableName];
      syncToSQL(tableName);
    }, delay);
  }

  function flushAllSyncs() {
    for (const name of Object.keys(syncTimers)) {
      clearTimeout(syncTimers[name]);
      delete syncTimers[name];
      syncToSQL(name);
    }
  }

  // ---- Row Context Menu ----
  function showRowContextMenu(x, y, tableName, rownum, win) {
    removeContextMenu();
    const menu = document.createElement('div');
    menu.className = 'context-menu';
    menu.style.left = x + 'px';
    menu.style.top = y + 'px';

    const insertBtn = document.createElement('button');
    insertBtn.textContent = 'Insert Row Above';
    insertBtn.addEventListener('click', () => {
      const t = tables[tableName];
      const idx = t.rows.findIndex(r => r._rownum === rownum);
      const newRow = { _rownum: 0 };
      t.columns.forEach(c => { newRow[c] = ''; });
      t.rows.splice(idx, 0, newRow);
      t.rows.forEach((r, i) => { r._rownum = i + 1; });
      markModified(tableName);
      syncToSQL(tableName);
      rebuildTable(win);
      removeContextMenu();
    });
    menu.appendChild(insertBtn);

    const delBtn = document.createElement('button');
    delBtn.textContent = 'Delete Row';
    delBtn.addEventListener('click', () => {
      deleteRow(tableName, rownum, win);
      removeContextMenu();
    });
    menu.appendChild(delBtn);

    document.body.appendChild(menu);
    setTimeout(() => {
      document.addEventListener('click', removeContextMenu, { once: true });
    }, 0);
  }

  function removeContextMenu() {
    document.querySelectorAll('.context-menu').forEach(m => m.remove());
  }

  // ---- SQL Console ----
  function setupKeyboard() {
    document.getElementById('sql-input').addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        executeQuery();
      }
    });

    document.addEventListener('keydown', (e) => {
      const mod = e.ctrlKey || e.metaKey;
      if (!mod) return;
      const key = e.key.toLowerCase();

      if (key === 'o' && !e.shiftKey) {
        e.preventDefault();
        openFile();
      } else if (key === 's' && !e.shiftKey) {
        e.preventDefault();
        saveActiveTable();
      } else if (key === 's' && e.shiftKey) {
        e.preventDefault();
        saveActiveTableAs();
      } else if (key === 'n' && !e.shiftKey) {
        e.preventDefault();
        newTable();
      } else if (key === 'w' && !e.shiftKey) {
        e.preventDefault();
        if (activeWinId) closeWindow(activeWinId);
      }
    });
  }

  function autoQuoteSQL(sql) {
    const names = Object.keys(tables).sort((a, b) => b.length - a.length);
    for (const name of names) {
      if (/^[a-zA-Z_][a-zA-Z0-9_]*$/.test(name)) continue;
      const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      sql = sql.replace(new RegExp('(?<!\\[)\\b' + escaped + '\\b(?!\\])', 'g'), '[' + name + ']');
    }
    return sql;
  }

  function extractIntoClause(sql) {
    // Match INTO [tablename] anywhere in a SELECT statement and strip it out
    // Supports: SELECT ... INTO name FROM ..., SELECT ... FROM ... INTO name WHERE ...
    const intoPattern = /\bINTO\s+\[?([^\]\s,;]+)\]?/i;
    const match = sql.match(intoPattern);
    if (match && /^\s*SELECT\b/i.test(sql)) {
      const targetName = match[1];
      const selectSQL = sql.replace(intoPattern, ' ').replace(/\s+/g, ' ').trim();
      return { targetName, selectSQL };
    }
    return null;
  }

  // Convert sql.js exec result [{columns, values}] to array of row objects
  function sqlResultToRows(result) {
    if (!result || result.length === 0) return [];
    const { columns, values } = result[0];
    return values.map(row => {
      const obj = {};
      columns.forEach((col, i) => { obj[col] = row[i]; });
      return obj;
    });
  }

  // List tables currently in the SQLite database
  function getDBTables() {
    const result = db.exec("SELECT name FROM sqlite_master WHERE type='table'");
    if (!result.length) return [];
    return result[0].values.map(r => r[0]);
  }

  // Max rows to materialize from a query result into a window
  const MAX_RESULT_ROWS = 100000;

  function executeQuery() {
    const input = document.getElementById('sql-input');
    const sql = autoQuoteSQL(input.value.trim());
    if (!sql) return;

    setStatus('Running query...', 'working');

    // Defer execution so browser can paint the "working" status
    setTimeout(() => {
      try {
        flushAllSyncs();
        const t0 = performance.now();

        // Handle SELECT ... INTO ... by running the SELECT and creating the table from results
        const intoInfo = extractIntoClause(sql);
        if (intoInfo) {
          const result = db.exec(intoInfo.selectSQL);
          const elapsed = performance.now() - t0;
          const rows = sqlResultToRows(result);
          if (rows.length === 0) {
            setStatus('INTO query returned no rows', 'error');
            return;
          }
          const name = sanitizeTableName(intoInfo.targetName);
          const uniqueName = tables[name] ? getUniqueTableName(name) : name;
          const rawColumns = Object.keys(rows[0]);
          const columns = sanitizeColumns(rawColumns);
          const tableRows = rows.map((r, i) => {
            const row = { _rownum: i + 1 };
            rawColumns.forEach((raw, j) => { row[columns[j]] = r[raw] != null ? String(r[raw]) : ''; });
            return row;
          });
          tables[uniqueName] = { columns, rows: tableRows, filename: null, modified: true, fileHandle: null };
          registerTable(uniqueName);
          createTableWindow(uniqueName);
          setStatus(`Created table "${uniqueName}" with ${tableRows.length} row(s) in ${formatElapsed(elapsed)}`, 'success');
          return;
        }

        // Snapshot existing DB tables before query
        const tablesBefore = new Set(getDBTables());

        const result = db.exec(sql);
        const elapsed = performance.now() - t0;

        // Detect new tables created by CREATE TABLE etc.
        const newTables = getDBTables().filter(n => !tablesBefore.has(n));

        const createMatch = sql.match(/\bCREATE\s+TABLE\s+(?:IF\s+NOT\s+EXISTS\s+)?\[?([^\]\s,(;]+)\]?/i);
        if (createMatch) {
          const createName = createMatch[1];
          if (!newTables.includes(createName) && !tables[createName]) {
            const dbTables = getDBTables();
            if (dbTables.includes(createName)) newTables.push(createName);
          }
        }

        if (newTables.length > 0) {
          importNewDBTables(newTables);
          setStatus(`Created table(s): ${newTables.join(', ')} in ${formatElapsed(elapsed)}`, 'success');
        } else if (result.length > 0 && result[0].columns && result[0].values.length > 0) {
          const totalRows = result[0].values.length;
          const truncated = totalRows > MAX_RESULT_ROWS;
          if (truncated) {
            result[0].values = result[0].values.slice(0, MAX_RESULT_ROWS);
          }
          const rows = sqlResultToRows(result);
          showQueryResult(sql, rows);
          const suffix = truncated ? ` (showing first ${MAX_RESULT_ROWS.toLocaleString()} of ${totalRows.toLocaleString()})` : '';
          setStatus(`Query returned ${totalRows.toLocaleString()} row(s) in ${formatElapsed(elapsed)}${suffix}`, 'success');
        } else {
          const msg = result.length > 0 ? `${result[0].values.length} row(s) affected` : 'OK';
          setStatus(`${msg} in ${formatElapsed(elapsed)}`, 'success');
          refreshAllTableWindows();
        }
      } catch (e) {
        setStatus(`Error: ${e.message}`, 'error');
      }
    }, 0);
  }

  function importNewDBTables(tableNames) {
    tableNames.forEach(name => {
      try {
        const result = db.exec(`SELECT * FROM [${name}]`);
        const rows = sqlResultToRows(result);
        const rawColumns = result.length > 0 ? result[0].columns : [];
        const columns = sanitizeColumns(rawColumns);
        const tableRows = rows.map((r, i) => {
          const row = { _rownum: i + 1 };
          rawColumns.forEach((raw, j) => { row[columns[j]] = r[raw] != null ? String(r[raw]) : ''; });
          return row;
        });
        tables[name] = { columns, rows: tableRows, filename: null, modified: true, fileHandle: null };
        registerTable(name);
        createTableWindow(name);
      } catch (e) {}
    });
  }

  function showQueryResult(sql, resultRows) {
    const rawColumns = Object.keys(resultRows[0]);
    const columns = sanitizeColumns(rawColumns);
    const tableName = '_query_' + nextWinId;

    // Store as a table so it can be saved
    const rows = resultRows.map((r, i) => {
      const row = { _rownum: i + 1 };
      rawColumns.forEach((raw, j) => { row[columns[j]] = r[raw] != null ? String(r[raw]) : ''; });
      return row;
    });
    tables[tableName] = { columns, rows, filename: null, modified: false };

    createSubwindow(tableName, (win, body) => {
      win.tableName = tableName;
      win.isQuery = true;
      renderTableView(win, body, tables[tableName]);
    }, { tableName, isQuery: true });
  }

  function refreshAllTableWindows() {
    windows.forEach(w => {
      if (w.tableName && tables[w.tableName] && !w.isQuery) {
        // Re-sync from SQL database
        try {
          const t = tables[w.tableName];
          const result = db.exec(`SELECT * FROM [${w.tableName}]`);
          if (result.length > 0) {
            const rows = sqlResultToRows(result);
            t.rows = rows.map((r, i) => {
              const row = { _rownum: i + 1 };
              t.columns.forEach(c => { row[c] = r[c] != null ? String(r[c]) : ''; });
              return row;
            });
          }
        } catch (e) {}
        rebuildTable(w);
      }
    });
  }

  function clearConsole() {
    document.getElementById('sql-input').value = '';
    setStatus('');
  }

  function setStatus(msg, type = '') {
    const el = document.getElementById('console-status');
    el.textContent = msg;
    el.className = type;
  }

  function formatElapsed(ms) {
    if (ms < 1) return '<1ms';
    if (ms < 1000) return Math.round(ms) + 'ms';
    return (ms / 1000).toFixed(2) + 's';
  }

  // ---- Console Resize ----
  function setupConsoleResize() {
    const handle = document.getElementById('console-resize-handle');
    const panel = document.getElementById('console-panel');
    let resizing = false, startY, origH;

    handle.addEventListener('mousedown', (e) => {
      resizing = true;
      startY = e.clientY;
      origH = panel.offsetHeight;
      e.preventDefault();
    });

    document.addEventListener('mousemove', (e) => {
      if (!resizing) return;
      const newH = origH - (e.clientY - startY);
      panel.style.height = Math.max(60, Math.min(window.innerHeight * 0.5, newH)) + 'px';
    });

    document.addEventListener('mouseup', () => { resizing = false; });
  }

  // ---- Menu close on outside click ----
  let _menuBarActive = false;
  let _menuDragging = false;

  function closeMenus() {
    _menuBarActive = false;
    _menuDragging = false;
    document.querySelectorAll('.menu-item').forEach(m => m.classList.remove('open'));
  }

  function setupMenuClose() {
    const menubar = document.getElementById('menubar');
    const allItems = document.querySelectorAll('.menu-item');

    function openItem(item) {
      allItems.forEach(m => m.classList.remove('open'));
      item.classList.add('open');
    }

    allItems.forEach(item => {
      item.querySelector('.menu-label').addEventListener('mousedown', (e) => {
        if (_menuBarActive) {
          closeMenus();
        } else {
          _menuBarActive = true;
          _menuDragging = true;
          openItem(item);
        }
        e.preventDefault();
      });

      // Clicking a label when already active (handled by mousedown above)
      item.querySelector('.menu-label').addEventListener('click', (e) => {
        e.stopPropagation();
      });

      item.addEventListener('mouseenter', () => {
        if (_menuBarActive) openItem(item);
      });
    });

    // On mouseup over a dropdown button during drag, activate it
    menubar.addEventListener('mouseup', (e) => {
      const btn = e.target.closest('.menu-dropdown button');
      if (_menuDragging && btn) {
        btn.click();
        closeMenus();
      }
      _menuDragging = false;
    });

    // Clicking a dropdown button closes the menu bar (non-drag case)
    menubar.addEventListener('click', (e) => {
      if (e.target.closest('.menu-dropdown button')) closeMenus();
    });

    document.addEventListener('mouseup', (e) => {
      if (_menuDragging && !e.target.closest('.menu-item')) {
        closeMenus();
      }
      _menuDragging = false;
    });

    document.addEventListener('click', (e) => {
      if (!e.target.closest('.menu-item')) closeMenus();
    });
  }

  // ---- Modals ----
  function showPrompt(title, label, defaultValue, callback) {
    const overlay = document.createElement('div');
    overlay.className = 'modal-overlay';
    overlay.innerHTML = `
      <div class="modal">
        <h3>${escHtml(title)}</h3>
        <label style="font-size:12px;display:block;margin-bottom:4px;">${escHtml(label)}</label>
        <input type="text" class="modal-input" value="${escHtml(defaultValue)}">
        <div class="modal-buttons">
          <button class="cancel">Cancel</button>
          <button class="primary ok">OK</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);
    const input = overlay.querySelector('.modal-input');
    input.focus();
    input.select();

    const close = (val) => { overlay.remove(); callback(val); };
    overlay.querySelector('.cancel').addEventListener('click', () => close(null));
    overlay.querySelector('.ok').addEventListener('click', () => close(input.value));
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') close(input.value);
      if (e.key === 'Escape') close(null);
    });
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) close(null);
    });
  }

  // ---- Utilities ----
  function escHtml(str) {
    const d = document.createElement('div');
    d.textContent = str;
    return d.innerHTML;
  }

  function closeActiveWindow() {
    if (activeWinId) closeWindow(activeWinId);
  }

  // ---- Public API ----
  return {
    init,
    openFile,
    openURL,
    saveActiveTable,
    saveActiveTableAs,
    newTable,
    closeActiveWindow,
    executeQuery,
    clearConsole,
    layoutTileH,
    layoutTileV,
    layoutGrid,
    layoutCascade,
    minimizeAll,
    restoreAll,
  };
})();

document.addEventListener('DOMContentLoaded', app.init);
