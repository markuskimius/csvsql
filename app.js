// ============================================================
// CSVSQL - CSV Database Application
// ============================================================

const app = (() => {
  let windows = [];
  let nextWinId = 1;
  let nextZIndex = 100;
  let activeWinId = null;
  let tables = {};  // tableName -> { columns, rows, filename, modified }

  // Virtual scrolling constants
  const ROW_HEIGHT = 26;
  const OVERSCAN = 10;

  // Debounced sync timers
  const syncTimers = {};

  // Sort optimization
  const collator = new Intl.Collator(undefined, { sensitivity: 'base' });

  // ---- Init ----
  function init() {
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
            { description: 'CSV files', accept: { 'text/csv': ['.csv', '.tsv', '.txt'] } },
            { description: 'Excel files', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'], 'application/vnd.ms-excel': ['.xls'] } },
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

  function openFileByType(file, fileHandle) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'xlsx' || ext === 'xls') {
      loadExcelFile(file);
    } else {
      loadCSVFile(file, fileHandle);
    }
  }

  function loadCSVFile(file, fileHandle) {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      dynamicTyping: false,
      complete(results) {
        const name = sanitizeTableName(file.name.replace(/\.[^.]+$/, ''));
        const uniqueName = getUniqueTableName(name);
        const columns = results.meta.fields || [];
        const rows = results.data.map((row, i) => ({ _rownum: i + 1, ...row }));
        tables[uniqueName] = { columns, rows, filename: file.name, modified: false, fileHandle: fileHandle || null };
        registerAlaSQL(uniqueName);
        createTableWindow(uniqueName);
      },
      error(err) {
        setStatus(`Error parsing ${file.name}: ${err.message}`, 'error');
      }
    });
  }

  function loadExcelFile(file) {
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
          const columns = Object.keys(jsonData[0]);
          const rows = jsonData.map((row, i) => {
            const r = { _rownum: i + 1 };
            columns.forEach(c => { r[c] = row[c] != null ? String(row[c]) : ''; });
            return r;
          });
          tables[uniqueName] = { columns, rows, filename: file.name, modified: false };
          registerAlaSQL(uniqueName);
          createTableWindow(uniqueName);
        });
        setStatus(`Opened ${file.name} (${workbook.SheetNames.length} sheet(s))`, 'success');
      } catch (err) {
        setStatus(`Error reading ${file.name}: ${err.message}`, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function sanitizeTableName(name) {
    return name.replace(/[^a-zA-Z0-9_]/g, '_').replace(/^[0-9]/, '_$&');
  }

  function getUniqueTableName(base) {
    let name = base;
    let i = 2;
    while (tables[name]) { name = base + '_' + i; i++; }
    return name;
  }

  function registerAlaSQL(tableName) {
    const t = tables[tableName];
    try { alasql(`DROP TABLE IF EXISTS [${tableName}]`); } catch (e) {}
    if (t.columns.length === 0) {
      alasql(`CREATE TABLE [${tableName}]`);
    } else {
      const colDefs = t.columns.map(c => `[${c}] STRING`).join(', ');
      alasql(`CREATE TABLE [${tableName}] (${colDefs})`);
    }
    if (!alasql.tables[tableName]) return;
    const insertRows = t.rows.map(r => {
      const obj = {};
      t.columns.forEach(c => { obj[c] = r[c] ?? ''; });
      return obj;
    });
    alasql.tables[tableName].data = insertRows;
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
    const csv = serializeCSV(tableName);
    const writable = await handle.createWritable();
    await writable.write(csv);
    await writable.close();
    const filename = handle.name;
    t.modified = false;
    t.filename = filename;
    updateWindowTitle(tableName);
    setStatus(`Saved ${filename}`, 'success');
  }

  function downloadCSV(tableName, filename) {
    flushAllSyncs();
    const t = tables[tableName];
    if (!t) return;
    const csv = serializeCSV(tableName);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
    t.modified = false;
    t.filename = filename;
    updateWindowTitle(tableName);
    setStatus(`Saved ${filename}`, 'success');
  }

  function serializeCSV(tableName) {
    const t = tables[tableName];
    const data = t.rows.map(r => {
      const obj = {};
      t.columns.forEach(c => { obj[c] = r[c]; });
      return obj;
    });
    return Papa.unparse(data, { columns: t.columns });
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
        registerAlaSQL(uniqueName);
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
      sortCol: null,
      sortDir: 'asc',
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
      try { alasql(`DROP TABLE IF EXISTS [${win.tableName}]`); } catch (e) {}
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

      // Re-register in AlaSQL under new name
      try { alasql(`DROP TABLE IF EXISTS [${oldName}]`); } catch (_) {}
      registerAlaSQL(uniqueName);

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
      <input type="text" class="filter-input" placeholder="Search all columns..." value="${escHtml(win.filterText)}">
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

  function buildTableHTML(win, container, tableData) {
    const { columns, rows } = tableData;
    let displayRows = [...rows];

    // Filter
    if (win.filterText) {
      const q = win.filterText.toLowerCase();
      displayRows = displayRows.filter(row =>
        columns.some(c => String(row[c] ?? '').toLowerCase().includes(q))
      );
    }

    // Sort
    if (win.sortCol) {
      const col = win.sortCol;
      const dir = win.sortDir === 'asc' ? 1 : -1;
      displayRows.sort((a, b) => {
        const va = a[col] ?? '', vb = b[col] ?? '';
        const na = Number(va), nb = Number(vb);
        if (!isNaN(na) && !isNaN(nb) && va !== '' && vb !== '') return (na - nb) * dir;
        return collator.compare(String(va), String(vb)) * dir;
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
      if (win.sortCol === col) {
        const arrow = document.createElement('span');
        arrow.className = 'sort-arrow';
        arrow.textContent = win.sortDir === 'asc' ? '\u25B2' : '\u25BC';
        th.appendChild(arrow);
      }
      th.addEventListener('click', () => {
        if (win.sortCol === col) {
          win.sortDir = win.sortDir === 'asc' ? 'desc' : 'asc';
        } else {
          win.sortCol = col;
          win.sortDir = 'asc';
        }
        rebuildTable(win);
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
        debouncedSyncToAlaSQL(win.tableName);
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
    debouncedSyncToAlaSQL(tableName);
  }

  function deleteRow(tableName, rownum, win) {
    const t = tables[tableName];
    t.rows = t.rows.filter(r => r._rownum !== rownum);
    // Renumber
    t.rows.forEach((r, i) => { r._rownum = i + 1; });
    markModified(tableName);
    debouncedSyncToAlaSQL(tableName);
    rebuildTable(win);
  }

  function addColumn(tableName, colName) {
    const t = tables[tableName];
    if (t.columns.includes(colName)) return;
    t.columns.push(colName);
    t.rows.forEach(r => { r[colName] = ''; });
    markModified(tableName);
    registerAlaSQL(tableName);
  }

  function markModified(tableName) {
    const t = tables[tableName];
    if (t) t.modified = true;
    updateWindowTitle(tableName);
  }

  function syncToAlaSQL(tableName) {
    const t = tables[tableName];
    if (!t) return;
    try {
      const data = t.rows.map(r => {
        const obj = {};
        t.columns.forEach(c => { obj[c] = r[c] ?? ''; });
        return obj;
      });
      alasql.tables[tableName].data = data;
    } catch (e) {}
  }

  function debouncedSyncToAlaSQL(tableName, delay = 500) {
    clearTimeout(syncTimers[tableName]);
    syncTimers[tableName] = setTimeout(() => {
      delete syncTimers[tableName];
      syncToAlaSQL(tableName);
    }, delay);
  }

  function flushAllSyncs() {
    for (const name of Object.keys(syncTimers)) {
      clearTimeout(syncTimers[name]);
      delete syncTimers[name];
      syncToAlaSQL(name);
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
      syncToAlaSQL(tableName);
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

  function executeQuery() {
    flushAllSyncs();
    const input = document.getElementById('sql-input');
    const sql = input.value.trim();
    if (!sql) return;

    const t0 = performance.now();

    // Handle SELECT ... INTO ... by running the SELECT and creating the table from results
    const intoInfo = extractIntoClause(sql);
    if (intoInfo) {
      try {
        const rows = alasql(intoInfo.selectSQL);
        const elapsed = performance.now() - t0;
        if (!Array.isArray(rows) || rows.length === 0 || typeof rows[0] !== 'object') {
          setStatus('INTO query returned no rows', 'error');
          return;
        }
        const name = sanitizeTableName(intoInfo.targetName);
        const uniqueName = tables[name] ? getUniqueTableName(name) : name;
        const columns = Object.keys(rows[0]);
        const tableRows = rows.map((r, i) => {
          const row = { _rownum: i + 1 };
          columns.forEach(c => { row[c] = r[c] != null ? String(r[c]) : ''; });
          return row;
        });
        tables[uniqueName] = { columns, rows: tableRows, filename: null, modified: true, fileHandle: null };
        registerAlaSQL(uniqueName);
        createTableWindow(uniqueName);
        setStatus(`Created table "${uniqueName}" with ${rows.length} row(s) in ${formatElapsed(elapsed)}`, 'success');
      } catch (e) {
        setStatus(`Error: ${e.message}`, 'error');
      }
      return;
    }

    // Snapshot existing AlaSQL tables before query
    const tablesBefore = new Set(Object.keys(alasql.tables));

    try {
      const result = alasql(sql);
      const elapsed = performance.now() - t0;

      // Detect new tables created by CREATE TABLE etc.
      const newAlaTables = Object.keys(alasql.tables).filter(n => !tablesBefore.has(n));

      const createMatch = sql.match(/\bCREATE\s+TABLE\s+(?:IF\s+NOT\s+EXISTS\s+)?\[?([^\]\s,(;]+)\]?/i);
      if (createMatch) {
        const createName = createMatch[1];
        if (!newAlaTables.includes(createName) && alasql.tables[createName] && !tables[createName]) {
          newAlaTables.push(createName);
        }
      }

      if (newAlaTables.length > 0) {
        importNewAlaTables(newAlaTables);
        setStatus(`Created table(s): ${newAlaTables.join(', ')} in ${formatElapsed(elapsed)}`, 'success');
      } else if (Array.isArray(result) && result.length > 0 && typeof result[0] === 'object') {
        showQueryResult(sql, result);
        setStatus(`Query returned ${result.length} row(s) in ${formatElapsed(elapsed)}`, 'success');
      } else if (Array.isArray(result)) {
        setStatus(`Query returned: ${JSON.stringify(result)} in ${formatElapsed(elapsed)}`, 'success');
        refreshAllTableWindows();
      } else {
        setStatus(`Result: ${JSON.stringify(result)} in ${formatElapsed(elapsed)}`, 'success');
        refreshAllTableWindows();
      }
    } catch (e) {
      const elapsed = performance.now() - t0;
      setStatus(`Error: ${e.message} (${formatElapsed(elapsed)})`, 'error');
    }
  }

  function importNewAlaTables(tableNames) {
    tableNames.forEach(name => {
      const alaTable = alasql.tables[name];
      if (!alaTable) return;
      const data = alaTable.data || [];
      const colDefs = alaTable.columns || [];
      const columns = data.length > 0
        ? Object.keys(data[0])
        : colDefs.map(c => c.columnid);
      const rows = data.map((r, i) => {
        const row = { _rownum: i + 1 };
        columns.forEach(c => { row[c] = r[c] != null ? String(r[c]) : ''; });
        return row;
      });
      tables[name] = { columns, rows, filename: null, modified: true, fileHandle: null };
      // Keep AlaSQL in sync with string values
      registerAlaSQL(name);
      createTableWindow(name);
    });
  }

  function showQueryResult(sql, resultRows) {
    const columns = Object.keys(resultRows[0]);
    const tableName = '_query_' + nextWinId;

    // Store as a table so it can be saved
    const rows = resultRows.map((r, i) => ({ _rownum: i + 1, ...r }));
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
        // Re-sync from alasql data
        const alaData = alasql.tables[w.tableName]?.data;
        if (alaData) {
          const t = tables[w.tableName];
          t.rows = alaData.map((r, i) => ({ _rownum: i + 1, ...r }));
        }
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
