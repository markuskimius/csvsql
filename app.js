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
  let _activeConsoleTab = 'sql';

  // Touch gesture state (shared across tables / windows)
  let _activeAutoFilter = null;            // { win, col, el } — currently open autofilter dropdown
  let _activePluginPopover = null;         // modal overlay element for plugin about dialog
  let _touchHeaderDrag = null;             // header 1.5-tap drag (→ reorder column)
  let _touchCellDrag = null;               // cell second-tap interaction (→ double-tap edit or 1.5-tap multi-select)
  let _touchWinDrag = null;                // titlebar 1.5-tap drag (→ move window)
  let _lastHeaderTap = { tableName: null, col: null, time: 0 };
  let _lastCellTap = { winId: null, rownum: null, col: null, time: 0 };
  let _lastTitleTap = { winId: null, time: 0 };
  const DOUBLE_TAP_WINDOW_MS = 500;        // tap 1 → tap 2 window for double-tap / 1.5-tap

  // Virtual scrolling constants
  const ROW_HEIGHT = 26;
  const OVERSCAN = 10;

  // Debounced sync timers
  const syncTimers = {};

  // Plugin state
  let plugins = [];
  let _columnTransformCache = {};

  // Sort optimization
  const collator = new Intl.Collator(undefined, { sensitivity: 'base' });

  // Zip group counter — tables from the same zip share a zipGroupId
  let nextZipGroupId = 1;

  // Excel group counter — tables from the same workbook share an excelGroupId
  let nextExcelGroupId = 1;

  function registerDBFunctions() {
    db.create_function('regexp', (pattern, value) => {
      try { return new RegExp(pattern, 'i').test(value) ? 1 : 0; } catch (_) { return 0; }
    });
  }

  // ---- Init ----
  async function init() {
    const SQL = await initSqlJs({
      locateFile: file => `lib/${file}`
    });
    db = new SQL.Database();
    registerDBFunctions();
    setupConsoleResize();
    setupFileInput();
    setupDragAndDrop();
    setupKeyboard();
    setupSQLHighlight();
    setupMenuClose();
    setupAI();
    setupBrowserResize();
    fixShortcutLabels();
    loadPersistedPlugins();
    updatePluginMenu();
    window._appReady = true;
  }

  function fixShortcutLabels() {
    if (navigator.platform.includes('Mac') || navigator.userAgent.includes('Mac')) {
      document.querySelectorAll('.shortcut').forEach(el => {
        el.textContent = el.textContent.replace('Ctrl+', '\u2318').replace('Shift+', '\u21E7');
      });
    }
  }

  // ---- File Menu ----
  // Track whether Shift is held — when true, files open without headers (columns become A, B, C, ...)
  let _shiftOpen = false;

  function setupFileInput() {
    document.getElementById('file-input').addEventListener('change', (e) => {
      const hasHeader = !_shiftOpen;
      for (const file of e.target.files) {
        openFileByType(file, null, undefined, hasHeader);
      }
      e.target.value = '';
      _shiftOpen = false;
    });
  }

  function isImageFile(file) {
    return file && (file.type.startsWith('image/') || /\.(png|jpe?g|gif|webp|svg|bmp)$/i.test(file.name));
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
        overlay.classList.remove('visible', 'shift');
      }
    });

    // Highlight AI panel when hovering over it during a drag
    const aiBody = document.getElementById('ai-body');
    let aiDragCounter = 0;
    aiBody.addEventListener('dragenter', () => { aiDragCounter++; aiBody.classList.add('ai-drag-over'); });
    aiBody.addEventListener('dragleave', () => { aiDragCounter--; if (aiDragCounter <= 0) { aiDragCounter = 0; aiBody.classList.remove('ai-drag-over'); } });

    document.addEventListener('dragover', (e) => {
      e.preventDefault();
      overlay.classList.toggle('shift', e.shiftKey);
    });

    document.addEventListener('drop', async (e) => {
      e.preventDefault();
      dragCounter = 0;
      aiDragCounter = 0;
      overlay.classList.remove('visible', 'shift');
      aiBody.classList.remove('ai-drag-over');

      // If dropped on AI panel and files are images, handle as AI image upload
      if (aiBody && aiBody.style.display !== 'none' && aiBody.contains(e.target)) {
        const files = [...e.dataTransfer.files];
        const images = files.filter(isImageFile);
        if (images.length > 0) {
          for (const img of images) addAIImage(img);
          return;
        }
      }

      closeMenus();
      const hasHeader = !e.shiftKey;
      const entries = [...e.dataTransfer.items].map(item => ({
        file: item.getAsFile(),
        handlePromise: item.getAsFileSystemHandle ? item.getAsFileSystemHandle().catch(() => null) : Promise.resolve(null),
      }));
      for (const entry of entries) {
        const handle = await entry.handlePromise;
        if (entry.file) {
          if (entry.file.name.toLowerCase().endsWith('.json')) {
            loadPluginFile(entry.file);
          } else {
            openFileByType(entry.file, handle, undefined, hasHeader);
          }
        }
      }
    });
  }

  async function openFile(shiftKey) {
    const hasHeader = !shiftKey;
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
          openFileByType(file, handle, undefined, hasHeader);
        }
      } catch (e) {
        if (e.name !== 'AbortError') setStatus(`Error opening file: ${e.message}`, 'error');
      }
    } else {
      _shiftOpen = shiftKey;
      document.getElementById('file-input').click();
    }
  }

  function openURL(noHeader) {
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
        openFileByType(file, null, undefined, !noHeader);
      } catch (e) {
        setStatus(`Error fetching URL: ${e.message}`, 'error');
      }
    });
  }

  // Compression extensions and the data file extensions they may wrap
  const COMPRESSION_EXTS = new Set(['gz', 'zip', 'bz2', 'xz', 'rar', '7z', 'zst']);
  const DATA_EXTS = new Set(['csv', 'tsv', 'psv', 'txt', 'xlsx', 'xls']);

  function openFileByType(file, fileHandle, compression, hasHeader) {
    if (hasHeader === undefined) hasHeader = true;
    const ext = file.name.split('.').pop().toLowerCase();
    if (COMPRESSION_EXTS.has(ext)) {
      decompressAndOpen(file, fileHandle, hasHeader);
    } else if (ext === 'xlsx' || ext === 'xls') {
      loadExcelFile(file, fileHandle, compression, hasHeader);
    } else {
      loadDelimitedFile(file, compression ? compression.fileHandle : fileHandle, compression, hasHeader);
    }
  }

  async function decompressAndOpen(file, fileHandle, hasHeader) {
    const ext = file.name.split('.').pop().toLowerCase();
    setStatus(`Decompressing ${file.name}...`, 'working');
    try {
      if (ext === 'gz') {
        await decompressGzip(file, fileHandle, hasHeader);
      } else if (ext === 'zip') {
        await decompressZip(file, fileHandle, hasHeader);
      } else {
        setStatus(`Unsupported compression format: .${ext} — please decompress the file first and open the decompressed file`, 'error');
      }
    } catch (e) {
      setStatus(`Error decompressing ${file.name}: ${e.message}`, 'error');
    }
  }

  async function decompressGzip(file, fileHandle, hasHeader) {
    const ds = new DecompressionStream('gzip');
    const decompressed = file.stream().pipeThrough(ds);
    const blob = await new Response(decompressed).blob();
    // Inner filename: strip .gz
    const innerName = file.name.replace(/\.gz$/i, '') || 'decompressed.csv';
    const innerFile = new File([blob], innerName, { type: 'application/octet-stream' });
    openFileByType(innerFile, null, { type: 'gz', compressedFilename: file.name, fileHandle: fileHandle || null }, hasHeader);
  }

  async function decompressZip(file, fileHandle, hasHeader) {
    const zip = await JSZip.loadAsync(file);
    const zipGroupId = nextZipGroupId++;

    // Collect recognized data files
    const dataFiles = [];
    for (const [name, entry] of Object.entries(zip.files)) {
      if (entry.dir) continue;
      const innerExt = name.split('.').pop().toLowerCase();
      if (DATA_EXTS.has(innerExt) || COMPRESSION_EXTS.has(innerExt)) {
        dataFiles.push({ name, entry });
      }
    }

    if (dataFiles.length === 0) {
      const firstEntry = Object.values(zip.files).find(e => !e.dir);
      if (firstEntry) {
        dataFiles.push({ name: firstEntry.name || 'data.csv', entry: firstEntry });
      } else {
        setStatus('ZIP archive is empty', 'error');
        return;
      }
    }

    const zipOriginalCount = dataFiles.length;
    for (const { name, entry } of dataFiles) {
      const blob = await entry.async('blob');
      const innerFile = new File([blob], name, { type: 'application/octet-stream' });
      openFileByType(innerFile, null, { type: 'zip', compressedFilename: file.name, fileHandle: fileHandle || null, zipGroupId, innerName: name, zipOriginalCount }, hasHeader);
    }
  }

  // Generate Excel-style column name: 0→A, 1→B, ..., 25→Z, 26→AA, 27→AB, ...
  function excelColName(index) {
    let name = '';
    let n = index;
    do {
      name = String.fromCharCode(65 + (n % 26)) + name;
      n = Math.floor(n / 26) - 1;
    } while (n >= 0);
    return name;
  }

  function delimiterForExt(filename) {
    const ext = filename.split('.').pop().toLowerCase();
    if (ext === 'tsv') return '\t';
    if (ext === 'psv') return '|';
    return undefined; // let Papa auto-detect (handles csv, txt)
  }

  function loadDelimitedFile(file, fileHandle, compression, hasHeader) {
    setStatus(`Loading ${file.name}...`, 'working');
    const t0 = performance.now();
    const delimiter = delimiterForExt(file.name);
    const allRows = [];
    let detectedDelimiter = delimiter || ',';
    let rawColumns = null;
    let columns = null;
    Papa.parse(file, {
      header: hasHeader,
      skipEmptyLines: true,
      dynamicTyping: false,
      delimiter,
      chunk(results) {
        if (hasHeader) {
          if (!rawColumns) {
            rawColumns = results.meta.fields || [];
            columns = sanitizeColumns(rawColumns);
            detectedDelimiter = delimiter || results.meta.delimiter || ',';
          }
          for (const row of results.data) {
            const r = { _rownum: allRows.length + 1 };
            rawColumns.forEach((raw, j) => { r[columns[j]] = row[raw] ?? ''; });
            allRows.push(r);
          }
        } else {
          if (!columns) {
            const numCols = results.data[0] ? results.data[0].length : 0;
            columns = Array.from({ length: numCols }, (_, i) => excelColName(i));
            detectedDelimiter = delimiter || results.meta.delimiter || ',';
          }
          for (const row of results.data) {
            const r = { _rownum: allRows.length + 1 };
            columns.forEach((col, j) => { r[col] = row[j] ?? ''; });
            allRows.push(r);
          }
        }
        setStatus(`Loading ${file.name}... ${allRows.length.toLocaleString()} rows`, 'working');
      },
      async complete() {
        if (!columns) {
          columns = [];
          rawColumns = [];
        }
        const name = sanitizeTableName(file.name.replace(/\.[^.]+$/, ''));
        const uniqueName = getUniqueTableName(name);
        tables[uniqueName] = { columns, rows: allRows, filename: file.name, modified: false, fileHandle: fileHandle || null, delimiter: detectedDelimiter, compression: compression || null };
        setStatus(`Indexing ${file.name}... 0 / ${allRows.length.toLocaleString()} rows`, 'working');
        await new Promise(r => setTimeout(r, 0));
        await registerTable(uniqueName);
        createTableWindow(uniqueName);
        const elapsed = performance.now() - t0;
        setStatus(`Opened ${file.name} (${allRows.length.toLocaleString()} rows) in ${formatElapsed(elapsed)}`, 'success');
      },
      error(err) {
        setStatus(`Error parsing ${file.name}: ${err.message}`, 'error');
      }
    });
  }

  function loadExcelFile(file, fileHandle, compression, hasHeader) {
    setStatus(`Loading ${file.name}...`, 'working');
    const t0 = performance.now();
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const workbook = XLSX.read(e.target.result, { type: 'array' });
        const excelGroupId = nextExcelGroupId++;
        const nonEmptySheets = workbook.SheetNames.filter(sn => {
          const s = workbook.Sheets[sn];
          return XLSX.utils.sheet_to_json(s, { header: 1, defval: '' }).some(r => r.length > 0);
        });
        const excelOriginalCount = nonEmptySheets.length;
        for (const sheetName of nonEmptySheets) {
          const sheet = workbook.Sheets[sheetName];
          let columns, rows;
          if (hasHeader) {
            const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: '' });
            if (jsonData.length === 0) continue;
            const rawColumns = Object.keys(jsonData[0]);
            columns = sanitizeColumns(rawColumns);
            rows = jsonData.map((row, i) => {
              const r = { _rownum: i + 1 };
              rawColumns.forEach((raw, j) => { r[columns[j]] = row[raw] != null ? String(row[raw]) : ''; });
              return r;
            });
          } else {
            const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
            if (rawData.length === 0) continue;
            const numCols = Math.max(...rawData.map(r => r.length));
            columns = Array.from({ length: numCols }, (_, i) => excelColName(i));
            rows = rawData.map((row, i) => {
              const r = { _rownum: i + 1 };
              columns.forEach((col, j) => { r[col] = row[j] != null ? String(row[j]) : ''; });
              return r;
            });
          }
          const name = sanitizeTableName(sheetName);
          const uniqueName = getUniqueTableName(name);
          const excelInfo = { excelGroupId, sheetName, excelOriginalCount, excelFilename: file.name, fileHandle: fileHandle || null };
          tables[uniqueName] = { columns, rows, filename: file.name, modified: false, compression: compression || null, excel: excelInfo };
          await registerTable(uniqueName);
          createTableWindow(uniqueName);
        }
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

  async function registerTable(tableName) {
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
    const total = t.rows.length;
    const BATCH = 50000;
    for (let i = 0; i < total; i += BATCH) {
      db.run('BEGIN TRANSACTION');
      const stmt = db.prepare(`INSERT INTO [${tableName}] VALUES (${placeholders})`);
      const end = Math.min(i + BATCH, total);
      for (let j = i; j < end; j++) {
        stmt.run(t.columns.map(c => t.rows[j][c] ?? ''));
      }
      stmt.free();
      db.run('COMMIT');
      if (end < total) {
        setStatus(`Indexing ${t.filename || tableName}... ${end.toLocaleString()} / ${total.toLocaleString()} rows`, 'working');
        await new Promise(r => setTimeout(r, 0));
      }
    }
    rebuildTransformCacheForTable(tableName);
  }

  async function saveActiveTable() {
    flushAllSyncs();
    const win = getActiveDataWindow();
    if (!win) return;
    const t = tables[win.tableName];
    if (!t) return;

    // Zip group save: re-pack all tables from the same zip
    if (t.compression && t.compression.type === 'zip' && t.compression.zipGroupId) {
      await saveZipGroup(t.compression);
      return;
    }

    // Excel group save: re-pack all sheets from the same workbook
    if (t.excel && t.excel.excelGroupId) {
      await saveExcelGroup(t.excel);
      return;
    }

    const handle = t.fileHandle || (t.compression && t.compression.fileHandle);
    if (handle) {
      await writeToHandle(win.tableName, handle);
    } else if (t.filename) {
      await downloadCSV(win.tableName, t.filename);
    } else {
      await saveActiveTableAs();
    }
  }

  function getZipGroupTables(zipGroupId) {
    const result = [];
    for (const [name, t] of Object.entries(tables)) {
      if (t.compression && t.compression.zipGroupId === zipGroupId) {
        result.push({ tableName: name, table: t });
      }
    }
    return result;
  }

  async function saveZipGroup(compression) {
    const { zipGroupId, compressedFilename } = compression;
    const groupTables = getZipGroupTables(zipGroupId);

    // Check if any tables from the group have been closed
    // We detect this by comparing against the original inner names
    const allInnerNames = new Set();
    const presentInnerNames = new Set();
    for (const { table: tbl } of groupTables) {
      if (tbl.compression && tbl.compression.innerName) {
        presentInnerNames.add(tbl.compression.innerName);
      }
    }

    // If no tables remain, fall through to Save As
    if (groupTables.length === 0) {
      await saveActiveTableAs();
      return;
    }

    // Check if any table was closed by comparing with original group
    // We track original count: if any table had a peer that's now gone
    // we can detect by checking if modified tables exist without the full set
    // Simpler: store original count on the compression object
    const originalCount = compression.zipOriginalCount;
    if (originalCount && groupTables.length < originalCount) {
      const missing = originalCount - groupTables.length;
      setStatus(`Warning: ${missing} table(s) from ${compressedFilename} no longer open — using Save As to avoid overwriting`, 'error');
      await saveActiveTableAs();
      return;
    }

    const handle = compression.fileHandle;
    const t0 = performance.now();
    setStatus(`Saving ${compressedFilename}...`, 'working');
    await new Promise(r => setTimeout(r, 0));

    const zip = new JSZip();
    for (const { tableName, table: tbl } of groupTables) {
      const innerName = tbl.compression.innerName || (tableName + '.csv');
      let blob;
      if (isExcelFilename(innerName)) {
        blob = serializeExcel(tbl);
      } else {
        const parts = [serializeHeader(tbl)];
        const CHUNK = 50000;
        for (let i = 0; i < tbl.rows.length; i += CHUNK) {
          parts.push(serializeChunk(tbl, i, Math.min(i + CHUNK, tbl.rows.length)));
        }
        blob = new Blob(parts, { type: 'text/csv;charset=utf-8;' });
      }
      zip.file(innerName, blob);
      setStatus(`Saving ${compressedFilename}... packed ${innerName}`, 'working');
    }

    const zipBlob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });

    if (handle) {
      const writable = await handle.createWritable();
      await writable.write(zipBlob);
      await writable.close();
    } else {
      const a = document.createElement('a');
      a.href = URL.createObjectURL(zipBlob);
      a.download = compressedFilename;
      a.click();
      URL.revokeObjectURL(a.href);
    }

    // Mark all tables as saved
    for (const { tableName, table: tbl } of groupTables) {
      tbl.modified = false;
      updateWindowTitle(tableName);
    }

    const totalRows = groupTables.reduce((sum, { table: tbl }) => sum + tbl.rows.length, 0);
    const elapsed = performance.now() - t0;
    setStatus(`Saved ${compressedFilename} (${groupTables.length} file(s), ${totalRows.toLocaleString()} rows) in ${formatElapsed(elapsed)}`, 'success');
  }

  function getExcelGroupTables(excelGroupId) {
    const result = [];
    for (const [name, t] of Object.entries(tables)) {
      if (t.excel && t.excel.excelGroupId === excelGroupId) {
        result.push({ tableName: name, table: t });
      }
    }
    return result;
  }

  async function saveExcelGroup(excelInfo) {
    const { excelGroupId, excelFilename, excelOriginalCount } = excelInfo;
    const groupTables = getExcelGroupTables(excelGroupId);

    if (groupTables.length === 0) {
      await saveActiveTableAs();
      return;
    }

    if (excelOriginalCount && groupTables.length < excelOriginalCount) {
      const missing = excelOriginalCount - groupTables.length;
      setStatus(`Warning: ${missing} sheet(s) from ${excelFilename} no longer open — using Save As to avoid overwriting`, 'error');
      await saveActiveTableAs();
      return;
    }

    const handle = excelInfo.fileHandle;
    const t0 = performance.now();
    setStatus(`Saving ${excelFilename}...`, 'working');
    await new Promise(r => setTimeout(r, 0));

    const wb = XLSX.utils.book_new();
    for (const { table: tbl } of groupTables) {
      const sheetName = tbl.excel.sheetName || 'Sheet1';
      const data = tbl.rows.map(row => {
        const obj = {};
        tbl.columns.forEach(c => { obj[c] = row[c] ?? ''; });
        return obj;
      });
      const ws = XLSX.utils.json_to_sheet(data, { header: tbl.columns });
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    }
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

    if (handle) {
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
    } else {
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = excelFilename;
      a.click();
      URL.revokeObjectURL(a.href);
    }

    for (const { tableName, table: tbl } of groupTables) {
      tbl.modified = false;
      updateWindowTitle(tableName);
    }

    const totalRows = groupTables.reduce((sum, { table: tbl }) => sum + tbl.rows.length, 0);
    const elapsed = performance.now() - t0;
    setStatus(`Saved ${excelFilename} (${groupTables.length} sheet(s), ${totalRows.toLocaleString()} rows) in ${formatElapsed(elapsed)}`, 'success');
  }

  async function saveActiveTableAs() {
    flushAllSyncs();
    const win = getActiveDataWindow();
    if (!win) return;
    const t = tables[win.tableName];
    if (!t) return;
    const baseFilename = t.filename || win.tableName + '.csv';
    const suggestedName = compressedFilename(baseFilename, t.compression);
    if (window.showSaveFilePicker) {
      try {
        const types = [
          { description: 'CSV files', accept: { 'text/csv': ['.csv'] } },
          { description: 'TSV files', accept: { 'text/tab-separated-values': ['.tsv'] } },
          { description: 'PSV files', accept: { 'text/plain': ['.psv'] } },
          { description: 'Excel files', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } },
          { description: 'Gzip compressed', accept: { 'application/gzip': ['.gz'] } },
          { description: 'ZIP compressed', accept: { 'application/zip': ['.zip'] } },
        ];
        const handle = await showSaveFilePicker({ suggestedName, types });
        // Detect compression from chosen filename
        const chosenName = handle.name;
        if (chosenName.endsWith('.gz')) {
          t.compression = { type: 'gz', compressedFilename: chosenName };
        } else if (chosenName.endsWith('.zip')) {
          t.compression = { type: 'zip', compressedFilename: chosenName };
        } else {
          t.compression = null;
        }
        await writeToHandle(win.tableName, handle);
        t.fileHandle = handle;
      } catch (e) {
        if (e.name !== 'AbortError') setStatus(`Error saving: ${e.message}`, 'error');
      }
    } else {
      showPrompt('Save As', 'Filename:', suggestedName, (newName) => {
        if (!newName) return;
        // Detect compression from typed filename
        if (newName.endsWith('.gz')) {
          t.compression = { type: 'gz', compressedFilename: newName };
        } else if (newName.endsWith('.zip')) {
          t.compression = { type: 'zip', compressedFilename: newName };
        } else {
          t.compression = null;
        }
        downloadCSV(win.tableName, newName.replace(/\.(gz|zip)$/i, '') || newName);
      });
    }
  }

  async function writeToHandle(tableName, handle) {
    const t = tables[tableName];
    if (!t) return;
    const saveName = handle.name || t.filename || tableName;
    const total = t.rows.length;
    setStatus(`Saving ${saveName}... 0 / ${total.toLocaleString()} rows`, 'working');
    const t0 = performance.now();
    const writable = await handle.createWritable();
    if (isExcelFilename(saveName)) {
      const blob = serializeExcel(t);
      await writable.write(blob);
    } else if (t.compression) {
      const parts = [serializeHeader(t)];
      const CHUNK = 50000;
      for (let i = 0; i < total; i += CHUNK) {
        const end = Math.min(i + CHUNK, total);
        parts.push(serializeChunk(t, i, end));
        if (end < total) {
          setStatus(`Saving ${saveName}... ${end.toLocaleString()} / ${total.toLocaleString()} rows`, 'working');
          await new Promise(r => setTimeout(r, 0));
        }
      }
      let blob = new Blob(parts, { type: 'text/csv;charset=utf-8;' });
      setStatus(`Compressing ${saveName}...`, 'working');
      blob = await compressBlob(blob, t.compression);
      await writable.write(blob);
    } else {
      await writable.write(serializeHeader(t));
      const CHUNK = 50000;
      for (let i = 0; i < total; i += CHUNK) {
        const end = Math.min(i + CHUNK, total);
        const chunk = serializeChunk(t, i, end);
        await writable.write(chunk);
        setStatus(`Saving ${saveName}... ${end.toLocaleString()} / ${total.toLocaleString()} rows`, 'working');
      }
    }
    await writable.close();
    const filename = handle.name;
    t.delimiter = delimiterForExt(filename) || t.delimiter || ',';
    t.modified = false;
    t.filename = filename;
    updateWindowTitle(tableName);
    const elapsed = performance.now() - t0;
    setStatus(`Saved ${saveName} (${total.toLocaleString()} rows) in ${formatElapsed(elapsed)}`, 'success');
  }

  async function downloadCSV(tableName, filename) {
    flushAllSyncs();
    const t = tables[tableName];
    if (!t) return;
    const saveName = compressedFilename(filename, t.compression);
    const total = t.rows.length;
    setStatus(`Saving ${saveName}... 0 / ${total.toLocaleString()} rows`, 'working');
    await new Promise(r => setTimeout(r, 0));
    const t0 = performance.now();
    let blob;
    if (isExcelFilename(filename)) {
      blob = serializeExcel(t);
    } else {
      const parts = [serializeHeader(t)];
      const CHUNK = 50000;
      for (let i = 0; i < total; i += CHUNK) {
        const end = Math.min(i + CHUNK, total);
        parts.push(serializeChunk(t, i, end));
        if (end < total) {
          setStatus(`Saving ${saveName}... ${end.toLocaleString()} / ${total.toLocaleString()} rows`, 'working');
          await new Promise(r => setTimeout(r, 0));
        }
      }
      blob = new Blob(parts, { type: 'text/csv;charset=utf-8;' });
      if (t.compression) {
        setStatus(`Compressing ${saveName}...`, 'working');
        blob = await compressBlob(blob, t.compression);
      }
    }
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = saveName;
    a.click();
    URL.revokeObjectURL(a.href);
    t.modified = false;
    t.filename = filename;
    updateWindowTitle(tableName);
    const elapsed = performance.now() - t0;
    setStatus(`Saved ${saveName} (${total.toLocaleString()} rows) in ${formatElapsed(elapsed)}`, 'success');
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

  function isExcelFilename(filename) {
    return /\.xlsx?$/i.test(filename);
  }

  function serializeExcel(t) {
    const data = t.rows.map(row => {
      const obj = {};
      t.columns.forEach(c => { obj[c] = row[c] ?? ''; });
      return obj;
    });
    const ws = XLSX.utils.json_to_sheet(data, { header: t.columns });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    return new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  async function compressBlob(blob, compression) {
    if (!compression) return blob;
    if (compression.type === 'gz') {
      const cs = new CompressionStream('gzip');
      const compressed = blob.stream().pipeThrough(cs);
      return await new Response(compressed).blob();
    }
    if (compression.type === 'zip') {
      const zip = new JSZip();
      const innerName = compression.compressedFilename
        ? compression.compressedFilename.replace(/\.zip$/i, '')
        : 'data.csv';
      zip.file(innerName, blob);
      return await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
    }
    return blob;
  }

  function compressedFilename(filename, compression) {
    if (!compression) return filename;
    if (compression.type === 'gz' && !filename.endsWith('.gz')) return filename + '.gz';
    if (compression.type === 'zip' && !filename.endsWith('.zip')) return filename + '.zip';
    return filename;
  }

  function newTable() {
    showPrompt('New Table', 'Table name:', '', (name) => {
      if (!name) return;
      const safeName = sanitizeTableName(name);
      const uniqueName = getUniqueTableName(safeName);
      showPrompt('Columns', 'Column names (comma-separated):', 'id, name, value', async (colStr) => {
        if (!colStr) return;
        const columns = colStr.split(',').map(c => c.trim()).filter(Boolean);
        tables[uniqueName] = { columns, rows: [], filename: null, modified: true };
        await registerTable(uniqueName);
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
      <div class="win-statusbar"><span class="status-left"></span><span class="status-center"></span><span class="status-right"></span></div>
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
      columnFilters: {},  // { colName: Set of allowed string values }
      selectedCol: null, // column name currently highlighted (target for Ctrl+Arrow reorder)
      selectedCells: new Set(), // keys "rownum:colName" for cells highlighted by selection
      anchorCell: null, // { rownum, col } — fixed corner for Ctrl+Shift+Arrow extension
      _copyWithHeader: false,
      disabledTransforms: new Set(),
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
    if (win.tableName && tables[win.tableName]) {
      const t = tables[win.tableName];
      if (!win.isQuery && t.modified) {
        if (!confirm(`Table "${win.tableName}" has unsaved changes. Close anyway?`)) return;
      }
      try { db.run(`DROP TABLE IF EXISTS [${win.tableName}]`); } catch (e) {}
      delete _columnTransformCache[win.tableName];
      delete tables[win.tableName];
    }
    win.el.remove();
    windows.splice(idx, 1);
    if (activeWinId === id) {
      activeWinId = windows.length ? windows[windows.length - 1].id : null;
      if (activeWinId) focusWindow(activeWinId);
    }
    updateWindowsList();
    if (_activeConsoleTab === 'ai') populateTableSelect();
  }

  function closeAllWindows() {
    const ids = windows.map(w => w.id);
    for (const id of ids) closeWindow(id);
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
      const area = document.getElementById('window-area').getBoundingClientRect();
      const cx = Math.max(area.left, Math.min(e.clientX, area.right));
      const cy = Math.max(area.top, Math.min(e.clientY, area.bottom));
      const dx = cx - startX, dy = cy - startY;
      if (Math.abs(dx) < 3 && Math.abs(dy) < 3) return;
      const areaW = area.right - area.left, areaH = area.bottom - area.top;
      const winW = win.el.offsetWidth, winH = win.el.offsetHeight;
      win.el.style.left = Math.max(0, Math.min(origX + dx, areaW - winW)) + 'px';
      win.el.style.top = Math.max(0, Math.min(origY + dy, areaH - winH)) + 'px';
      if (win.maximized) win.maximized = false;
    });

    document.addEventListener('mouseup', () => {
      if (dragging) {
        dragging = false;
        titlebar.style.cursor = 'grab';
      }
    });

    // Touch: 1.5-tap (tap once, then tap-and-pan) to move the window. A single
    // tap is reserved for normal taps on titlebar children (focus, buttons).
    titlebar.addEventListener('touchstart', (e) => {
      if (e.touches.length !== 1) return;
      if (e.target.tagName === 'BUTTON') return;
      const touch = e.touches[0];
      const now = Date.now();
      const prev = _lastTitleTap;
      const isSecondTap = prev.winId === win.id &&
                          (now - prev.time) < DOUBLE_TAP_WINDOW_MS;
      if (!isSecondTap) {
        _lastTitleTap = { winId: win.id, time: now };
        return;
      }
      _lastTitleTap = { winId: null, time: 0 };
      focusWindow(win.id);
      _touchWinDrag = {
        win,
        startX: touch.clientX, startY: touch.clientY,
        origX: parseInt(win.el.style.left),
        origY: parseInt(win.el.style.top),
        dragging: false,
      };
    }, { passive: true });

    titlebar.addEventListener('touchmove', (e) => {
      if (!_touchWinDrag || _touchWinDrag.win !== win) return;
      const touch = e.touches[0];
      if (!touch) return;
      const dx = touch.clientX - _touchWinDrag.startX;
      const dy = touch.clientY - _touchWinDrag.startY;
      if (!_touchWinDrag.dragging && (Math.abs(dx) > 5 || Math.abs(dy) > 5)) {
        _touchWinDrag.dragging = true;
        titlebar.style.cursor = 'grabbing';
        if (navigator.vibrate) navigator.vibrate(10);
      }
      if (_touchWinDrag.dragging) {
        if (e.cancelable) e.preventDefault();
        const area = document.getElementById('window-area').getBoundingClientRect();
        const cx = Math.max(area.left, Math.min(touch.clientX, area.right));
        const cy = Math.max(area.top, Math.min(touch.clientY, area.bottom));
        const ax = cx - _touchWinDrag.startX;
        const ay = cy - _touchWinDrag.startY;
        const areaW = area.right - area.left, areaH = area.bottom - area.top;
        const winW = win.el.offsetWidth, winH = win.el.offsetHeight;
        win.el.style.left = Math.max(0, Math.min(_touchWinDrag.origX + ax, areaW - winW)) + 'px';
        win.el.style.top = Math.max(0, Math.min(_touchWinDrag.origY + ay, areaH - winH)) + 'px';
        if (win.maximized) win.maximized = false;
      }
    }, { passive: false });

    const endWinTouch = (e) => {
      if (!_touchWinDrag || _touchWinDrag.win !== win) return;
      const state = _touchWinDrag;
      _touchWinDrag = null;
      if (!state.dragging) return;
      titlebar.style.cursor = 'grab';
      if (e.cancelable) e.preventDefault();
    };
    titlebar.addEventListener('touchend', endWinTouch);
    titlebar.addEventListener('touchcancel', endWinTouch);
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
        closeAutoFilter();
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
        const area = document.getElementById('window-area').getBoundingClientRect();
        const cx = Math.max(area.left, Math.min(e.clientX, area.right));
        const cy = Math.max(area.top, Math.min(e.clientY, area.bottom));
        const dx = cx - startX;
        const dy = cy - startY;
        const areaW = area.right - area.left, areaH = area.bottom - area.top;
        if (resizeR) win.el.style.width = Math.min(Math.max(280, origW + dx), areaW - origLeft) + 'px';
        if (resizeB) win.el.style.height = Math.min(Math.max(160, origH + dy), areaH - origTop) + 'px';
        if (resizeL) {
          const newW = Math.max(280, origW - dx);
          const newLeft = Math.max(0, origLeft + origW - newW);
          win.el.style.width = (origLeft + origW - newLeft) + 'px';
          win.el.style.left = newLeft + 'px';
        }
        if (resizeT) {
          const newH = Math.max(160, origH - dy);
          const newTop = Math.max(0, origTop + origH - newH);
          win.el.style.height = (origTop + origH - newTop) + 'px';
          win.el.style.top = newTop + 'px';
        }
      });

      document.addEventListener('mouseup', () => { resizing = false; });
    });
  }

  function setupWindowButtons(win) {
    win.el.querySelector('.btn-close').addEventListener('click', (e) => {
      if (e.ctrlKey || e.metaKey) { closeAllWindows(); } else { closeWindow(win.id); }
    });
    win.el.querySelector('.btn-min').addEventListener('click', () => minimizeWindow(win.id));
    win.el.querySelector('.btn-max').addEventListener('click', () => toggleMaximize(win.id));
    // Ctrl/Cmd-click title text to rename, double-click titlebar to maximize
    win.el.querySelector('.win-title').addEventListener('click', (e) => {
      if ((e.ctrlKey || e.metaKey) && win.tableName) {
        e.stopPropagation();
        startInlineRename(win);
      }
    });
    win.el.querySelector('.win-titlebar').addEventListener('dblclick', (e) => {
      if (e.target.tagName !== 'BUTTON') {
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

  function getDisplayFilename(t) {
    if (!t) return '';
    if (t.excel && t.excel.excelFilename) {
      return t.excel.excelFilename + ' [' + (t.excel.sheetName || 'Sheet1') + ']';
    }
    if (!t.filename) return '';
    if (t.compression && t.compression.type === 'zip' && t.compression.compressedFilename) {
      return t.compression.compressedFilename + '/' + (t.compression.innerName || t.filename);
    }
    return compressedFilename(t.filename, t.compression);
  }

  function updateWindowTitle(tableName) {
    windows.filter(w => w.tableName === tableName).forEach(w => {
      const t = tables[tableName];
      const mod = t && t.modified ? ' *' : '';
      const displayName = getDisplayFilename(t);
      const fname = displayName ? ' — ' + displayName : '';
      w.el.querySelector('.win-title').textContent = tableName + fname + mod;
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

  let _prevAreaWidth = 0;
  let _prevAreaHeight = 0;

  function scaleWindowsToArea() {
    const area = document.getElementById('window-area');
    const rect = area.getBoundingClientRect();
    const w = rect.width;
    const h = rect.height;
    if (_prevAreaWidth === 0 || _prevAreaHeight === 0) {
      _prevAreaWidth = w;
      _prevAreaHeight = h;
      return;
    }
    const scaleX = w / _prevAreaWidth;
    const scaleY = h / _prevAreaHeight;
    if (scaleX === 1 && scaleY === 1) return;
    windows.forEach(win => {
      if (win.el.classList.contains('minimized')) return;
      if (win.maximized) {
        win.el.style.width = w + 'px';
        win.el.style.height = h + 'px';
        if (win.prevBounds) {
          win.prevBounds.left = Math.round(win.prevBounds.left * scaleX);
          win.prevBounds.top = Math.round(win.prevBounds.top * scaleY);
          win.prevBounds.width = Math.round(win.prevBounds.width * scaleX);
          win.prevBounds.height = Math.round(win.prevBounds.height * scaleY);
        }
        return;
      }
      const left = parseFloat(win.el.style.left) || 0;
      const top = parseFloat(win.el.style.top) || 0;
      const width = parseFloat(win.el.style.width) || 0;
      const height = parseFloat(win.el.style.height) || 0;
      win.el.style.left = Math.round(left * scaleX) + 'px';
      win.el.style.top = Math.round(top * scaleY) + 'px';
      win.el.style.width = Math.round(width * scaleX) + 'px';
      win.el.style.height = Math.round(height * scaleY) + 'px';
    });
    _prevAreaWidth = w;
    _prevAreaHeight = h;
  }

  function setupBrowserResize() {
    const area = document.getElementById('window-area');
    const rect = area.getBoundingClientRect();
    _prevAreaWidth = rect.width;
    _prevAreaHeight = rect.height;
    window.addEventListener('resize', scaleWindowsToArea);
  }

  // ---- Table Window ----
  function createTableWindow(tableName) {
    const t = tables[tableName];
    const displayName = getDisplayFilename(t);
    const fname = displayName ? ' — ' + displayName : '';
    const mod = t && t.modified ? ' *' : '';
    createSubwindow(tableName + fname + mod, (win, body) => {
      win.tableName = tableName;
      renderTableView(win, body, t);
    }, { tableName });
    if (_activeConsoleTab === 'ai') populateTableSelect();
  }

  function renderTableView(win, body, tableData) {
    body.innerHTML = '';

    // Toolbar
    const toolbar = document.createElement('div');
    toolbar.className = 'win-toolbar';
    toolbar.innerHTML = `
      <label>Filter:</label>
      <input type="text" class="filter-input" placeholder="WHERE clause, e.g. age > 30 AND name LIKE '%Smith%'" value="${escHtml(win.filterText)}" spellcheck="false">
      <button class="btn-add-row">+ Row</button>
      <button class="btn-add-col">+ Col</button>
    `;
    body.appendChild(toolbar);

    const filterInput = toolbar.querySelector('.filter-input');

    // Wrap filter input with highlight overlay
    const filterWrap = document.createElement('div');
    filterWrap.className = 'filter-highlight-wrap';
    const filterOverlay = document.createElement('div');
    filterOverlay.className = 'filter-highlight';
    filterOverlay.setAttribute('aria-hidden', 'true');
    filterInput.parentNode.insertBefore(filterWrap, filterInput);
    filterWrap.appendChild(filterOverlay);
    filterWrap.appendChild(filterInput);
    function updateFilterHighlight() {
      filterOverlay.innerHTML = sqlHighlightHTML(filterInput.value);
    }
    if (win.filterText) updateFilterHighlight();
    filterInput.addEventListener('scroll', () => {
      filterOverlay.scrollLeft = filterInput.scrollLeft;
    });

    let filterTimeout;
    filterInput.addEventListener('input', () => {
      updateFilterHighlight();
      clearTimeout(filterTimeout);
      filterTimeout = setTimeout(() => {
        win.filterText = filterInput.value;
        if (syncTimers[win.tableName]) {
          clearTimeout(syncTimers[win.tableName]);
          delete syncTimers[win.tableName];
          syncToSQL(win.tableName);
        }
        rebuildTable(win);
        if (win._filterError) {
          filterInput.classList.add('filter-error');
          filterInput.title = win._filterError;
        } else {
          filterInput.classList.remove('filter-error');
          filterInput.title = '';
        }
      }, 200);
    });

    // Escape returns focus from the filter to the selected cell (or middle cell).
    filterInput.addEventListener('keydown', (e) => {
      if (e.key !== 'Escape') return;
      e.preventDefault();
      const displayIdx = win.anchorCell
        ? win._displayRows.findIndex(r => r._rownum === win.anchorCell.rownum)
        : -1;
      const colIdx = win.anchorCell ? win._columns.indexOf(win.anchorCell.col) : -1;
      if (displayIdx >= 0 && colIdx >= 0) {
        focusCellAt(win, displayIdx, colIdx);
      } else {
        focusMiddleCell(win);
      }
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
    closeAutoFilter();
    const { columns, rows } = tableData;
    let displayRows = [...rows];

    // SQL WHERE filter
    if (win.filterText && win.tableName && db) {
      try {
        const sql = `SELECT rowid FROM [${win.tableName}] WHERE ${win.filterText}`;
        const result = db.exec(sql);
        if (result.length > 0) {
          const matchIds = new Set(result[0].values.map(r => r[0]));
          displayRows = displayRows.filter(row => matchIds.has(row._rownum));
        } else {
          displayRows = [];
        }
        win._filterError = null;
      } catch (e) {
        win._filterError = e.message;
      }
    } else {
      win._filterError = null;
    }

    // Column autofilters (AND together)
    const cfKeys = Object.keys(win.columnFilters);
    if (cfKeys.length > 0) {
      displayRows = displayRows.filter(row => {
        for (const c of cfKeys) {
          if (!win.columnFilters[c].has(String(row[c] ?? ''))) return false;
        }
        return true;
      });
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
          }
          const cmp = collator.compare(String(va), String(vb));
          if (cmp !== 0) return cmp * m;
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

    const colgroup = document.createElement('colgroup');
    const rowNumColEl = document.createElement('col');
    rowNumColEl.style.width = '50px';
    colgroup.appendChild(rowNumColEl);
    columns.forEach((col, colIdx) => {
      const colEl = document.createElement('col');
      if (win.colWidths && win.colWidths[colIdx] != null) {
        colEl.style.width = win.colWidths[colIdx] + 'px';
      }
      colgroup.appendChild(colEl);
    });
    table.appendChild(colgroup);
    win._colgroup = colgroup;
    if (win.colWidths) {
      table.classList.add('fixed-layout');
      table.style.width = (50 + win.colWidths.reduce((a, b) => a + b, 0)) + 'px';
    }

    // Header
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    const rowNumTh = document.createElement('th');
    rowNumTh.className = 'row-num-header';
    rowNumTh.textContent = '#';
    headerRow.appendChild(rowNumTh);

    columns.forEach((col, colIdx) => {
      const th = document.createElement('th');
      const thInner = document.createElement('div');
      thInner.className = 'th-inner';
      const colLabel = document.createElement('span');
      colLabel.className = 'col-name';
      colLabel.textContent = col;
      thInner.appendChild(colLabel);
      th.dataset.colIdx = colIdx;
      const sortIdx = win.sortCols.findIndex(s => s.col === col);
      if (sortIdx !== -1) {
        th.classList.add('sorted');
        const arrow = document.createElement('span');
        arrow.className = 'sort-arrow';
        const dir = win.sortCols[sortIdx].dir;
        arrow.textContent = dir === 'asc' ? '\u25B2' : '\u25BC';
        if (win.sortCols.length > 1) arrow.textContent += (sortIdx + 1);
        thInner.appendChild(arrow);
      }
      if (win.columnFilters[col]) th.classList.add('col-filtered');
      if (win.selectedCol === col) th.classList.add('col-selected');

      if (hasDisplayTransform(win.tableName, col)) {
        const fxIcon = document.createElement('span');
        const colDisabled = win.disabledTransforms.has(col);
        fxIcon.className = 'col-transform-icon' + (colDisabled ? ' disabled' : '');
        fxIcon.textContent = '\u{1F50C}';
        fxIcon.title = colDisabled ? 'Plugin transform disabled — click to enable' : 'Plugin transform active — click to disable';
        fxIcon.addEventListener('mousedown', (e) => { e.stopPropagation(); });
        fxIcon.addEventListener('click', (e) => {
          e.stopPropagation();
          e.preventDefault();
          if (win.disabledTransforms.has(col)) win.disabledTransforms.delete(col);
          else win.disabledTransforms.add(col);
          rebuildTable(win);
        });
        thInner.appendChild(fxIcon);
      }

      // AutoFilter button
      const filterBtn = document.createElement('span');
      filterBtn.className = 'col-filter-btn';
      if (win.columnFilters[col]) filterBtn.classList.add('active');
      filterBtn.textContent = '\u2630';
      filterBtn.addEventListener('mousedown', (e) => { e.stopPropagation(); });
      filterBtn.addEventListener('touchstart', (e) => { e.stopPropagation(); }, { passive: true });
      filterBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        e.preventDefault();
        openAutoFilter(win, col, th);
      });
      thInner.appendChild(filterBtn);
      th.appendChild(thInner);

      const resizeHandle = document.createElement('div');
      resizeHandle.className = 'col-resize-handle';
      resizeHandle.addEventListener('mousedown', (e) => {
        e.stopPropagation();
        e.preventDefault();
        startColResize(win, colIdx, e);
      });
      resizeHandle.addEventListener('dblclick', (e) => {
        e.stopPropagation();
        e.preventDefault();
        autoFitColumn(win, colIdx);
      });
      th.appendChild(resizeHandle);

      // Drag to reorder columns; click to sort/select; Ctrl/Cmd+click to rename
      th.addEventListener('mousedown', (e) => {
        if (e.button !== 0 || th._renaming) return;
        e.preventDefault();
        const startX = e.clientX;
        const startY = e.clientY;
        let dragging = false;
        let ghost = null;
        th._didDrag = false;
        const onMove = (me) => {
          if (!dragging && Math.abs(me.clientX - startX) > 5) {
            dragging = true;
            th._didDrag = true;
            th.classList.add('col-dragging');
            ghost = document.createElement('div');
            ghost.className = 'col-drag-ghost';
            ghost.textContent = col;
            ghost.style.left = me.clientX + 'px';
            ghost.style.top = startY + 'px';
            document.body.appendChild(ghost);
          }
          if (dragging) {
            ghost.style.left = me.clientX + 'px';
            const ths = headerRow.querySelectorAll('th:not(.row-num-header)');
            ths.forEach(h => h.classList.remove('col-drop-left', 'col-drop-right'));
            for (const h of ths) {
              const rect = h.getBoundingClientRect();
              const mid = rect.left + rect.width / 2;
              if (me.clientX >= rect.left && me.clientX <= rect.right) {
                h.classList.add(me.clientX < mid ? 'col-drop-left' : 'col-drop-right');
                break;
              }
            }
          }
        };
        const onUp = (ue) => {
          document.removeEventListener('mousemove', onMove);
          document.removeEventListener('mouseup', onUp);
          th.classList.remove('col-dragging');
          if (ghost) ghost.remove();
          const ths = headerRow.querySelectorAll('th:not(.row-num-header)');
          ths.forEach(h => h.classList.remove('col-drop-left', 'col-drop-right'));
          if (dragging) {
            let dropIdx = colIdx;
            for (const h of ths) {
              const rect = h.getBoundingClientRect();
              if (ue.clientX >= rect.left && ue.clientX <= rect.right) {
                const mid = rect.left + rect.width / 2;
                dropIdx = parseInt(h.dataset.colIdx);
                if (ue.clientX >= mid && dropIdx < columns.length) dropIdx++;
                break;
              }
            }
            if (dropIdx !== colIdx) {
              reorderColumn(win, colIdx, dropIdx);
            }
          }
        };
        document.addEventListener('mousemove', onMove);
        document.addEventListener('mouseup', onUp);
      });

      // Touch: 1.5-tap-to-drag (tap once, then tap-and-hold + pan) to reorder
      th.addEventListener('touchstart', (e) => {
        if (e.touches.length !== 1 || th._renaming) return;
        const touch = e.touches[0];
        const now = Date.now();
        const prev = _lastHeaderTap;
        const isSecondTap = prev.tableName === win.tableName && prev.col === col &&
                            (now - prev.time) < DOUBLE_TAP_WINDOW_MS;
        if (!isSecondTap) {
          // First tap — let native click/sort happen; just record for a potential second tap.
          _lastHeaderTap = { tableName: win.tableName, col, time: now };
          return;
        }
        // Second tap: enter drag-pending mode; drag visuals start on first movement.
        _lastHeaderTap = { tableName: null, col: null, time: 0 };
        _touchHeaderDrag = {
          th, col, colIdx,
          startX: touch.clientX, startY: touch.clientY,
          dragging: false, ghost: null,
        };
        th._didDrag = false;
      }, { passive: true });

      th.addEventListener('touchmove', (e) => {
        if (!_touchHeaderDrag || _touchHeaderDrag.th !== th) return;
        const touch = e.touches[0];
        if (!touch) return;
        const dx = touch.clientX - _touchHeaderDrag.startX;
        const dy = touch.clientY - _touchHeaderDrag.startY;
        if (!_touchHeaderDrag.dragging && (Math.abs(dx) > 5 || Math.abs(dy) > 5)) {
          _touchHeaderDrag.dragging = true;
          th._didDrag = true;
          th.classList.add('col-dragging');
          const ghost = document.createElement('div');
          ghost.className = 'col-drag-ghost';
          ghost.textContent = col;
          ghost.style.left = touch.clientX + 'px';
          ghost.style.top = touch.clientY + 'px';
          document.body.appendChild(ghost);
          _touchHeaderDrag.ghost = ghost;
          if (navigator.vibrate) navigator.vibrate(10);
        }
        if (_touchHeaderDrag.dragging) {
          e.preventDefault();
          _touchHeaderDrag.ghost.style.left = touch.clientX + 'px';
          _touchHeaderDrag.ghost.style.top = touch.clientY + 'px';
          const ths = headerRow.querySelectorAll('th:not(.row-num-header)');
          ths.forEach(h => h.classList.remove('col-drop-left', 'col-drop-right'));
          const el = document.elementFromPoint(touch.clientX, touch.clientY);
          const targetTh = el && el.closest && el.closest('th');
          if (targetTh && headerRow.contains(targetTh) && !targetTh.classList.contains('row-num-header')) {
            const rect = targetTh.getBoundingClientRect();
            const mid = rect.left + rect.width / 2;
            targetTh.classList.add(touch.clientX < mid ? 'col-drop-left' : 'col-drop-right');
          }
        }
      }, { passive: false });

      const endHeaderTouch = (e) => {
        if (!_touchHeaderDrag || _touchHeaderDrag.th !== th) return;
        const state = _touchHeaderDrag;
        _touchHeaderDrag = null;
        if (!state.dragging) return;
        th.classList.remove('col-dragging');
        if (state.ghost) state.ghost.remove();
        const ths = headerRow.querySelectorAll('th:not(.row-num-header)');
        ths.forEach(h => h.classList.remove('col-drop-left', 'col-drop-right'));
        const last = (e.changedTouches && e.changedTouches[0]) || null;
        if (!last) return;
        let dropIdx = colIdx;
        const el = document.elementFromPoint(last.clientX, last.clientY);
        const targetTh = el && el.closest && el.closest('th');
        if (targetTh && headerRow.contains(targetTh) && !targetTh.classList.contains('row-num-header')) {
          const rect = targetTh.getBoundingClientRect();
          const mid = rect.left + rect.width / 2;
          dropIdx = parseInt(targetTh.dataset.colIdx);
          if (last.clientX >= mid && dropIdx < columns.length) dropIdx++;
        }
        if (dropIdx !== colIdx) reorderColumn(win, colIdx, dropIdx);
      };
      th.addEventListener('touchend', endHeaderTouch);
      th.addEventListener('touchcancel', endHeaderTouch);

      th.addEventListener('click', (e) => {
        if (th._renaming) return;
        if (th._didDrag) { th._didDrag = false; return; }
        if (e.ctrlKey || e.metaKey) {
          e.stopPropagation();
          startColumnRename(win, th, col);
          return;
        }
        win.selectedCol = col;
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
      if (td.tagName !== 'TD' || !td.classList.contains('data-cell')) return;
      if (td.getAttribute('contenteditable') !== 'true') return;
      const tr = td.parentElement;
      const displayIdx = parseInt(tr.dataset.displayIdx, 10);
      const colIdx = parseInt(td.dataset.colIdx, 10);
      if (isNaN(displayIdx) || isNaN(colIdx)) return;
      const row = win._displayRows[displayIdx];
      const col = win._columns[colIdx];
      if (!row || col == null) { exitEditMode(td); return; }
      const newVal = td.textContent;
      const oldVal = String(row[col] ?? '');
      if (newVal !== oldVal) {
        row[col] = newVal;
        td.classList.add('modified');
        markModified(win.tableName);
        const t = tables[win.tableName];
        if (t) {
          if (!t._dirtyCells) t._dirtyCells = [];
          t._dirtyCells.push({ rownum: row._rownum, col, value: newVal });
          if (!t._undoStack) t._undoStack = [];
          t._undoStack.push({ type: 'edit', changes: [{ rownum: row._rownum, col, oldValue: oldVal, newValue: newVal }] });
          t._redoStack = [];
        }
        debouncedSync(win.tableName);
      }
      exitEditMode(td);
      if (win.tableName && col && !win.disabledTransforms.has(col) && hasDisplayTransform(win.tableName, col)) {
        td.textContent = getDisplayValue(win.tableName, col, row);
      }
    }, true); // capture phase for blur

    table.addEventListener('focusin', (e) => {
      const td = e.target;
      if (td.tagName !== 'TD' || !td.classList.contains('data-cell')) return;
      if (win._programmaticFocus) return;
      const tr = td.parentElement;
      const di = parseInt(tr.dataset.displayIdx, 10);
      const ci = parseInt(td.dataset.colIdx, 10);
      if (isNaN(di) || isNaN(ci)) return;
      const row = win._displayRows[di];
      if (!row) return;
      const col = win._columns[ci];
      win.anchorCell = { rownum: row._rownum, col };
      win.selectedCells = new Set([`${row._rownum}:${col}`]);
      win._copyWithHeader = false;
      applyCellHighlights(win);
    });

    // Mouse multi-select: Shift+click extends; plain mousedown+drag sweeps a rectangle.
    let dragState = null;
    const onDragMove = (ev) => {
      if (!dragState) return;
      const el = document.elementFromPoint(ev.clientX, ev.clientY);
      const td = el && el.closest && el.closest('td.data-cell');
      if (!td || !table.contains(td)) return;
      const di = parseInt(td.parentElement.dataset.displayIdx, 10);
      const ci = parseInt(td.dataset.colIdx, 10);
      if (isNaN(di) || isNaN(ci)) return;
      if (di === dragState.lastDi && ci === dragState.lastCi) return;
      if (!dragState.isDragging && (di !== dragState.startDi || ci !== dragState.startCi)) {
        dragState.isDragging = true;
        table.classList.add('drag-selecting');
      }
      if (dragState.isDragging) {
        const sel = window.getSelection && window.getSelection();
        if (sel && sel.removeAllRanges) sel.removeAllRanges();
        rebuildSelectionRect(win, di, ci);
        applyCellHighlights(win);
      }
      dragState.lastDi = di;
      dragState.lastCi = ci;
    };
    const onDragEnd = () => {
      if (dragState && dragState.isDragging) {
        table.classList.remove('drag-selecting');
        focusCellAt(win, dragState.lastDi, dragState.lastCi);
      }
      dragState = null;
      document.removeEventListener('mousemove', onDragMove);
      document.removeEventListener('mouseup', onDragEnd);
    };
    table.addEventListener('mousedown', (e) => {
      if (e.button !== 0) return;
      const td = e.target.closest && e.target.closest('td.data-cell');
      if (!td || !table.contains(td)) return;
      // Click on a cell already in edit mode: let the native click place the caret.
      if (td.getAttribute('contenteditable') === 'true') return;
      const tr = td.parentElement;
      const di = parseInt(tr.dataset.displayIdx, 10);
      const ci = parseInt(td.dataset.colIdx, 10);
      if (isNaN(di) || isNaN(ci)) return;
      win._copyWithHeader = false;
      if (!e.shiftKey && !(e.ctrlKey || e.metaKey)) {
        const row = win._displayRows[di];
        if (row) {
          const col = win._columns[ci];
          win.anchorCell = { rownum: row._rownum, col };
          win.selectedCells = new Set([`${row._rownum}:${col}`]);
          applyCellHighlights(win);
        }
      }
      if (e.shiftKey && win.anchorCell) {
        e.preventDefault();
        rebuildSelectionRect(win, di, ci);
        focusCellAt(win, di, ci);
        applyCellHighlights(win);
        return;
      }
      if ((e.ctrlKey || e.metaKey) && !e.shiftKey) {
        e.preventDefault();
        td.focus();
        enterEditMode(td);
        return;
      }
      dragState = { startDi: di, startCi: ci, lastDi: di, lastCi: ci, isDragging: false };
      document.addEventListener('mousemove', onDragMove);
      document.addEventListener('mouseup', onDragEnd);
    });

    // Select-all via corner cell click
    table.addEventListener('click', (e) => {
      const th = e.target.closest && e.target.closest('th.row-num-header');
      if (!th || !table.contains(th)) return;
      selectAllCells(win);
    });

    // Row selection via row-number cell click/drag
    let rowDragState = null;
    const onRowDragMove = (ev) => {
      if (!rowDragState) return;
      const el = document.elementFromPoint(ev.clientX, ev.clientY);
      const td = el && el.closest && el.closest('td.row-num');
      if (!td || !table.contains(td)) return;
      const di = parseInt(td.parentElement.dataset.displayIdx, 10);
      if (isNaN(di) || di === rowDragState.lastDi) return;
      rowDragState.lastDi = di;
      const sel = window.getSelection && window.getSelection();
      if (sel && sel.removeAllRanges) sel.removeAllRanges();
      selectRows(win, rowDragState.anchorDi, di);
    };
    const onRowDragEnd = () => {
      if (rowDragState) {
        focusCellAt(win, rowDragState.lastDi, 0);
      }
      rowDragState = null;
      document.removeEventListener('mousemove', onRowDragMove);
      document.removeEventListener('mouseup', onRowDragEnd);
    };
    table.addEventListener('mousedown', (e) => {
      if (e.button !== 0) return;
      const td = e.target.closest && e.target.closest('td.row-num');
      if (!td || !table.contains(td)) return;
      const di = parseInt(td.parentElement.dataset.displayIdx, 10);
      if (isNaN(di)) return;
      e.preventDefault();
      if (e.shiftKey && win.anchorCell) {
        const anchorDi = win._displayRows.findIndex(r => r._rownum === win.anchorCell.rownum);
        if (anchorDi >= 0) {
          selectRows(win, anchorDi, di);
          focusCellAt(win, di, 0);
        }
        return;
      }
      selectRows(win, di, di);
      rowDragState = { anchorDi: di, lastDi: di };
      document.addEventListener('mousemove', onRowDragMove);
      document.addEventListener('mouseup', onRowDragEnd);
    });

    // Cell touch handling — two gestures, both starting from a second tap
    // within DOUBLE_TAP_WINDOW_MS of the first:
    //  - Double-tap (quick two taps, no pan) → enter edit mode on the tapped cell
    //  - 1.5-tap (tap, then tap-and-pan) → extend the cell selection rectangle
    //    from the anchor cell to the panned cell (like a mouse drag-select)
    // The two are disambiguated at touchend: if the second tap moved >5 px we
    // already switched into drag mode; otherwise we treat it as a double-tap.
    //
    // For edit mode, iOS opens the virtual keyboard only when a native
    // tap/click lands on an already-editable element — programmatic focus()
    // inside a touch handler doesn't count. So touchend sets
    // contenteditable="true" synchronously and does NOT preventDefault: the
    // browser's synthesized click then lands on the now-editable cell and iOS
    // opens the keyboard natively.
    const getCellCoords = (clientX, clientY) => {
      const el = document.elementFromPoint(clientX, clientY);
      const td = el && el.closest && el.closest('td.data-cell');
      if (!td || !table.contains(td)) return null;
      const tr = td.parentElement;
      const di = parseInt(tr.dataset.displayIdx, 10);
      const ci = parseInt(td.dataset.colIdx, 10);
      if (isNaN(di) || isNaN(ci)) return null;
      return { di, ci };
    };

    table.addEventListener('touchstart', (e) => {
      if (e.touches.length !== 1) return;
      const touch = e.touches[0];
      const td = e.target.closest && e.target.closest('td.data-cell');
      if (!td || !table.contains(td)) return;
      if (td.getAttribute('contenteditable') === 'true') return;

      const tr = td.parentElement;
      const di = parseInt(tr.dataset.displayIdx, 10);
      const ci = parseInt(td.dataset.colIdx, 10);
      const row = !isNaN(di) ? win._displayRows[di] : null;
      const col = !isNaN(ci) ? win._columns[ci] : null;

      const now = Date.now();
      const prev = _lastCellTap;
      const isSecondTap = prev.winId === win.id &&
                          (now - prev.time) < DOUBLE_TAP_WINDOW_MS;
      if (isSecondTap) {
        // Seed the anchor from the first-tap cell so if the second tap pans
        // into a drag-select, the rectangle starts from tap 1. Real devices
        // get this for free via the synthesized click → focusin; we set it
        // explicitly to be robust across browsers and test environments.
        if (prev.rownum != null && prev.col != null) {
          win.anchorCell = { rownum: prev.rownum, col: prev.col };
        }
        _lastCellTap = { winId: null, rownum: null, col: null, time: 0 };
        _touchCellDrag = {
          win,
          td,
          startX: touch.clientX, startY: touch.clientY,
          dragging: false,
          lastDi: isNaN(di) ? -1 : di,
          lastCi: isNaN(ci) ? -1 : ci,
        };
        return;
      }

      // First tap: record the cell identity for a potential second tap.
      _lastCellTap = {
        winId: win.id,
        rownum: row ? row._rownum : null,
        col,
        time: now,
      };
    }, { passive: true });

    table.addEventListener('touchmove', (e) => {
      if (!_touchCellDrag || _touchCellDrag.win !== win) return;
      const touch = e.touches[0];
      if (!touch) return;
      const dx = touch.clientX - _touchCellDrag.startX;
      const dy = touch.clientY - _touchCellDrag.startY;
      if (!_touchCellDrag.dragging && (Math.abs(dx) > 5 || Math.abs(dy) > 5)) {
        _touchCellDrag.dragging = true;
        table.classList.add('drag-selecting');
        const sel = window.getSelection && window.getSelection();
        if (sel && sel.removeAllRanges) sel.removeAllRanges();
        if (navigator.vibrate) navigator.vibrate(10);
      }
      if (_touchCellDrag.dragging) {
        if (e.cancelable) e.preventDefault();
        const c = getCellCoords(touch.clientX, touch.clientY);
        if (c && (c.di !== _touchCellDrag.lastDi || c.ci !== _touchCellDrag.lastCi)) {
          _touchCellDrag.lastDi = c.di;
          _touchCellDrag.lastCi = c.ci;
          rebuildSelectionRect(win, c.di, c.ci);
          applyCellHighlights(win);
        }
      }
    }, { passive: false });

    const endCellInteraction = (e) => {
      if (!_touchCellDrag || _touchCellDrag.win !== win) return;
      const state = _touchCellDrag;
      _touchCellDrag = null;
      if (state.dragging) {
        // Drag-select: finalize the rectangle and focus the panned cell.
        // preventDefault suppresses the synthesized click so the focusin
        // handler doesn't collapse the selection back to a single cell.
        table.classList.remove('drag-selecting');
        if (e.cancelable) e.preventDefault();
        if (state.lastDi >= 0 && state.lastCi >= 0) {
          focusCellAt(win, state.lastDi, state.lastCi);
        }
      } else {
        // Double-tap without pan → enter edit mode. Setting contenteditable
        // synchronously (and NOT calling preventDefault) lets the synthesized
        // click on the now-editable cell open the iOS virtual keyboard.
        if (state.td && state.td.isConnected) {
          state.td.setAttribute('contenteditable', 'true');
        }
      }
    };
    table.addEventListener('touchend', endCellInteraction);
    table.addEventListener('touchcancel', endCellInteraction);

    table.addEventListener('keydown', (e) => {
      const td = e.target;
      if (td.tagName !== 'TD' || !td.classList.contains('data-cell')) return;
      const tr = td.parentElement;
      const inEdit = td.getAttribute('contenteditable') === 'true';
      // Vim h/j/k/l in select mode map to arrow-key equivalents.
      const vimMap = { h: 'ArrowLeft', j: 'ArrowDown', k: 'ArrowUp', l: 'ArrowRight' };
      const navKey = (!inEdit && !e.ctrlKey && !e.metaKey && !e.altKey && e.key.length === 1)
        ? (vimMap[e.key.toLowerCase()] || e.key)
        : e.key;
      const isArrow = navKey === 'ArrowUp' || navKey === 'ArrowDown' ||
                      navKey === 'ArrowLeft' || navKey === 'ArrowRight';

      // F2, Enter, i, or Ctrl/Cmd+U: enter edit mode (select mode only)
      const noMods = !e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey;
      if (!inEdit && (e.key === 'F2' ||
          (noMods && (e.key === 'Enter' || e.key === 'i')) ||
          ((e.ctrlKey || e.metaKey) && !e.shiftKey && !e.altKey && (e.key === 'u' || e.key === 'U')))) {
        e.preventDefault();
        enterEditMode(td);
        return;
      }

      // "/" in select mode → jump to this window's filter input
      if (!inEdit && noMods && e.key === '/') {
        const filterInput = win.el.querySelector('.filter-input');
        if (filterInput) {
          e.preventDefault();
          filterInput.focus();
          return;
        }
      }

      // Select-mode arrow key handling
      if (!inEdit && isArrow) {
        const plain = !e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey;
        const shiftOnly = e.shiftKey && !e.ctrlKey && !e.metaKey && !e.altKey;
        const ctrlOnly = (e.ctrlKey || e.metaKey) && !e.shiftKey && !e.altKey;
        if (plain) {
          e.preventDefault();
          moveSingleCellSelection(win, tr, td, navKey);
          return;
        }
        if (shiftOnly) {
          e.preventDefault();
          extendCellSelection(win, tr, td, navKey);
          return;
        }
        if (ctrlOnly && (navKey === 'ArrowLeft' || navKey === 'ArrowRight')) {
          e.preventDefault();
          moveSelectionColumns(win, tr, td, navKey);
          return;
        }
      }

      // Select-mode Tab / Shift+Tab: cycle table windows.
      // Edit-mode Tab still moves between cells in the row.
      if (e.key === 'Tab' && !inEdit && !e.ctrlKey && !e.metaKey && !e.altKey) {
        e.preventDefault();
        cycleTableWindow(win, e.shiftKey ? -1 : 1);
        return;
      }
      // Ctrl/Cmd (+Shift) + H/J/K/L in select mode:
      //   +Shift L/H → cycle next / previous table window
      //   plain H/J/K/L → nudge the window 5 px in the vim direction
      if (!inEdit && (e.ctrlKey || e.metaKey) && !e.altKey && e.key.length === 1) {
        const k = e.key.toLowerCase();
        if (e.shiftKey && (k === 'l' || k === 'h')) {
          e.preventDefault();
          cycleTableWindow(win, k === 'l' ? 1 : -1);
          return;
        }
        if (!e.shiftKey) {
          const nudgeMap = { h: [-5, 0], j: [0, 5], k: [0, -5], l: [5, 0] };
          if (nudgeMap[k]) {
            e.preventDefault();
            nudgeWindow(win, nudgeMap[k][0], nudgeMap[k][1]);
            return;
          }
        }
      }

      // Copy / Paste / Undo / Redo — select mode only (edit mode falls through to native)
      if (!inEdit && (e.ctrlKey || e.metaKey) && !e.altKey) {
        const k = e.key.toLowerCase();
        if (k === 'x' && !e.shiftKey) { e.preventDefault(); cutSelectedCells(win); return; }
        if (k === 'c' && !e.shiftKey) { e.preventDefault(); copySelectedCells(win); return; }
        if (k === 'v' && !e.shiftKey) { e.preventDefault(); pasteAtAnchor(win); return; }
        if (k === 'z' && !e.shiftKey) { e.preventDefault(); undoTable(win.tableName, win); return; }
        if (k === 'z' && e.shiftKey) { e.preventDefault(); redoTable(win.tableName, win); return; }
        if (k === 'a' && e.shiftKey) { e.preventDefault(); selectAllCells(win); return; }
      }

      if (e.key === 'Tab') {
        e.preventDefault();
        const cells = [...tr.querySelectorAll('td.data-cell')];
        const idx = cells.indexOf(td);
        const next = e.shiftKey ? cells[idx - 1] : cells[idx + 1];
        if (next) next.focus();
      } else if (e.key === 'Enter' && !e.shiftKey) {
        if (!inEdit) return;
        e.preventDefault();
        td.blur();
        const nextTr = tr.nextElementSibling;
        if (nextTr && !nextTr.classList.contains('virtual-pad')) {
          const colIdx = [...tr.children].indexOf(td);
          const nextTd = nextTr.children[colIdx];
          if (nextTd && nextTd.classList.contains('data-cell')) nextTd.focus();
        }
      } else if (e.key === 'Escape') {
        e.preventDefault();
        if (inEdit) {
          // Revert edit, exit edit mode, keep selection and focus on this cell.
          const displayIdx = parseInt(tr.dataset.displayIdx, 10);
          const colIdx = parseInt(td.dataset.colIdx, 10);
          if (!isNaN(displayIdx) && !isNaN(colIdx)) {
            const row = win._displayRows[displayIdx];
            const col = win._columns[colIdx];
            if (row && col != null) {
              td.textContent = row[col] ?? '';
              if (win.tableName && !win.disabledTransforms.has(col) && hasDisplayTransform(win.tableName, col)) {
                td.textContent = getDisplayValue(win.tableName, col, row);
              }
            }
          }
          exitEditMode(td);
          win._programmaticFocus = true;
          td.focus();
          win._programmaticFocus = false;
        } else {
          td.blur();
          win.selectedCells = new Set();
          win.anchorCell = null;
          applyCellHighlights(win);
        }
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

    if (!win.colWidths) {
      const ths = table.querySelectorAll('thead th:not(.row-num-header)');
      win.colWidths = [];
      for (const th of ths) win.colWidths.push(th.offsetWidth);
      const cols = colgroup.querySelectorAll('col');
      for (let i = 0; i < win.colWidths.length; i++) {
        cols[i + 1].style.width = win.colWidths[i] + 'px';
      }
      table.classList.add('fixed-layout');
      updateTableWidth(win);
    }

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
    const hasColumnFilters = Object.keys(win.columnFilters).length > 0;
    const hasWhereFilter = !!win.filterText;
    if (hasColumnFilters || hasWhereFilter) {
      statusLeft.textContent = `${displayRows.length} of ${rows.length} rows — `;
      const clearLink = document.createElement('span');
      clearLink.className = 'status-clear-filters';
      clearLink.textContent = 'Clear Filters';
      clearLink.addEventListener('click', () => {
        win.columnFilters = {};
        win.filterText = '';
        const filterInput = win.el.querySelector('.filter-input');
        if (filterInput) filterInput.value = '';
        rebuildTable(win);
      });
      statusLeft.appendChild(clearLink);
    } else {
      statusLeft.textContent = `${displayRows.length} of ${rows.length} rows`;
    }
    statusRight.textContent = `${columns.length} columns`;

    const statusCenter = win.el.querySelector('.status-center');
    statusCenter.innerHTML = '';
    const transformedCols = win.tableName && _columnTransformCache[win.tableName]
      ? Object.keys(_columnTransformCache[win.tableName]) : [];
    if (transformedCols.length > 0) {
      const allOff = transformedCols.every(c => win.disabledTransforms.has(c));
      const allOn = !transformedCols.some(c => win.disabledTransforms.has(c));
      const toggle = document.createElement('span');
      const state = allOn ? 'on' : allOff ? 'off' : 'partial';
      toggle.className = 'status-plugin-toggle' + (state !== 'on' ? ' disabled' : '');
      toggle.title = state === 'on' ? 'All plugins enabled — click to disable all'
        : 'Some plugins disabled — click to enable all';
      toggle.textContent = state === 'on' ? '⊕ Plugins on'
        : state === 'off' ? '⊘ Plugins off'
        : '⊕ Plugins partial';
      toggle.addEventListener('click', (e) => {
        e.stopPropagation();
        if (allOn) transformedCols.forEach(c => win.disabledTransforms.add(c));
        else transformedCols.forEach(c => win.disabledTransforms.delete(c));
        rebuildTable(win);
      });
      statusCenter.appendChild(toggle);
    }
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
        td.textContent = win.disabledTransforms.has(columns[c]) ? (row[columns[c]] ?? '') : getDisplayValue(win.tableName, columns[c], row);
        td.className = 'data-cell';
        td.tabIndex = 0;
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

    applyCellHighlights(win);
  }

  function rebuildSelectionRect(win, leadDisplay, leadCol) {
    if (!win.anchorCell) {
      const row = win._displayRows[leadDisplay];
      if (row) win.anchorCell = { rownum: row._rownum, col: win._columns[leadCol] };
    }
    const anchorDisplay = win.anchorCell
      ? win._displayRows.findIndex(r => r._rownum === win.anchorCell.rownum)
      : -1;
    const anchorCol = win.anchorCell ? win._columns.indexOf(win.anchorCell.col) : -1;
    const next = new Set();
    if (anchorDisplay === -1 || anchorCol === -1) {
      const row = win._displayRows[leadDisplay];
      if (row) {
        const col = win._columns[leadCol];
        win.anchorCell = { rownum: row._rownum, col };
        next.add(`${row._rownum}:${col}`);
      }
    } else {
      const r1 = Math.min(anchorDisplay, leadDisplay);
      const r2 = Math.max(anchorDisplay, leadDisplay);
      const c1 = Math.min(anchorCol, leadCol);
      const c2 = Math.max(anchorCol, leadCol);
      for (let r = r1; r <= r2; r++) {
        const row = win._displayRows[r];
        if (!row) continue;
        for (let c = c1; c <= c2; c++) {
          next.add(`${row._rownum}:${win._columns[c]}`);
        }
      }
    }
    win.selectedCells = next;
  }

  function focusMiddleCell(win) {
    const container = win._container;
    const table = win._table;
    const tbody = win._tbody;
    const displayRows = win._displayRows;
    if (!container || !table || !tbody || !displayRows || !displayRows.length) return false;
    const rect = container.getBoundingClientRect();
    const thead = table.querySelector('thead');
    const theadH = thead ? thead.offsetHeight : 0;
    const cx = rect.left + rect.width / 2;
    const cy = rect.top + theadH + (rect.height - theadH) / 2;
    const el = document.elementFromPoint(cx, cy);
    const td = el && el.closest && el.closest('td.data-cell');
    if (td && table.contains(td)) {
      td.focus();
      return true;
    }
    return false;
  }

  function enterEditMode(td) {
    const tr = td.parentElement;
    const winEl = td.closest('.subwindow');
    const win = winEl && windows.find(w => w.el === winEl);
    if (win && win.tableName) {
      const colIdx = parseInt(td.dataset.colIdx, 10);
      const displayIdx = parseInt(tr.dataset.displayIdx, 10);
      if (!isNaN(colIdx) && !isNaN(displayIdx) && win._columns && win._displayRows) {
        const col = win._columns[colIdx];
        if (col && !win.disabledTransforms.has(col) && hasDisplayTransform(win.tableName, col)) {
          const row = win._displayRows[displayIdx];
          if (row) td.textContent = row[col] ?? '';
        }
      }
    }
    td.setAttribute('contenteditable', 'true');
    td.focus();
    const range = document.createRange();
    range.selectNodeContents(td);
    range.collapse(false); // caret at end
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }

  function exitEditMode(td) {
    if (td.getAttribute('contenteditable') === 'true') {
      td.removeAttribute('contenteditable');
    }
  }

  function copySelectedCells(win) {
    if (!win.selectedCells || win.selectedCells.size === 0) return;
    const t = tables[win.tableName];
    if (!t) return;
    const colOrder = new Map(t.columns.map((c, i) => [c, i]));
    const displayOrder = new Map(win._displayRows.map((r, i) => [r._rownum, i]));
    const cells = [];
    for (const key of win.selectedCells) {
      const sep = key.indexOf(':');
      const rownum = parseInt(key.slice(0, sep), 10);
      const col = key.slice(sep + 1);
      if (displayOrder.has(rownum) && colOrder.has(col)) cells.push({ rownum, col });
    }
    if (cells.length === 0) return;
    cells.sort((a, b) => (displayOrder.get(a.rownum) - displayOrder.get(b.rownum)) || (colOrder.get(a.col) - colOrder.get(b.col)));
    const rowMap = new Map(t.rows.map(r => [r._rownum, r]));
    const rows = [];
    let curRownum = null, curRow = [];
    for (const { rownum, col } of cells) {
      if (rownum !== curRownum) {
        if (curRow.length) rows.push(curRow);
        curRow = [];
        curRownum = rownum;
      }
      const row = rowMap.get(rownum);
      curRow.push(String(row ? (row[col] ?? '') : ''));
    }
    if (curRow.length) rows.push(curRow);
    if (win._copyWithHeader) {
      const selCols = [...new Set(cells.map(c => c.col))].sort((a, b) => colOrder.get(a) - colOrder.get(b));
      rows.unshift(selCols);
    }
    const tsv = rows.map(r => r.join('\t')).join('\n');
    try { navigator.clipboard.writeText(tsv); } catch (_) {}
  }

  function cutSelectedCells(win) {
    if (!win.selectedCells || win.selectedCells.size === 0) return;
    const t = tables[win.tableName];
    if (!t) return;
    copySelectedCells(win);
    const changes = [];
    if (!t._dirtyCells) t._dirtyCells = [];
    for (const key of win.selectedCells) {
      const sep = key.indexOf(':');
      const rownum = parseInt(key.slice(0, sep), 10);
      const col = key.slice(sep + 1);
      const row = t.rows.find(r => r._rownum === rownum);
      if (!row || !t.columns.includes(col)) continue;
      const oldValue = String(row[col] ?? '');
      if (oldValue !== '') {
        row[col] = '';
        changes.push({ rownum, col, oldValue, newValue: '' });
        t._dirtyCells.push({ rownum, col, value: '' });
      }
    }
    if (changes.length === 0) return;
    if (!t._undoStack) t._undoStack = [];
    t._undoStack.push({ type: 'cut', changes });
    t._redoStack = [];
    markModified(win.tableName);
    debouncedSync(win.tableName);
    windows.filter(w => w.tableName === win.tableName).forEach(w => {
      w._renderStart = -1; w._renderEnd = -1;
      renderVisibleRows(w);
    });
    refocusAnchorCell(win);
  }

  function refocusAnchorCell(win) {
    if (!win.anchorCell || !win._displayRows || !win._columns) return;
    const di = win._displayRows.findIndex(r => r._rownum === win.anchorCell.rownum);
    const ci = win._columns.indexOf(win.anchorCell.col);
    if (di < 0 || ci < 0) return;
    const tr = win._tbody && win._tbody.querySelector(`tr[data-display-idx="${di}"]`);
    if (!tr) return;
    const td = tr.children[ci + 1];
    if (!td) return;
    win._programmaticFocus = true;
    td.focus();
    win._programmaticFocus = false;
  }

  function selectAllCells(win) {
    if (!win._displayRows || !win._columns || win._columns.length === 0) return;
    if (win._displayRows.length === 0) return;
    win.selectedCells = new Set();
    for (const row of win._displayRows) {
      for (const col of win._columns) {
        win.selectedCells.add(`${row._rownum}:${col}`);
      }
    }
    win.anchorCell = { rownum: win._displayRows[0]._rownum, col: win._columns[0] };
    win._copyWithHeader = true;
    applyCellHighlights(win);
    refocusAnchorCell(win);
  }

  function selectRows(win, fromDi, toDi) {
    if (!win._displayRows || !win._columns || win._columns.length === 0) return;
    const r1 = Math.min(fromDi, toDi);
    const r2 = Math.max(fromDi, toDi);
    win.selectedCells = new Set();
    for (let di = r1; di <= r2; di++) {
      const row = win._displayRows[di];
      if (!row) continue;
      for (const col of win._columns) {
        win.selectedCells.add(`${row._rownum}:${col}`);
      }
    }
    win.anchorCell = { rownum: win._displayRows[fromDi]._rownum, col: win._columns[0] };
    win._copyWithHeader = true;
    applyCellHighlights(win);
  }

  async function pasteAtAnchor(win) {
    if (!win.anchorCell || !win.tableName) return;
    const t = tables[win.tableName];
    if (!t) return;
    let text;
    try { text = await navigator.clipboard.readText(); } catch (_) { return; }
    if (!text) return;
    if (!tables[win.tableName] || !win._displayRows) return;
    const lines = text.replace(/\n$/, '').split('\n');
    const pasteData = lines.map(line => line.split('\t'));
    const startDi = win._displayRows.findIndex(r => r._rownum === win.anchorCell.rownum);
    const startCi = win._columns.indexOf(win.anchorCell.col);
    if (startDi < 0 || startCi < 0) return;
    const changes = [];
    for (let ri = 0; ri < pasteData.length; ri++) {
      const di = startDi + ri;
      if (di >= win._displayRows.length) break;
      const row = win._displayRows[di];
      for (let ci = 0; ci < pasteData[ri].length; ci++) {
        const colIdx = startCi + ci;
        if (colIdx >= win._columns.length) break;
        const col = win._columns[colIdx];
        const oldValue = String(row[col] ?? '');
        const newValue = pasteData[ri][ci];
        if (oldValue !== newValue) {
          row[col] = newValue;
          changes.push({ rownum: row._rownum, col, oldValue, newValue });
          if (!t._dirtyCells) t._dirtyCells = [];
          t._dirtyCells.push({ rownum: row._rownum, col, value: newValue });
        }
      }
    }
    if (changes.length === 0) return;
    if (!t._undoStack) t._undoStack = [];
    t._undoStack.push({ type: 'paste', changes });
    t._redoStack = [];
    markModified(win.tableName);
    debouncedSync(win.tableName);
    const endDi = Math.min(startDi + pasteData.length - 1, win._displayRows.length - 1);
    const endCi = Math.min(startCi + Math.max(...pasteData.map(r => r.length)) - 1, win._columns.length - 1);
    win.selectedCells = new Set();
    for (let di = startDi; di <= endDi; di++) {
      const rn = win._displayRows[di]._rownum;
      for (let ci = startCi; ci <= endCi; ci++) {
        win.selectedCells.add(`${rn}:${win._columns[ci]}`);
      }
    }
    windows.filter(w => w.tableName === win.tableName).forEach(w => {
      w._renderStart = -1; w._renderEnd = -1;
      renderVisibleRows(w);
    });
    refocusAnchorCell(win);
  }

  function undoTable(tableName, win) {
    const t = tables[tableName];
    if (!t) return;
    if (!t._undoStack || t._undoStack.length === 0) return;
    const entry = t._undoStack.pop();
    const rowMap = new Map(t.rows.map(r => [r._rownum, r]));
    if (!t._dirtyCells) t._dirtyCells = [];
    for (const { rownum, col, oldValue } of entry.changes) {
      const row = rowMap.get(rownum);
      if (!row || !t.columns.includes(col)) continue;
      row[col] = oldValue;
      t._dirtyCells.push({ rownum, col, value: oldValue });
    }
    if (!t._redoStack) t._redoStack = [];
    t._redoStack.push(entry);
    markModified(tableName);
    debouncedSync(tableName);
    windows.filter(w => w.tableName === tableName).forEach(w => {
      w._renderStart = -1; w._renderEnd = -1;
      renderVisibleRows(w);
    });
    refocusAnchorCell(win);
  }

  function redoTable(tableName, win) {
    const t = tables[tableName];
    if (!t) return;
    if (!t._redoStack || t._redoStack.length === 0) return;
    const entry = t._redoStack.pop();
    const rowMap = new Map(t.rows.map(r => [r._rownum, r]));
    if (!t._dirtyCells) t._dirtyCells = [];
    for (const { rownum, col, newValue } of entry.changes) {
      const row = rowMap.get(rownum);
      if (!row || !t.columns.includes(col)) continue;
      row[col] = newValue;
      t._dirtyCells.push({ rownum, col, value: newValue });
    }
    if (!t._undoStack) t._undoStack = [];
    t._undoStack.push(entry);
    markModified(tableName);
    debouncedSync(tableName);
    windows.filter(w => w.tableName === tableName).forEach(w => {
      w._renderStart = -1; w._renderEnd = -1;
      renderVisibleRows(w);
    });
    refocusAnchorCell(win);
  }

  function moveSelectionColumns(win, tr, td, key) {
    const t = tables[win.tableName];
    if (!t) return;
    const leadCi = parseInt(td.dataset.colIdx, 10);
    const displayIdx = parseInt(tr.dataset.displayIdx, 10);
    if (isNaN(leadCi) || isNaN(displayIdx)) return;
    const anchorCi = win.anchorCell ? t.columns.indexOf(win.anchorCell.col) : leadCi;
    const c1 = Math.min(anchorCi, leadCi);
    const c2 = Math.max(anchorCi, leadCi);
    const focusedCol = win._columns[leadCi];
    const refocus = () => {
      const newCi = tables[win.tableName].columns.indexOf(focusedCol);
      if (newCi >= 0) focusCellAt(win, displayIdx, newCi);
    };
    if (key === 'ArrowLeft') {
      if (c1 === 0) return;
      reorderColumn(win, c1 - 1, c2 + 1).then(refocus);
    } else {
      if (c2 >= t.columns.length - 1) return;
      reorderColumn(win, c2 + 1, c1).then(refocus);
    }
  }

  function extendCellSelection(win, tr, td, key) {
    const displayIdx = parseInt(tr.dataset.displayIdx, 10);
    const colIdx = parseInt(td.dataset.colIdx, 10);
    if (isNaN(displayIdx) || isNaN(colIdx)) return;
    const maxDisplay = win._displayRows.length - 1;
    const maxCol = win._columns.length - 1;
    let newDisplay = displayIdx;
    let newCol = colIdx;
    if (key === 'ArrowUp') newDisplay = Math.max(0, displayIdx - 1);
    else if (key === 'ArrowDown') newDisplay = Math.min(maxDisplay, displayIdx + 1);
    else if (key === 'ArrowLeft') newCol = Math.max(0, colIdx - 1);
    else if (key === 'ArrowRight') newCol = Math.min(maxCol, colIdx + 1);
    if (newDisplay === displayIdx && newCol === colIdx) return; // at edge
    rebuildSelectionRect(win, newDisplay, newCol);
    focusCellAt(win, newDisplay, newCol);
    applyCellHighlights(win);
  }

  function moveSingleCellSelection(win, tr, td, key) {
    const displayIdx = parseInt(tr.dataset.displayIdx, 10);
    const colIdx = parseInt(td.dataset.colIdx, 10);
    if (isNaN(displayIdx) || isNaN(colIdx)) return;
    const maxDisplay = win._displayRows.length - 1;
    const maxCol = win._columns.length - 1;
    let newDisplay = displayIdx;
    let newCol = colIdx;
    if (key === 'ArrowUp') newDisplay = Math.max(0, displayIdx - 1);
    else if (key === 'ArrowDown') newDisplay = Math.min(maxDisplay, displayIdx + 1);
    else if (key === 'ArrowLeft') newCol = Math.max(0, colIdx - 1);
    else if (key === 'ArrowRight') newCol = Math.min(maxCol, colIdx + 1);
    if (newDisplay === displayIdx && newCol === colIdx) return; // at edge
    const row = win._displayRows[newDisplay];
    const col = win._columns[newCol];
    if (!row || col == null) return;
    win.anchorCell = { rownum: row._rownum, col };
    win.selectedCells = new Set([`${row._rownum}:${col}`]);
    focusCellAt(win, newDisplay, newCol);
    applyCellHighlights(win);
  }

  function focusCellAt(win, displayIdx, colIdx) {
    const container = win._container;
    const thead = win._table?.querySelector('thead');
    if (!container) return;
    const theadH = thead ? thead.offsetHeight : 0;
    const cellTop = displayIdx * ROW_HEIGHT;
    const cellBottom = cellTop + ROW_HEIGHT;
    const viewTop = Math.max(0, container.scrollTop - theadH);
    const viewBottom = viewTop + container.clientHeight - theadH;
    if (cellTop < viewTop) {
      container.scrollTop = cellTop;
    } else if (cellBottom > viewBottom) {
      container.scrollTop = cellBottom - (container.clientHeight - theadH);
    }
    // Force render of the target cell before focusing
    win._renderStart = -1; win._renderEnd = -1;
    renderVisibleRows(win);
    const tr = win._tbody.querySelector(`tr[data-display-idx="${displayIdx}"]`);
    if (!tr) return;
    const td = tr.children[colIdx + 1]; // +1 for row-num column
    if (!td) return;
    win._programmaticFocus = true;
    td.focus();
    win._programmaticFocus = false;
  }

  function cycleTableWindow(currentWin, dir) {
    const list = windows.filter(w =>
      w.tableName && tables[w.tableName] && !w.el.classList.contains('minimized')
    );
    if (list.length < 2) return false;
    const idx = list.indexOf(currentWin);
    if (idx === -1) return false;
    const next = list[(idx + dir + list.length) % list.length];
    focusWindow(next.id);
    const dIdx = next.anchorCell
      ? next._displayRows.findIndex(r => r._rownum === next.anchorCell.rownum)
      : -1;
    const cIdx = next.anchorCell ? next._columns.indexOf(next.anchorCell.col) : -1;
    if (dIdx >= 0 && cIdx >= 0) focusCellAt(next, dIdx, cIdx);
    else focusMiddleCell(next);
    return true;
  }

  function nudgeWindow(win, dx, dy) {
    const left = parseInt(win.el.style.left) || 0;
    const top = parseInt(win.el.style.top) || 0;
    const area = document.getElementById('window-area');
    win.el.style.left = Math.max(0, Math.min(left + dx, area.clientWidth - win.el.offsetWidth)) + 'px';
    win.el.style.top = Math.max(0, Math.min(top + dy, area.clientHeight - win.el.offsetHeight)) + 'px';
    if (win.maximized) win.maximized = false;
  }

  function applyCellHighlights(win) {
    const tbody = win._tbody;
    const table = win._table;
    if (!tbody || !table) return;
    const thead = table.querySelector('thead tr');
    const selected = win.selectedCells || new Set();
    const hiRows = new Set();
    const hiCols = new Set();
    for (const key of selected) {
      const sep = key.indexOf(':');
      if (sep < 0) continue;
      hiRows.add(parseInt(key.slice(0, sep), 10));
      hiCols.add(key.slice(sep + 1));
    }
    if (thead) {
      for (const th of thead.querySelectorAll('th')) {
        const ci = parseInt(th.dataset.colIdx, 10);
        if (isNaN(ci)) continue;
        const name = win._columns[ci];
        th.classList.toggle('col-highlight', hiCols.has(name));
      }
    }
    for (const tr of tbody.querySelectorAll('tr')) {
      if (tr.classList.contains('virtual-pad')) continue;
      const di = parseInt(tr.dataset.displayIdx, 10);
      if (isNaN(di)) continue;
      const row = win._displayRows[di];
      if (!row) continue;
      tr.classList.toggle('row-highlight', hiRows.has(row._rownum));
      for (const td of tr.children) {
        const ci = parseInt(td.dataset.colIdx, 10);
        if (isNaN(ci)) continue;
        const name = win._columns[ci];
        td.classList.toggle('col-highlight', hiCols.has(name));
        td.classList.toggle('cell-selected', selected.has(`${row._rownum}:${name}`));
      }
    }
  }

  function rebuildTable(win) {
    const t = tables[win.tableName];
    if (!t) return;
    const container = win.el.querySelector('.table-container');
    if (!container) return;

    const oldFilter = win._lastFilterText;
    const filterChanged = oldFilter !== undefined && oldFilter !== win.filterText;
    win._lastFilterText = win.filterText;

    const cfKey = JSON.stringify(Object.keys(win.columnFilters).sort().map(k => [k, [...win.columnFilters[k]].sort()]));
    const colFilterChanged = win._lastColFilterKey !== undefined && win._lastColFilterKey !== cfKey;
    win._lastColFilterKey = cfKey;

    const savedScrollLeft = container.scrollLeft;
    buildTableHTML(win, container, t);

    container.scrollLeft = savedScrollLeft;
    if (filterChanged || colFilterChanged) {
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

  async function addColumn(tableName, colName) {
    const t = tables[tableName];
    if (t.columns.includes(colName)) return;
    t.columns.push(colName);
    t.rows.forEach(r => { r[colName] = ''; });
    markModified(tableName);
    windows.filter(w => w.tableName === tableName).forEach(w => { w.colWidths = null; });
    await registerTable(tableName);
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
    if (win.selectedCol === oldCol) win.selectedCol = newCol;
    if (win.columnFilters[oldCol]) {
      win.columnFilters[newCol] = win.columnFilters[oldCol];
      delete win.columnFilters[oldCol];
    }
    markModified(tableName);
    try { db.run(`ALTER TABLE [${tableName}] RENAME COLUMN [${oldCol}] TO [${newCol}]`); } catch (_) {}
    rebuildTransformCacheForTable(tableName);
    rebuildTable(win);
  }

  function updateTableWidth(win) {
    if (!win._table || !win.colWidths) return;
    win._table.style.width = (50 + win.colWidths.reduce((a, b) => a + b, 0)) + 'px';
  }

  function startColResize(win, colIdx, e) {
    closeAutoFilter();
    const colEl = win._colgroup.children[colIdx + 1];
    const startX = e.clientX;
    const startWidth = win.colWidths[colIdx];
    const MIN_COL_WIDTH = 40;

    document.body.classList.add('col-resizing');
    const th = win._table.querySelector(`thead th[data-col-idx="${colIdx}"]`);
    const handle = th && th.querySelector('.col-resize-handle');
    if (handle) handle.classList.add('active');

    const onMove = (me) => {
      const newWidth = Math.max(MIN_COL_WIDTH, startWidth + me.clientX - startX);
      win.colWidths[colIdx] = newWidth;
      colEl.style.width = newWidth + 'px';
      updateTableWidth(win);
    };
    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      document.body.classList.remove('col-resizing');
      if (handle) handle.classList.remove('active');
      if (th) th._didDrag = true;
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  }

  function autoFitColumn(win, colIdx) {
    const MIN_COL_WIDTH = 40;
    const MAX_COL_WIDTH = 600;
    const measurer = document.createElement('span');
    measurer.style.cssText = 'visibility:hidden;position:absolute;white-space:nowrap;padding:0 8px;font-size:12px;font-family:inherit;';
    document.body.appendChild(measurer);

    const col = win._columns[colIdx];
    measurer.textContent = col;
    let maxW = measurer.offsetWidth + 40;

    const displayRows = win._displayRows;
    const start = win._renderStart || 0;
    const end = Math.min(win._renderEnd || displayRows.length, displayRows.length);
    for (let i = start; i < end; i++) {
      measurer.textContent = String(displayRows[i][col] ?? '');
      maxW = Math.max(maxW, measurer.offsetWidth);
    }
    measurer.remove();

    const finalWidth = Math.min(MAX_COL_WIDTH, Math.max(MIN_COL_WIDTH, maxW));
    win.colWidths[colIdx] = finalWidth;
    win._colgroup.children[colIdx + 1].style.width = finalWidth + 'px';
    updateTableWidth(win);
  }

  async function reorderColumn(win, fromIdx, toIdx) {
    const t = tables[win.tableName];
    if (!t) return;
    const col = t.columns.splice(fromIdx, 1)[0];
    const w = win.colWidths ? win.colWidths.splice(fromIdx, 1)[0] : null;
    if (toIdx > fromIdx) toIdx--;
    t.columns.splice(toIdx, 0, col);
    if (win.colWidths && w != null) win.colWidths.splice(toIdx, 0, w);
    markModified(win.tableName);
    await registerTable(win.tableName);
    rebuildTable(win);
  }

  function startColumnRename(win, th, oldCol) {
    const currentWidth = th.offsetWidth;
    const input = document.createElement('input');
    input.className = 'inline-rename';
    input.value = oldCol;
    input.style.width = currentWidth + 'px';
    input.style.boxSizing = 'border-box';
    th.innerHTML = '';
    th.appendChild(input);
    th._renaming = true;
    input.focus();
    input.select();

    let done = false;
    function finish() {
      if (done) return;
      done = true;
      th._renaming = false;
      const raw = input.value.trim();
      if (input.parentNode) input.remove();
      if (!raw || raw === oldCol) {
        rebuildTable(win);
        return;
      }
      renameColumn(win.tableName, oldCol, raw, win);
    }
    input.addEventListener('blur', finish);
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); input.blur(); }
      if (e.key === 'Escape') { done = true; th._renaming = false; input.remove(); rebuildTable(win); }
      e.stopPropagation();
    });
    input.addEventListener('click', (e) => e.stopPropagation());
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

  // ---- AutoFilter Dropdown ----

  function closeAutoFilter() {
    if (!_activeAutoFilter) return;
    _activeAutoFilter.el.remove();
    document.removeEventListener('mousedown', _autoFilterOutsideClick);
    document.removeEventListener('keydown', _autoFilterEscapeKey);
    window.removeEventListener('resize', closeAutoFilter);
    if (_activeAutoFilter.scrollHandler) {
      const container = _activeAutoFilter.win.el.querySelector('.table-container');
      if (container) container.removeEventListener('scroll', _activeAutoFilter.scrollHandler);
    }
    _activeAutoFilter = null;
  }

  function _autoFilterOutsideClick(e) {
    if (_activeAutoFilter && !_activeAutoFilter.el.contains(e.target)) {
      closeAutoFilter();
    }
  }

  function _autoFilterEscapeKey(e) {
    if (e.key === 'Escape' && _activeAutoFilter) {
      closeAutoFilter();
    }
  }

  function openAutoFilter(win, col, anchorEl) {
    if (_activeAutoFilter && _activeAutoFilter.win === win && _activeAutoFilter.col === col) {
      closeAutoFilter();
      return;
    }
    closeAutoFilter();

    const t = tables[win.tableName];
    if (!t) return;

    // Gather unique values from all rows (not just filtered)
    const useTransform = !win.disabledTransforms.has(col) && hasDisplayTransform(win.tableName, col);
    const seen = new Set();
    const values = [];
    for (const row of t.rows) {
      const v = String(row[col] ?? '');
      if (!seen.has(v)) {
        seen.add(v);
        values.push({ raw: v, display: useTransform ? String(getDisplayValue(win.tableName, col, row)) : v });
      }
    }
    values.sort((a, b) => collator.compare(a.display, b.display));

    const existingFilter = win.columnFilters[col];
    const items = values.map(v => ({
      value: v.raw,
      display: v.display,
      checked: existingFilter ? existingFilter.has(v.raw) : true,
    }));

    // Build dropdown DOM
    const dropdown = document.createElement('div');
    dropdown.className = 'autofilter-dropdown';

    // Search
    const searchWrap = document.createElement('div');
    searchWrap.className = 'autofilter-search-wrap';
    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.className = 'autofilter-search';
    searchInput.placeholder = 'Search values...';
    searchWrap.appendChild(searchInput);
    dropdown.appendChild(searchWrap);

    // Select All
    const controls = document.createElement('div');
    controls.className = 'autofilter-controls';
    const selectAllLabel = document.createElement('label');
    selectAllLabel.className = 'autofilter-select-all';
    const selectAllCb = document.createElement('input');
    selectAllCb.type = 'checkbox';
    selectAllLabel.appendChild(selectAllCb);
    selectAllLabel.appendChild(document.createTextNode(' Select All'));
    controls.appendChild(selectAllLabel);
    dropdown.appendChild(controls);

    // Value list
    const listEl = document.createElement('div');
    listEl.className = 'autofilter-list';
    dropdown.appendChild(listEl);

    let searchText = '';
    let visibleItems = items;

    function getVisibleItems() {
      if (!searchText) return items;
      const lower = searchText.toLowerCase();
      return items.filter(it => it.display.toLowerCase().includes(lower));
    }

    function updateSelectAll() {
      const vis = getVisibleItems();
      const allChecked = vis.length > 0 && vis.every(it => it.checked);
      const someChecked = vis.some(it => it.checked);
      selectAllCb.checked = allChecked;
      selectAllCb.indeterminate = someChecked && !allChecked;
    }

    function renderList() {
      visibleItems = getVisibleItems();
      listEl.innerHTML = '';

      if (visibleItems.length <= 200) {
        for (const item of visibleItems) {
          const label = document.createElement('label');
          label.className = 'autofilter-item';
          const cb = document.createElement('input');
          cb.type = 'checkbox';
          cb.checked = item.checked;
          cb.addEventListener('change', () => {
            item.checked = cb.checked;
            updateSelectAll();
          });
          const txt = document.createElement('span');
          txt.textContent = item.display === '' ? '(empty)' : item.display;
          label.appendChild(cb);
          label.appendChild(txt);
          listEl.appendChild(label);
        }
      } else {
        // Virtual scrolling for large lists
        const ITEM_H = 24;
        const spacer = document.createElement('div');
        spacer.style.height = (visibleItems.length * ITEM_H) + 'px';
        spacer.style.position = 'relative';
        listEl.appendChild(spacer);

        const renderVirtual = () => {
          const st = listEl.scrollTop;
          const ch = listEl.clientHeight;
          const start = Math.max(0, Math.floor(st / ITEM_H) - 5);
          const end = Math.min(visibleItems.length, Math.ceil((st + ch) / ITEM_H) + 5);
          while (spacer.firstChild) spacer.removeChild(spacer.firstChild);
          for (let i = start; i < end; i++) {
            const item = visibleItems[i];
            const label = document.createElement('label');
            label.className = 'autofilter-item';
            label.style.position = 'absolute';
            label.style.top = (i * ITEM_H) + 'px';
            label.style.left = '0';
            label.style.right = '0';
            label.style.height = ITEM_H + 'px';
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.checked = item.checked;
            cb.addEventListener('change', () => {
              item.checked = cb.checked;
              updateSelectAll();
            });
            const txt = document.createElement('span');
            txt.textContent = item.display === '' ? '(empty)' : item.display;
            label.appendChild(cb);
            label.appendChild(txt);
            spacer.appendChild(label);
          }
        };
        renderVirtual();
        listEl.addEventListener('scroll', () => requestAnimationFrame(renderVirtual));
      }
      updateSelectAll();
    }

    selectAllCb.addEventListener('change', () => {
      const checked = selectAllCb.checked;
      for (const it of getVisibleItems()) it.checked = checked;
      renderList();
    });

    let searchTimeout;
    searchInput.addEventListener('input', () => {
      clearTimeout(searchTimeout);
      searchTimeout = setTimeout(() => {
        searchText = searchInput.value.trim();
        renderList();
      }, 150);
    });

    // Buttons
    const buttons = document.createElement('div');
    buttons.className = 'autofilter-buttons';
    const clearBtn = document.createElement('button');
    clearBtn.className = 'autofilter-clear';
    clearBtn.textContent = 'Clear';
    clearBtn.addEventListener('click', () => {
      delete win.columnFilters[col];
      closeAutoFilter();
      rebuildTable(win);
    });
    const applyBtn = document.createElement('button');
    applyBtn.className = 'autofilter-apply';
    applyBtn.textContent = 'Apply';
    applyBtn.addEventListener('click', () => {
      const checked = new Set();
      for (const it of items) {
        if (it.checked) checked.add(it.value);
      }
      if (checked.size === items.length) {
        delete win.columnFilters[col];
      } else {
        win.columnFilters[col] = checked;
      }
      closeAutoFilter();
      rebuildTable(win);
    });
    buttons.appendChild(clearBtn);
    buttons.appendChild(applyBtn);
    dropdown.appendChild(buttons);

    renderList();
    document.body.appendChild(dropdown);

    // Position below the anchor header cell, right-aligned to its right edge
    const rect = anchorEl.getBoundingClientRect();
    let top = rect.bottom + 2;
    const dRect = dropdown.getBoundingClientRect();
    let left = rect.right - dRect.width;
    if (top + dRect.height > window.innerHeight) top = rect.top - dRect.height - 2;
    if (left < 0) left = 4;
    if (left + dRect.width > window.innerWidth) left = window.innerWidth - dRect.width - 8;
    dropdown.style.top = top + 'px';
    dropdown.style.left = left + 'px';

    searchInput.focus();

    // Close on outside click or Escape
    const scrollHandler = () => closeAutoFilter();
    const container = win.el.querySelector('.table-container');
    if (container) container.addEventListener('scroll', scrollHandler);

    _activeAutoFilter = { win, col, el: dropdown, scrollHandler };
    setTimeout(() => {
      document.addEventListener('mousedown', _autoFilterOutsideClick);
      document.addEventListener('keydown', _autoFilterEscapeKey);
      window.addEventListener('resize', closeAutoFilter);
    }, 0);
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

  // Prompt history (Up/Down arrow to navigate previous commands)
  const _aiHistory = [];
  let _aiHistoryIdx = _aiHistory.length;
  let _aiHistoryDraft = '';
  const MAX_HISTORY = 100;

  function pushHistory(arr, value) {
    if (!value) return;
    const idx = arr.indexOf(value);
    if (idx !== -1) arr.splice(idx, 1);
    arr.push(value);
    if (arr.length > MAX_HISTORY) arr.shift();
  }

  function handleHistoryKey(e, history, getIdx, setIdx, getDraft, setDraft) {
    const el = e.target;
    if (e.key === 'ArrowUp') {
      // Only navigate history when cursor is at the very start
      if (el.selectionStart !== 0 || el.selectionEnd !== 0) return;
      if (history.length === 0) return;
      e.preventDefault();
      // Save draft on first up
      if (getIdx() === history.length) setDraft(el.value);
      if (getIdx() > 0) {
        setIdx(getIdx() - 1);
        el.value = history[getIdx()];
        el.setSelectionRange(0, 0);
      }
    } else if (e.key === 'ArrowDown') {
      // Only navigate history when cursor is at the very end
      if (el.selectionStart !== el.value.length || el.selectionEnd !== el.value.length) return;
      if (getIdx() >= history.length) return;
      e.preventDefault();
      setIdx(getIdx() + 1);
      if (getIdx() === history.length) {
        el.value = getDraft();
      } else {
        el.value = history[getIdx()];
      }
      el.setSelectionRange(el.value.length, el.value.length);
    }
  }

  function setupKeyboard() {
    document.getElementById('sql-input').addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        executeQuery();
      }
    });

    document.getElementById('ai-input').addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        runAI();
        return;
      }
      handleHistoryKey(e, _aiHistory,
        () => _aiHistoryIdx, v => _aiHistoryIdx = v,
        () => _aiHistoryDraft, v => _aiHistoryDraft = v);
    });

    document.addEventListener('keydown', (e) => {
      // Plain arrow key on an active data window with no cell selected: focus the cell in the middle of the view.
      if (!e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey &&
          (e.key === 'ArrowUp' || e.key === 'ArrowDown' || e.key === 'ArrowLeft' || e.key === 'ArrowRight')) {
        const tgt = e.target;
        if (tgt && tgt.matches && tgt.matches('input, textarea, [contenteditable="true"]')) return;
        const win = getActiveDataWindow();
        if (!win || win.anchorCell) return;
        if (focusMiddleCell(win)) e.preventDefault();
        return;
      }
      if (!(e.ctrlKey || e.metaKey)) return;
      if (e.shiftKey && e.key.toLowerCase() === 'a') {
        const tgt = e.target;
        if (tgt && tgt.matches && tgt.matches('input, textarea, [contenteditable="true"], td.data-cell')) return;
        const win = getActiveDataWindow();
        if (win && win.tableName) { e.preventDefault(); selectAllCells(win); return; }
      }
      switch (e.key) {
        case 's': e.preventDefault(); saveActiveTable(); break;
        case 'o': e.preventDefault(); openFile(e.shiftKey); break;
        case 'n': e.preventDefault(); newTable(); break;
        case 'w': e.preventDefault(); if (activeWinId) closeWindow(activeWinId); break;
        case 'ArrowLeft':
        case 'ArrowRight': {
          const tgt = e.target;
          if (tgt && tgt.matches && tgt.matches('input, textarea, [contenteditable="true"], td.data-cell')) return;
          const win = getActiveDataWindow();
          if (!win || !win.selectedCol) return;
          const t = tables[win.tableName];
          if (!t) return;
          const idx = t.columns.indexOf(win.selectedCol);
          if (idx === -1) return;
          const dir = e.key === 'ArrowLeft' ? -1 : 1;
          const target = idx + dir;
          if (target < 0 || target >= t.columns.length) return;
          e.preventDefault();
          const toIdx = dir === -1 ? target : target + 1;
          reorderColumn(win, idx, toIdx);
          break;
        }
      }
    });
  }

  // SQL syntax highlighting
  const _sqlKeywords = new Set([
    'SELECT','FROM','WHERE','INSERT','UPDATE','DELETE','INTO','JOIN','ON','AND','OR','NOT','IN',
    'LIKE','AS','ORDER','BY','GROUP','HAVING','LIMIT','OFFSET','UNION','ALL','DISTINCT','BETWEEN',
    'EXISTS','IS','NULL','CREATE','TABLE','DROP','ALTER','SET','VALUES','COUNT','SUM','AVG','MIN',
    'MAX','CASE','WHEN','THEN','ELSE','END','ASC','DESC','LEFT','RIGHT','INNER','OUTER','CROSS',
    'NATURAL','FULL','USING','WITH','RECURSIVE','CAST','OVER','PARTITION','WINDOW','ROWS','RANGE',
    'UNBOUNDED','PRECEDING','FOLLOWING','CURRENT','ROW','IF','REPLACE','BEGIN','COMMIT','ROLLBACK',
    'TRANSACTION','INDEX','PRIMARY','KEY','FOREIGN','REFERENCES','CONSTRAINT','DEFAULT','CHECK',
    'UNIQUE','AUTOINCREMENT','TEMP','TEMPORARY','VIEW','TRIGGER','EXPLAIN','PRAGMA','VACUUM',
    'ATTACH','DETACH','REINDEX','ANALYZE','GLOB','REGEXP','ESCAPE','COLLATE','NOCASE','ABORT',
    'FAIL','IGNORE','CONFLICT','INSTEAD','OF','EACH','FOR','AFTER','BEFORE','NO','ACTION',
    'CASCADE','RESTRICT','DEFERRABLE','DEFERRED','IMMEDIATE','INITIALLY','RAISE','EXCEPT',
    'INTERSECT','COALESCE','IFNULL','NULLIF','TYPEOF','LENGTH','SUBSTR','UPPER','LOWER','TRIM',
    'LTRIM','RTRIM','INSTR','HEX','QUOTE','RANDOM','ABS','ROUND','TOTAL','GROUP_CONCAT',
    'REPLACE','ZEROBLOB','UNICODE','CHAR','PRINTF','FORMAT','DATE','TIME','DATETIME','STRFTIME',
    'JULIANDAY','NOT','BOOLEAN','TRUE','FALSE','INTEGER','REAL','TEXT','BLOB','NUMERIC',
  ]);

  function sqlHighlightHTML(text) {
    let out = '';
    let i = 0;
    const len = text.length;
    while (i < len) {
      const ch = text[i];
      // Whitespace
      if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
        out += ch;
        i++;
        continue;
      }
      // Single-line comment --
      if (ch === '-' && i + 1 < len && text[i + 1] === '-') {
        let end = text.indexOf('\n', i);
        if (end === -1) end = len;
        out += '<span class="sql-cmt">' + escHtml(text.slice(i, end)) + '</span>';
        i = end;
        continue;
      }
      // Block comment /* */
      if (ch === '/' && i + 1 < len && text[i + 1] === '*') {
        let end = text.indexOf('*/', i + 2);
        if (end === -1) end = len; else end += 2;
        out += '<span class="sql-cmt">' + escHtml(text.slice(i, end)) + '</span>';
        i = end;
        continue;
      }
      // Single-quoted string
      if (ch === "'") {
        let j = i + 1;
        while (j < len) {
          if (text[j] === "'" && j + 1 < len && text[j + 1] === "'") { j += 2; continue; }
          if (text[j] === "'") { j++; break; }
          j++;
        }
        out += '<span class="sql-str">' + escHtml(text.slice(i, j)) + '</span>';
        i = j;
        continue;
      }
      // Bracket-quoted identifier [...]
      if (ch === '[') {
        let end = text.indexOf(']', i + 1);
        if (end === -1) end = len - 1;
        out += '<span class="sql-brk">' + escHtml(text.slice(i, end + 1)) + '</span>';
        i = end + 1;
        continue;
      }
      // Numbers
      if ((ch >= '0' && ch <= '9') || (ch === '.' && i + 1 < len && text[i + 1] >= '0' && text[i + 1] <= '9')) {
        let j = i;
        while (j < len && ((text[j] >= '0' && text[j] <= '9') || text[j] === '.')) j++;
        out += '<span class="sql-num">' + escHtml(text.slice(i, j)) + '</span>';
        i = j;
        continue;
      }
      // Words (identifiers/keywords)
      if ((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || ch === '_') {
        let j = i + 1;
        while (j < len && ((text[j] >= 'a' && text[j] <= 'z') || (text[j] >= 'A' && text[j] <= 'Z') || (text[j] >= '0' && text[j] <= '9') || text[j] === '_')) j++;
        const word = text.slice(i, j);
        if (_sqlKeywords.has(word.toUpperCase())) {
          out += '<span class="sql-kw">' + escHtml(word) + '</span>';
        } else {
          out += escHtml(word);
        }
        i = j;
        continue;
      }
      // Operators and other characters
      out += escHtml(ch);
      i++;
    }
    return out;
  }

  function setupSQLHighlight() {
    const input = document.getElementById('sql-input');
    const overlay = document.getElementById('sql-highlight');
    function update() {
      overlay.innerHTML = sqlHighlightHTML(input.value) + '\n';
    }
    input.addEventListener('input', update);
    input.addEventListener('scroll', () => {
      overlay.scrollTop = input.scrollTop;
      overlay.scrollLeft = input.scrollLeft;
    });
    update();
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

  // Worker-based query execution for interruptibility and live timer
  let _queryWorker = null;
  let _queryTimer = null;

  function makeQueryWorker() {
    const libBase = new URL('lib/', location.href).href;
    const src = `
      importScripts('${libBase}sql-wasm.js');
      let SQL = null;
      let db = null;
      initSqlJs({ locateFile: f => '${libBase}' + f }).then((sql) => {
        SQL = sql;
        postMessage({ type: 'ready' });
      });
      onmessage = (e) => {
        if (e.data.type === 'load') {
          try {
            const newDb = new SQL.Database(e.data.data);
            newDb.create_function('regexp', (pattern, value) => {
              try { return new RegExp(pattern, 'i').test(value) ? 1 : 0; } catch (_) { return 0; }
            });
            if (db) db.close();
            db = newDb;
            postMessage({ type: 'loaded' });
          } catch (err) {
            postMessage({ type: 'error', message: err.message });
          }
        } else if (e.data.type === 'exec') {
          try {
            const result = db.exec(e.data.sql);
            const tablesResult = db.exec("SELECT name FROM sqlite_master WHERE type='table'");
            const tableNames = tablesResult.length ? tablesResult[0].values.map(r => r[0]) : [];
            postMessage({ type: 'result', result, tableNames });
          } catch (err) {
            postMessage({ type: 'error', message: err.message });
          }
        }
      };
    `;
    const blob = new Blob([src], { type: 'application/javascript' });
    return new Worker(URL.createObjectURL(blob));
  }

  function cancelQuery() {
    if (_queryWorker) {
      _queryWorker.terminate();
      _queryWorker = null;
    }
    if (_queryTimer) {
      clearInterval(_queryTimer);
      _queryTimer = null;
    }
    setStatus('Query interrupted', 'error');
    // Restore interrupt button back to normal
    const btn = document.getElementById('btn-interrupt');
    if (btn) btn.style.display = 'none';
  }

  function executeQuery() {
    const input = document.getElementById('sql-input');
    const sql = autoQuoteSQL(input.value.trim());
    if (!sql) return;

    flushAllSyncs();

    // Handle SELECT ... INTO ... by stripping INTO before sending to worker
    const intoInfo = extractIntoClause(sql);
    const workerSQL = intoInfo ? intoInfo.selectSQL : sql;

    // Export current database state for the worker
    const dbData = db.export();
    registerDBFunctions(); // db.export() destroys custom functions in sql.js
    const t0 = performance.now();

    // Show timer and interrupt button
    setStatus('Running query... 0s', 'working');
    showInterruptButton(true);
    _queryTimer = setInterval(() => {
      const elapsed = ((performance.now() - t0) / 1000).toFixed(1);
      setStatus(`Running query... ${elapsed}s`, 'working');
    }, 100);

    // Create worker and run query
    _queryWorker = makeQueryWorker();
    _queryWorker.onmessage = async (e) => {
      if (e.data.type === 'ready') {
        _queryWorker.postMessage({ type: 'load', data: dbData });
      } else if (e.data.type === 'loaded') {
        _queryWorker.postMessage({ type: 'exec', sql: workerSQL });
      } else if (e.data.type === 'result') {
        clearInterval(_queryTimer);
        _queryTimer = null;
        showInterruptButton(false);
        const elapsed = performance.now() - t0;
        try {
          await handleQueryResult(sql, e.data.result, e.data.tableNames, elapsed, intoInfo);
        } catch (err) {
          setStatus(`Error: ${err.message}`, 'error');
        }
        _queryWorker.terminate();
        _queryWorker = null;
      } else if (e.data.type === 'error') {
        clearInterval(_queryTimer);
        _queryTimer = null;
        showInterruptButton(false);
        setStatus(`Error: ${e.data.message}`, 'error');
        _queryWorker.terminate();
        _queryWorker = null;
      }
    };
    _queryWorker.onerror = (err) => {
      clearInterval(_queryTimer);
      _queryTimer = null;
      showInterruptButton(false);
      setStatus(`Worker error: ${err.message}`, 'error');
      _queryWorker.terminate();
      _queryWorker = null;
    };
  }

  function showInterruptButton(show, handler) {
    let btn = document.getElementById('btn-interrupt');
    if (!btn) {
      btn = document.createElement('button');
      btn.id = 'btn-interrupt';
      document.getElementById('console-actions').appendChild(btn);
    }
    btn.textContent = 'Interrupt';
    btn.onclick = handler || cancelQuery;
    btn.style.display = show ? '' : 'none';
  }

  async function handleQueryResult(sql, result, workerTableNames, elapsed, intoInfo) {
    // Handle SELECT ... INTO ...
    if (intoInfo) {
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
      await registerTable(uniqueName);
      createTableWindow(uniqueName);
      setStatus(`Created table "${uniqueName}" with ${tableRows.length} row(s) in ${formatElapsed(elapsed)}`, 'success');
      return;
    }

    // Detect new tables: compare worker's table list with our known tables
    const tablesBefore = new Set(getDBTables());

    // Re-run DDL statements on main db so tables persist
    const isDDL = /^\s*(CREATE|DROP|ALTER|INSERT|UPDATE|DELETE|REPLACE)\b/i.test(sql);
    if (isDDL) {
      try { db.exec(sql); } catch (e) {}
    }

    const newTables = workerTableNames.filter(n => !tablesBefore.has(n) && !tables[n]);

    const createMatch = sql.match(/\bCREATE\s+TABLE\s+(?:IF\s+NOT\s+EXISTS\s+)?\[?([^\]\s,(;]+)\]?/i);
    if (createMatch) {
      const createName = createMatch[1];
      if (!newTables.includes(createName) && !tables[createName]) {
        const dbTables = getDBTables();
        if (dbTables.includes(createName)) newTables.push(createName);
      }
    }

    if (newTables.length > 0) {
      await importNewDBTables(newTables);
      setStatus(`Created table(s): ${newTables.join(', ')} in ${formatElapsed(elapsed)}`, 'success');
    } else if (result.length > 0 && result[0].columns && result[0].values.length > 0) {
      const totalRows = result[0].values.length;
      const truncated = totalRows > MAX_RESULT_ROWS;
      if (truncated) {
        result[0].values = result[0].values.slice(0, MAX_RESULT_ROWS);
      }
      const rows = sqlResultToRows(result);
      await showQueryResult(sql, rows);
      const suffix = truncated ? ` (showing first ${MAX_RESULT_ROWS.toLocaleString()} of ${totalRows.toLocaleString()})` : '';
      setStatus(`Query returned ${totalRows.toLocaleString()} row(s) in ${formatElapsed(elapsed)}${suffix}`, 'success');
    } else {
      const msg = result.length > 0 ? `${result[0].values.length} row(s) affected` : 'OK';
      setStatus(`${msg} in ${formatElapsed(elapsed)}`, 'success');
      refreshAllTableWindows();
    }
  }

  async function importNewDBTables(tableNames) {
    for (const name of tableNames) {
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
        await registerTable(name);
        createTableWindow(name);
      } catch (e) {}
    }
  }

  async function showQueryResult(sql, resultRows) {
    const rawColumns = Object.keys(resultRows[0]);
    const columns = sanitizeColumns(rawColumns);
    const tableName = '_query_' + nextWinId;

    // Store as a table so it can be saved
    const rows = resultRows.map((r, i) => {
      const row = { _rownum: i + 1 };
      rawColumns.forEach((raw, j) => { row[columns[j]] = r[raw] != null ? String(r[raw]) : ''; });
      return row;
    });
    tables[tableName] = { columns, rows, filename: null, modified: true };
    await registerTable(tableName);

    createSubwindow(tableName + ' *', (win, body) => {
      win.tableName = tableName;
      win.isQuery = true;
      renderTableView(win, body, tables[tableName]);
    }, { tableName, isQuery: true });
  }

  function refreshAllTableWindows() {
    const dbTables = new Set(getDBTables());
    // Collect windows to close (dropped tables) — iterate in reverse to avoid index shift
    const toClose = [];
    windows.forEach(w => {
      if (!w.tableName || !tables[w.tableName] || w.isQuery) return;
      if (!dbTables.has(w.tableName)) {
        toClose.push(w.id);
        return;
      }
      // Re-sync from SQL database
      try {
        const t = tables[w.tableName];
        const result = db.exec(`SELECT * FROM [${w.tableName}]`);
        if (result.length > 0) {
          const columns = sanitizeColumns(result[0].columns);
          const rows = sqlResultToRows(result);
          t.columns = columns;
          t.rows = rows.map((r, i) => {
            const row = { _rownum: i + 1 };
            columns.forEach(c => { row[c] = r[c] != null ? String(r[c]) : ''; });
            return row;
          });
        } else {
          t.columns = [];
          t.rows = [];
        }
      } catch (e) {}
      rebuildTable(w);
    });
    // Close windows for dropped tables (skip unsaved-changes prompt since SQL already dropped them)
    for (const id of toClose) {
      const win = windows.find(w => w.id === id);
      if (win) {
        delete tables[win.tableName];
        win.el.remove();
        windows.splice(windows.indexOf(win), 1);
        if (activeWinId === id) {
          activeWinId = windows.length ? windows[windows.length - 1].id : null;
          if (activeWinId) focusWindow(activeWinId);
        }
      }
    }
    if (toClose.length) updateWindowsList();
  }

  function clearConsole() {
    if (_activeConsoleTab === 'ai') {
      document.getElementById('ai-response').innerHTML = '';
      document.getElementById('ai-input').value = '';
      _aiConversation = [];
      _aiImages = {};
      renderImageThumbs();
    } else {
      document.getElementById('sql-input').value = '';
      document.getElementById('sql-highlight').innerHTML = '';
    }
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
      scaleWindowsToArea();
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

    // Wire up menu buttons that require window state
    document.getElementById('btn-save').addEventListener('click', () => saveActiveTable());
    document.getElementById('btn-save-as').addEventListener('click', () => saveActiveTableAs());
    document.getElementById('btn-close-window').addEventListener('click', () => closeActiveWindow());
    document.getElementById('btn-tile-h').addEventListener('click', () => layoutTileH());
    document.getElementById('btn-tile-v').addEventListener('click', () => layoutTileV());
    document.getElementById('btn-grid').addEventListener('click', () => layoutGrid());
    document.getElementById('btn-cascade').addEventListener('click', () => layoutCascade());
    document.getElementById('btn-minimize-all').addEventListener('click', () => minimizeAll());
    document.getElementById('btn-restore-all').addEventListener('click', () => restoreAll());
    document.getElementById('btn-load-plugin').addEventListener('click', () => loadPluginFromFile());
    document.getElementById('btn-expr-ref').addEventListener('click', () => showExpressionReference());
    document.getElementById('btn-undo').addEventListener('click', () => {
      const win = getActiveDataWindow();
      if (win && win.tableName) undoTable(win.tableName, win);
    });
    document.getElementById('btn-redo').addEventListener('click', () => {
      const win = getActiveDataWindow();
      if (win && win.tableName) redoTable(win.tableName, win);
    });
    document.getElementById('btn-cut').addEventListener('click', () => {
      const win = getActiveDataWindow();
      if (win && win.tableName) cutSelectedCells(win);
    });
    document.getElementById('btn-copy').addEventListener('click', () => {
      const win = getActiveDataWindow();
      if (win && win.tableName) copySelectedCells(win);
    });
    document.getElementById('btn-paste').addEventListener('click', () => {
      const win = getActiveDataWindow();
      if (win && win.tableName) pasteAtAnchor(win);
    });
    document.getElementById('btn-select-all').addEventListener('click', () => {
      const win = getActiveDataWindow();
      if (win && win.tableName) selectAllCells(win);
    });

    function updateMenuState() {
      const hasActive = !!activeWinId;
      const hasAny = windows.length > 0;
      document.getElementById('btn-save').disabled = !hasActive;
      document.getElementById('btn-save-as').disabled = !hasActive;
      document.getElementById('btn-close-window').disabled = !hasActive;
      document.getElementById('btn-tile-h').disabled = !hasAny;
      document.getElementById('btn-tile-v').disabled = !hasAny;
      document.getElementById('btn-grid').disabled = !hasAny;
      document.getElementById('btn-cascade').disabled = !hasAny;
      document.getElementById('btn-minimize-all').disabled = !hasAny;
      document.getElementById('btn-restore-all').disabled = !hasAny;
      const win = getActiveDataWindow();
      const t = win && win.tableName && tables[win.tableName];
      const undoEntry = t && t._undoStack && t._undoStack.length > 0 ? t._undoStack[t._undoStack.length - 1] : null;
      const redoEntry = t && t._redoStack && t._redoStack.length > 0 ? t._redoStack[t._redoStack.length - 1] : null;
      const actionLabel = (e) => e ? e.type.charAt(0).toUpperCase() + e.type.slice(1) : '';
      const btnUndo = document.getElementById('btn-undo');
      const btnRedo = document.getElementById('btn-redo');
      btnUndo.disabled = !undoEntry;
      btnRedo.disabled = !redoEntry;
      btnUndo.firstChild.textContent = undoEntry ? `Undo ${actionLabel(undoEntry)}` : 'Undo';
      btnRedo.firstChild.textContent = redoEntry ? `Redo ${actionLabel(redoEntry)}` : 'Redo';
      const hasSel = !!(win && win.selectedCells && win.selectedCells.size > 0);
      document.getElementById('btn-cut').disabled = !hasSel;
      document.getElementById('btn-copy').disabled = !hasSel;
      document.getElementById('btn-paste').disabled = !(win && win.anchorCell);
      document.getElementById('btn-select-all').disabled = !(win && win.tableName && tables[win.tableName]);
    }

    function openItem(item) {
      allItems.forEach(m => m.classList.remove('open'));
      item.classList.add('open');
      updateMenuState();
    }

    allItems.forEach(item => {
      item.querySelector('.menu-label').addEventListener('mousedown', (e) => {
        if (item.classList.contains('open')) {
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
      if (_menuDragging && btn && !btn.disabled) {
        btn.click();
        closeMenus();
      }
      _menuDragging = false;
    });

    // Clicking a dropdown button closes the menu bar (non-drag case)
    menubar.addEventListener('click', (e) => {
      if (e.target.closest('.menu-dropdown button')) closeMenus();
    });

    document.addEventListener('mouseup', () => {
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

  function showToast(message, type) {
    type = type || 'success';
    const toast = document.createElement('div');
    toast.className = 'toast toast-' + type;
    toast.textContent = message;
    document.body.appendChild(toast);
    let offset = 12;
    document.querySelectorAll('.toast').forEach(t => {
      if (t !== toast) offset += t.offsetHeight + 8;
    });
    toast.style.bottom = offset + 'px';
    setTimeout(() => {
      toast.classList.add('toast-fade-out');
      toast.addEventListener('animationend', () => toast.remove());
    }, 3000);
  }

  function closeActiveWindow() {
    if (activeWinId) closeWindow(activeWinId);
  }

  // ---- Help ----
  function showHelpWindow(title, bodyHTML) {
    const existing = windows.find(w => w.title === title);
    if (existing) {
      if (existing.el.classList.contains('minimized')) restoreWindow(existing.id);
      else focusWindow(existing.id);
      return;
    }
    const area = document.getElementById('window-area');
    const rect = area.getBoundingClientRect();
    const w = Math.min(600, rect.width - 60);
    const h = Math.min(500, rect.height - 40);
    createSubwindow(title, (win, body) => {
      body.innerHTML = `<div class="help-body">${bodyHTML}</div>`;
    }, { width: w, height: h });
  }

  function showAbout() {
    const license = `MIT License

Copyright (c) 2026 Mark Kim

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.`;
    showHelpWindow('About CSVSQL', `
      <p><strong>CSVSQL</strong> &mdash; A browser-based CSV database with SQL query support.</p>
      <p>Version 0.20.0 &mdash; &copy; 2026 Mark Kim</p>
      <h4>License</h4>
      <div class="about-text">${escHtml(license)}</div>
    `);
  }

  function showManual() {
    showHelpWindow("User's Manual", `
<h4>Overview</h4>
<p>CSVSQL treats CSV and other data files as database tables. Open files, edit cells, run SQL queries, and save &mdash; all in the browser with no server required.</p>
<p>Install from PyPI with <code>pip install csvsql</code>, then run <code>csvsql</code> to start. If <code>csvsql</code> conflicts with another command on your system, use <code>csvsqlw</code> instead &mdash; it&rsquo;s an identical alias.</p>

<h4>Opening Files</h4>
<p>Use <strong>File &rarr; Open</strong> (<code>Ctrl+O</code> / <code>&#8984;O</code>), <strong>File &rarr; Open URL</strong>, or drag and drop files onto the window. Hold <strong>Shift</strong> while opening or dropping to load files without headers &mdash; columns will be named A, B, C, &hellip; Z, AA, AB, etc.</p>

<table>
<tr><th>Format</th><th>Extensions</th><th>Notes</th></tr>
<tr><td>CSV</td><td>.csv, .txt</td><td>Delimiter auto-detected (comma, tab, pipe, etc.)</td></tr>
<tr><td>TSV</td><td>.tsv</td><td>Tab-delimited</td></tr>
<tr><td>PSV</td><td>.psv</td><td>Pipe-delimited</td></tr>
<tr><td>Excel</td><td>.xlsx, .xls</td><td>Each non-empty worksheet opens as a separate table</td></tr>
<tr><td>Gzip</td><td>.csv.gz, etc.</td><td>Decompressed in browser; inner file opened by type</td></tr>
<tr><td>ZIP</td><td>.zip</td><td>All recognized data files inside the archive are opened</td></tr>
</table>
<p>Compressed formats (.bz2, .xz, .rar, .7z, .zst) are recognized but not yet supported for in-browser decompression &mdash; decompress these externally first.</p>

<h4>Saving Files</h4>
<p><strong>Save</strong> (<code>Ctrl+S</code> / <code>&#8984;S</code>) writes directly back to the original file if the browser supports the File System Access API (Chrome, Edge). On browsers without this API (Firefox), Save triggers a download.</p>
<p><strong>Save As</strong> always prompts for a new filename and location. You can choose CSV, TSV, PSV, Excel (.xlsx), Gzip, or ZIP format.</p>
<p><strong>When Save acts as Save As:</strong> Save falls back to Save As when the table has no associated file handle &mdash; for example, tables created via <code>New Table</code>, SQL query results, or tables created with <code>SELECT INTO</code>.</p>

<p><strong>Multi-file save behavior:</strong></p>
<ul>
<li><strong>ZIP archives:</strong> Saving any table from a ZIP re-packs all tables from that archive into the same ZIP. If any table from the archive has been closed, Save falls back to Save As to avoid data loss.</li>
<li><strong>Excel workbooks:</strong> Saving any sheet from an Excel file re-packs all sheets into the same workbook. If any sheet has been closed, Save falls back to Save As.</li>
<li><strong>Gzip files:</strong> Saved back to the original .gz file with compression.</li>
</ul>

<h4>Editing</h4>
<ul>
<li><strong>Edit cells:</strong> Click a cell to select it; press <code>Enter</code>, <code>i</code>, <code>F2</code>, or <code>Ctrl</code>/<code>&#8984;</code>+<code>U</code> to enter edit mode, or <code>Ctrl</code>/<code>&#8984;</code>+click a cell to edit it directly. <code>Tab</code>/<code>Shift+Tab</code> moves between cells, <code>Enter</code> saves and moves down, <code>Escape</code> reverts the edit (and clears the selection when not editing).</li>
<li><strong>Highlight row &amp; column:</strong> Clicking a cell also highlights its row and column. Move the selection with arrow keys or vim-style <code>h</code>/<code>j</code>/<code>k</code>/<code>l</code>; extend to a rectangle of cells with <code>Shift</code>+arrow (or <code>Shift</code>+<code>H</code>/<code>J</code>/<code>K</code>/<code>L</code>), <code>Shift</code>+click on another cell, or click-and-drag across cells &mdash; every selected cell's row and column is highlighted so you can see what lines up with what. Click a row number to select an entire row; drag across row numbers or <code>Shift</code>+click another row number to select a range. Click the <code>#</code> corner cell or press <code>Ctrl</code>/<code>&#8984;</code>+<code>Shift</code>+<code>A</code> to select all. Pressing an arrow key with no cell selected focuses the cell in the middle of the current view. <code>Esc</code> clears the selection.</li>
<li><strong>Cut / Copy / Paste:</strong> Select cells and use <code>Ctrl</code>/<code>&#8984;</code>+<code>X</code>, <code>Ctrl</code>/<code>&#8984;</code>+<code>C</code>, <code>Ctrl</code>/<code>&#8984;</code>+<code>V</code>. Data is copied as tab-separated values. Select All and row selection copies include the column header row. In edit mode, these shortcuts pass through to native browser behavior for text within the cell.</li>
<li><strong>Undo / Redo:</strong> <code>Ctrl</code>/<code>&#8984;</code>+<code>Z</code> to undo, <code>Ctrl</code>/<code>&#8984;</code>+<code>Shift</code>+<code>Z</code> to redo. Undoes cell edits, paste, and cut operations. Multi-cell paste and cut undo as a single step. Also available from the Edit menu.</li>
<li><strong>Add rows:</strong> Click <code>+ Row</code> in the toolbar, or right-click a row number to insert above.</li>
<li><strong>Delete rows:</strong> Right-click a row number and choose Delete Row.</li>
<li><strong>Add columns:</strong> Click <code>+ Col</code> in the toolbar.</li>
<li><strong>Rename columns:</strong> <code>Ctrl</code>/<code>&#8984;</code>+click a column header.</li>
<li><strong>Select a column:</strong> Click a column header to select it (highlighted) and sort it. Selection is the target for Ctrl+&larr;/&rarr;.</li>
<li><strong>Reorder columns:</strong> Drag a column header to a new position. With a column selected by clicking its header, press <code>Ctrl</code>/<code>&#8984;</code>+<code>&larr;</code>/<code>&rarr;</code> to nudge it. With cells selected (in select mode, not editing), <code>Ctrl</code>/<code>&#8984;</code>+<code>&larr;</code>/<code>&rarr;</code> moves the columns spanned by the selection.</li>
<li><strong>Resize columns:</strong> Drag the right edge of a column header to resize. Double-click the edge to auto-fit the column to its content. Column widths are fixed after initial load and survive sorting and filtering.</li>
<li><strong>Rename tables:</strong> <code>Ctrl</code>/<code>&#8984;</code>+click the window title.</li>
</ul>

<h4>Touch Gestures</h4>
<ul>
<li><strong>Edit a cell:</strong> Double-tap the cell to enter edit mode.</li>
<li><strong>Select a rectangle of cells:</strong> Tap a cell, then tap and pan from any cell to draw a selection rectangle from the first cell to the panned cell.</li>
<li><strong>Reorder a column:</strong> Tap a column header, then tap-and-hold the same header and pan to the target position.</li>
<li><strong>Move a window:</strong> Tap the window's title bar, then tap the title bar again and pan to move the window.</li>
</ul>

<h4>Sorting &amp; Filtering</h4>
<ul>
<li><strong>Sort:</strong> Click a column header to sort ascending &rarr; descending &rarr; unsorted.</li>
<li><strong>Multi-column sort:</strong> Shift+click additional column headers. Numbers next to arrows indicate sort priority.</li>
<li><strong>Filter:</strong> Type a SQL <code>WHERE</code> clause in the filter bar (without the <code>WHERE</code> keyword). For example: <code>age > 30 AND name LIKE '%Smith%'</code></li>
<li>The filter supports all SQLite expressions including <code>REGEXP</code> (see below).</li>
<li><strong>Column autofilter:</strong> Click the <code>&#x2630;</code> icon on any column header to open a dropdown with checkboxes for each unique value. Use the search box to narrow the list. Uncheck values and click Apply to hide matching rows. Multiple column filters AND together and combine with the WHERE filter bar. Filtered columns show a green border. Click Clear to remove a column&rsquo;s filter. When any filters are active, a &ldquo;Clear Filters&rdquo; link appears in the status bar to reset all column autofilters and the WHERE filter at once.</li>
</ul>

<h4>SQL Console</h4>
<p>The SQL Console at the bottom runs queries against all open tables using SQLite syntax. Press <code>Ctrl+Enter</code> / <code>&#8984;+Enter</code> to execute. The console and filter inputs feature SQL syntax highlighting for keywords, strings, numbers, comments, and bracket-quoted identifiers.</p>
<p>Tables are referenced by the name shown in their window title bar. Names are sanitized to <code>[a-zA-Z0-9_]</code> characters.</p>
<p>Query results open in new windows and are automatically registered as queryable tables &mdash; you can run further SQL queries or use the filter bar on any result set.</p>

<h4>SQL Syntax Reference</h4>
<p>Standard SQLite syntax is supported. All column values are stored as TEXT.</p>

<pre>SELECT column1, column2 FROM tablename
  WHERE condition
  ORDER BY column1 ASC
  LIMIT 100</pre>

<p><strong>Joins, subqueries, aggregates, GROUP BY, HAVING, UNION, CASE</strong> &mdash; all standard SQLite features work.</p>

<h4>REGEXP</h4>
<p>CSVSQL adds a <code>REGEXP</code> function (not available in standard SQLite). It performs a case-insensitive regular expression match.</p>
<pre>SELECT * FROM employees
  WHERE name REGEXP '^(John|Jane)'

-- In the filter bar:
name REGEXP 'smith|jones'</pre>

<h4>SELECT INTO</h4>
<p>Use <code>SELECT ... INTO tablename ...</code> to create a new table from query results. The <code>INTO</code> clause can appear anywhere in the SELECT statement.</p>
<pre>SELECT name, salary INTO high_earners
  FROM employees WHERE salary > 100000

SELECT * FROM orders
  INTO us_orders
  WHERE country = 'US'</pre>
<p>The new table opens in its own window and can be edited, queried, and saved like any other table.</p>

<h4>CREATE TABLE</h4>
<p>New tables created via SQL automatically open as editable windows:</p>
<pre>CREATE TABLE projects (id, name, status)

INSERT INTO projects VALUES ('1', 'Alpha', 'active')</pre>

<h4>DDL &amp; DML</h4>
<p><code>INSERT</code>, <code>UPDATE</code>, <code>DELETE</code>, <code>ALTER TABLE</code>, and <code>DROP TABLE</code> all work. Changes to existing tables are reflected in their windows immediately after execution.</p>

<h4>Window Management</h4>
<ul>
<li><strong>Move:</strong> Drag the title bar. Windows are constrained to the workspace area.</li>
<li><strong>Resize:</strong> Drag any edge or corner. Windows cannot be resized beyond the workspace boundaries.</li>
<li><strong>Maximize/Restore:</strong> Double-click the title bar, or click the maximize button.</li>
<li><strong>Minimize:</strong> Click the minimize button. Restore from the Windows menu.</li>
<li><strong>Close:</strong> Click the close button. <code>Ctrl</code>/<code>&#8984;</code>+click closes all windows.</li>
<li><strong>Layout:</strong> Use the Windows menu to tile, grid, or cascade all windows.</li>
<li><strong>Proportional scaling:</strong> Windows reposition and resize proportionally when the browser window or console panel is resized.</li>
</ul>

<h4>Keyboard Shortcuts</h4>
<table>
<tr><th>Shortcut</th><th>Action</th></tr>
<tr><td><code>Ctrl+O</code> / <code>&#8984;O</code></td><td>Open file</td></tr>
<tr><td><code>Ctrl+S</code> / <code>&#8984;S</code></td><td>Save table</td></tr>
<tr><td><code>Ctrl+N</code> / <code>&#8984;N</code></td><td>New table</td></tr>
<tr><td><code>Ctrl+W</code> / <code>&#8984;W</code></td><td>Close window</td></tr>
<tr><td><code>Ctrl+&larr;</code> / <code>Ctrl+&rarr;</code> (or <code>&#8984;</code>+arrow)</td><td>Move selected header column, or cell-selection's columns, left / right</td></tr>
<tr><td><code>Enter</code>, <code>i</code>, <code>F2</code>, or <code>Ctrl</code>/<code>&#8984;</code>+<code>U</code></td><td>Enter edit mode on the selected cell</td></tr>
<tr><td><code>/</code> (cell selected, not editing)</td><td>Jump to the window's filter input</td></tr>
<tr><td><code>Escape</code> (in filter input)</td><td>Return focus to the selected cell</td></tr>
<tr><td>Arrow keys (no cell selected)</td><td>Focus the cell in the middle of the visible table</td></tr>
<tr><td>Arrow keys or <code>h</code>/<code>j</code>/<code>k</code>/<code>l</code> (cell selected, not editing)</td><td>Move selection to the adjacent cell</td></tr>
<tr><td><code>Shift+</code>arrow or <code>Shift+H</code>/<code>J</code>/<code>K</code>/<code>L</code></td><td>Extend cell selection (highlights row &amp; column of every selected cell)</td></tr>
<tr><td><code>Ctrl</code>/<code>&#8984;</code>+<code>Shift</code>+<code>A</code></td><td>Select all cells</td></tr>
<tr><td><code>Ctrl</code>/<code>&#8984;</code>+<code>X</code></td><td>Cut selected cells</td></tr>
<tr><td><code>Ctrl</code>/<code>&#8984;</code>+<code>C</code></td><td>Copy selected cells</td></tr>
<tr><td><code>Ctrl</code>/<code>&#8984;</code>+<code>V</code></td><td>Paste at selected cell</td></tr>
<tr><td><code>Ctrl</code>/<code>&#8984;</code>+<code>Z</code></td><td>Undo</td></tr>
<tr><td><code>Ctrl</code>/<code>&#8984;</code>+<code>Shift</code>+<code>Z</code></td><td>Redo</td></tr>
<tr><td><code>Tab</code>/<code>Shift+Tab</code> or <code>Ctrl</code>/<code>&#8984;</code>+<code>Shift</code>+<code>L</code>/<code>H</code> (cell selected, not editing)</td><td>Switch to next / previous table window</td></tr>
<tr><td><code>Ctrl</code>/<code>&#8984;</code>+<code>H</code>/<code>J</code>/<code>K</code>/<code>L</code> (cell selected, not editing)</td><td>Nudge the active window 5 px left / down / up / right</td></tr>
<tr><td><code>Ctrl+Enter</code></td><td>Execute SQL query</td></tr>
<tr><td><code>Enter</code></td><td>Send AI prompt</td></tr>
<tr><td><code>Shift+Enter</code></td><td>Newline in AI prompt</td></tr>
<tr><td><code>Up</code> / <code>Down</code></td><td>AI prompt history</td></tr>
</table>

<h4>AI Analysis <em>(experimental)</em></h4>
<p>The AI tab in the console panel lets you analyze your data using natural language. The AI has full SQL access to your data &mdash; it writes and executes queries automatically to answer your questions with exact results, regardless of dataset size. You can also chat with the AI without any tables loaded.</p>
<p><strong>Four provider options:</strong></p>
<ul>
<li><strong>WebLLM (default):</strong> Runs entirely in the browser via WebGPU. Requires Chrome/Edge 113+. No install, no API key, no data leaves your machine.</li>
<li><strong>Ollama:</strong> Local AI server. Install from <a href="https://ollama.com">ollama.com</a>, then run <code>ollama pull llama3.2</code>. Larger models than WebLLM, still fully local.</li>
<li><strong>Claude (Anthropic):</strong> Cloud provider. Requires an API key from <a href="https://console.anthropic.com">console.anthropic.com</a>. Best reasoning quality.</li>
<li><strong>OpenAI:</strong> Cloud provider. Requires an API key from <a href="https://platform.openai.com">platform.openai.com</a>.</li>
</ul>
<p><strong>Usage:</strong> Switch to the AI tab, select one or more tables, type your question, and press <code>Enter</code> or click Run. Use <code>Shift+Enter</code> for multiline prompts. Press <code>Up</code>/<code>Down</code> arrow to recall previous prompts.</p>
<p><strong>How it works:</strong> The AI receives column statistics and sample rows for context, then writes SQL queries in <code>\`\`\`sql</code> code blocks. These queries are executed automatically against the full dataset, and the results are fed back to the AI for analysis. This loop repeats (up to 5 rounds) until the AI has enough data to answer.</p>

<h4>Rich AI Output</h4>
<p>The AI can produce rich output beyond plain text:</p>
<ul>
<li><strong>Charts:</strong> Ask for a chart or visualization and the AI will render an interactive Chart.js chart inline. Example: &ldquo;show me a bar chart of sales by region&rdquo;</li>
<li><strong>Formatted tables:</strong> Ask for a formatted table and the AI will render a styled HTML table. Example: &ldquo;show the top 10 rows as a table&rdquo;</li>
<li><strong>PDF reports:</strong> Ask for a PDF or report and the AI will generate a downloadable PDF file. PDFs can include text, tables, embedded charts, and images. Example: &ldquo;generate a PDF report with charts and summary statistics&rdquo;</li>
</ul>
<p><strong>Images:</strong> Drag and drop PNG or JPG images onto the AI chat area to upload them. Uploaded images appear as thumbnails above the input field and can be included in PDF reports (e.g., as a logo). Click &times; to remove an image.</p>
<p>The AI queries your data with SQL first, then uses the results to build the rich output. Chart.js and jsPDF libraries are loaded on demand when first needed.</p>
<p>Click the gear icon &#9881; to configure the provider, model, and API keys.</p>

<h4>Plugins</h4>
<p>Plugins customize how cell values are displayed. A plugin is a JSON config file that maps table and column name patterns to display expressions written in the CSVSQL expression language.</p>

<p><strong>Loading &amp; unloading:</strong> Use <strong>Plugins &rarr; Load Plugin</strong> to open a <code>.json</code> plugin file, or drag and drop a <code>.json</code> file onto the app. A toast notification confirms success or shows errors. Loaded plugins appear in the Plugins menu with an &times; button to unload them. Click a plugin&rsquo;s name to open an About dialog showing its metadata, column rules, and an Unload button. Plugins persist across page reloads.</p>

<p><strong>How matching works:</strong> Each plugin has a <code>table</code> regex matched against table names, and column rules with a <code>match</code> regex matched against column names. Multiple plugins can be loaded simultaneously and stack on the same table &mdash; each column is governed by the <em>last-loaded</em> plugin with a matching rule. Unloading a plugin reveals any earlier plugin&rsquo;s rule that was shadowed.</p>

<p><strong>Plugin config format:</strong></p>
<pre>{
  "name": "Plugin Name",
  "version": "1.0.0",
  "author": "Author Name",
  "created": "2026-01-01",
  "description": "Optional description",
  "table": "regex for table names",
  "columns": [
    {
      "match": "regex for column names",
      "display": "expression"
    }
  ]
}</pre>
<p>The <code>version</code>, <code>author</code>, <code>created</code>, and <code>description</code> fields are optional metadata shown in the About dialog.</p>

<p><strong>Per-column toggle:</strong> Columns with an active plugin transform show a &#x1F50C; icon in the header. Click the icon to disable the transform for that column &mdash; the icon dims but stays visible. Click again to re-enable. The status bar shows a bulk <strong>Plugins on/off/partial</strong> toggle to enable or disable all transforms at once.</p>

<p><strong>Editing:</strong> When you enter edit mode on a cell with a display transform, the raw value is shown so you edit the actual data. The formatted display returns when you finish editing.</p>

<p><strong>Autofilter dropdowns:</strong> When a plugin transform is active for a column, the autofilter dropdown shows the formatted display values (not raw values). Searching within the dropdown also matches against the formatted text.</p>

<p><strong>Saving, SQL, filtering, sorting:</strong> All operate on raw values, not the plugin&rsquo;s display output.</p>

<p><strong>Expression language:</strong> See <strong>Help &rarr; Plugin Expression Reference</strong> for the full language reference, or the <a href="https://github.com/markuskimius/csvsql" target="_blank">README</a> for a quick overview. Common examples:</p>
<pre>date(value, 'locale')                           Locale-format a date
'$' + fixed(num(value), 2)                      Currency
choose(value, 'A', 'Active', 'I', 'Inactive')   Value mapping
value || 'N/A'                                   Default for empty</pre>

<p>Bundled example plugins are available in the <code>plugins/</code> directory.</p>

<h4>Links</h4>
<p><a href="https://github.com/markuskimius/csvsql" target="_blank">GitHub</a></p>
    `);
  }

  // ---- AI Analysis ----
  let _aiProvider = null;
  let _aiAbort = null;
  let _webllmEngine = null;
  let _aiConversation = []; // accumulated user/assistant message history
  let _aiImages = {};       // name -> data URL for images dropped into AI chat

  let aiSettings = JSON.parse(localStorage.getItem('csvsql_ai_settings') || 'null') || {
    provider: 'webllm',
    model: '',
    ollamaUrl: 'http://localhost:11434',
    claudeApiKey: '',
    openaiApiKey: '',
  };
  // Ensure keys exist for older saved settings
  if (!('claudeApiKey' in aiSettings)) aiSettings.claudeApiKey = '';
  if (!('openaiApiKey' in aiSettings)) aiSettings.openaiApiKey = '';

  function saveAISettings() {
    localStorage.setItem('csvsql_ai_settings', JSON.stringify(aiSettings));
  }

  function setAIStatus(msg, type = '') {
    setStatus(msg, type);
  }

  // Tab switching
  function setupConsoleTabs() {
    const tabs = document.querySelectorAll('.console-tab');
    tabs.forEach(tab => {
      tab.addEventListener('click', () => {
        switchConsoleTab(tab.dataset.tab);
      });
    });
  }

  function switchConsoleTab(tab) {
    _activeConsoleTab = tab;
    document.querySelectorAll('.console-tab').forEach(t => {
      t.classList.toggle('active', t.dataset.tab === tab);
    });
    document.getElementById('console-body').style.display = tab === 'sql' ? '' : 'none';
    document.getElementById('ai-body').style.display = tab === 'ai' ? '' : 'none';
    if (tab === 'ai') {
      populateTableSelect();
      detectAIProvider();
      document.getElementById('ai-input').focus();
    } else {
      document.getElementById('sql-input').focus();
    }
  }

  function runConsole() {
    if (_activeConsoleTab === 'ai') runAI();
    else executeQuery();
  }

  function populateTableSelect() {
    const sel = document.getElementById('ai-table-select');
    if (!sel) return;
    const prev = new Set([...sel.selectedOptions].map(o => o.value));
    const hadPrev = prev.size > 0;
    sel.innerHTML = '';
    for (const name of Object.keys(tables)) {
      const opt = document.createElement('option');
      opt.value = name;
      opt.textContent = `${name} (${tables[name].rows.length} rows, ${tables[name].columns.length} cols)`;
      // Select all by default; preserve previous selection if user had one
      opt.selected = hadPrev ? prev.has(name) : true;
      sel.appendChild(opt);
    }
  }

  // Provider detection
  async function detectAIProvider() {
    const badge = document.getElementById('ai-provider-badge');
    if (!badge) return;

    if (aiSettings.provider === 'ollama' || aiSettings.provider === 'auto') {
      try {
        const r = await fetch(aiSettings.ollamaUrl + '/api/tags', { signal: AbortSignal.timeout(2000) });
        if (r.ok) {
          const data = await r.json();
          const models = (data.models || []).map(m => m.name);
          _aiProvider = 'ollama';
          if (!aiSettings.model || !models.includes(aiSettings.model)) {
            aiSettings.model = models[0] || '';
            saveAISettings();
          }
          badge.textContent = 'Ollama: ' + (aiSettings.model || 'no models');
          return;
        }
      } catch {}
    }

    if (aiSettings.provider === 'claude') {
      if (aiSettings.claudeApiKey) {
        _aiProvider = 'claude';
        if (!aiSettings.model) {
          aiSettings.model = 'claude-opus-4-20250514';
          saveAISettings();
        }
        badge.textContent = 'Claude: ' + aiSettings.model;
        return;
      }
    }

    if (aiSettings.provider === 'openai') {
      if (aiSettings.openaiApiKey) {
        _aiProvider = 'openai';
        if (!aiSettings.model) {
          aiSettings.model = 'o3';
          saveAISettings();
        }
        badge.textContent = 'OpenAI: ' + aiSettings.model;
        return;
      }
    }

    if (aiSettings.provider === 'webllm' || aiSettings.provider === 'auto') {
      if (navigator.gpu) {
        _aiProvider = 'webllm';
        badge.textContent = 'WebLLM (WebGPU)';
        if (!aiSettings.model) {
          aiSettings.model = 'Qwen3-8B-q4f16_1-MLC';
          saveAISettings();
        }
        return;
      }
    }

    _aiProvider = null;
    badge.textContent = 'No AI provider';
    showSetupHelp();
  }

  function showSetupHelp() {
    const resp = document.getElementById('ai-response');
    if (!resp || resp.innerHTML.includes('ai-setup-help')) return;
    resp.innerHTML = `<div class="ai-setup-help">
<strong>No AI provider detected.</strong> Options:<br><br>
<strong>Option 1: Claude or OpenAI (cloud)</strong><br>
Click the gear icon, select Claude or OpenAI, and enter your API key.<br>
Best reasoning quality for data analysis.<br><br>
<strong>Option 2: Ollama (local)</strong><br>
1. Install from <a href="https://ollama.com" target="_blank">ollama.com</a><br>
2. Run: <code>ollama pull llama3.2</code><br>
3. Ollama runs on localhost:11434 by default<br><br>
<strong>Option 3: WebLLM (in-browser)</strong><br>
Requires Chrome 113+ with WebGPU enabled.<br>
Smaller models, runs entirely in the browser — no install needed.<br>
</div>`;
  }

  // Build data context for the AI prompt
  // Budget for data context chars. The AI also has SQL query access to the full
  // dataset, so this budget is for column stats + sample rows to orient the model.
  function getDataCharBudget() {
    if (_aiProvider === 'claude' || _aiProvider === 'openai') return 500000;
    if (_aiProvider === 'webllm') return 20000; // 16K token context; leave room for system prompt, conversation, and response
    return 100000; // ollama
  }

  function buildColumnStats(tableName, columns) {
    const stats = [];
    const maxDetailCols = 30;
    for (let ci = 0; ci < columns.length; ci++) {
      const col = columns[ci];
      const fullStats = ci < maxDetailCols;
      try {
        const basic = db.exec(`SELECT COUNT([${col}]), COUNT(CASE WHEN [${col}] IS NULL OR [${col}] = '' THEN 1 END), COUNT(DISTINCT [${col}]), MIN([${col}]), MAX([${col}]) FROM [${tableName}]`);
        if (!basic.length) continue;
        const [total, nullEmpty, distinct, mn, mx] = basic[0].values[0];
        let line = `  ${col}: ${total} values, ${nullEmpty} null/empty, ${distinct} distinct, min=${mn}, max=${mx}`;

        if (!fullStats) { stats.push(line); continue; }

        // Check if numeric
        const numCheck = db.exec(`SELECT COUNT(*) FROM [${tableName}] WHERE [${col}] GLOB '*[0-9]*' AND TYPEOF(CAST([${col}] AS REAL)) = 'real'`);
        const numCount = numCheck.length ? numCheck[0].values[0][0] : 0;
        const isNumeric = numCount > total * 0.5;

        if (isNumeric) {
          try {
            const agg = db.exec(`SELECT AVG(CAST([${col}] AS REAL)), AVG(CAST([${col}] AS REAL) * CAST([${col}] AS REAL)) FROM [${tableName}] WHERE [${col}] GLOB '*[0-9]*'`);
            if (agg.length) {
              const mean = agg[0].values[0][0];
              const meanSq = agg[0].values[0][1];
              const variance = meanSq - mean * mean;
              const stddev = variance > 0 ? Math.sqrt(variance) : 0;
              line += `\n    Numeric: mean=${Number(mean).toFixed(2)}, stddev=${Number(stddev).toFixed(2)}`;
            }
            // Percentiles via OFFSET
            const cntRes = db.exec(`SELECT COUNT(*) FROM [${tableName}] WHERE [${col}] GLOB '*[0-9]*'`);
            const cnt = cntRes.length ? cntRes[0].values[0][0] : 0;
            if (cnt > 4) {
              const pcts = [0.25, 0.5, 0.75];
              const pVals = [];
              for (const p of pcts) {
                const off = Math.floor(cnt * p);
                const pr = db.exec(`SELECT CAST([${col}] AS REAL) FROM [${tableName}] WHERE [${col}] GLOB '*[0-9]*' ORDER BY CAST([${col}] AS REAL) LIMIT 1 OFFSET ${off}`);
                pVals.push(pr.length ? Number(pr[0].values[0][0]).toFixed(2) : '?');
              }
              line += `\n    Percentiles: p25=${pVals[0]}, median=${pVals[1]}, p75=${pVals[2]}`;
            }
          } catch {}
        } else if (distinct <= 1000 && distinct > 0) {
          // Categorical: top 20 value counts
          try {
            const topRes = db.exec(`SELECT [${col}], COUNT(*) as cnt FROM [${tableName}] WHERE [${col}] IS NOT NULL AND [${col}] != '' GROUP BY [${col}] ORDER BY cnt DESC LIMIT 20`);
            if (topRes.length) {
              const vals = topRes[0].values.map(r => `${r[0]} (${r[1]})`).join(', ');
              line += `\n    Top values: ${vals}`;
              if (distinct > 20) line += ` ... and ${distinct - 20} more`;
            }
          } catch {}
        }

        stats.push(line);
      } catch {}
    }
    return stats.length ? 'Column statistics:\n' + stats.join('\n') + '\n' : '';
  }

  function buildDataContext(tableNames) {
    const parts = [];
    if (tableNames.length === 0) return '';
    const budgetPerTable = Math.floor(getDataCharBudget() / tableNames.length);

    for (const name of tableNames) {
      const t = tables[name];
      if (!t || t.columns.length === 0) continue;

      let info = `Table: ${name}\nColumns: ${t.columns.join(', ')}\nTotal rows: ${t.rows.length}\n\n`;

      // For small tables (<=50 rows), include all rows directly
      if (t.rows.length <= 50) {
        const header = t.columns.join(' | ');
        const maxCellLen = 100;
        const rowLines = t.rows.map(row =>
          t.columns.map(c => {
            const v = String(row[c] ?? '');
            return v.length > maxCellLen ? v.substring(0, maxCellLen) + '...' : v;
          }).join(' | ')
        );
        info += 'Data:\n' + header + '\n' + rowLines.join('\n') + '\n';
        parts.push(info);
        continue;
      }

      // For larger tables: compute column stats first, then fill with sample rows
      const statsText = buildColumnStats(name, t.columns);
      info += statsText + '\n';

      // Determine sample size
      const sampleSize = t.rows.length <= 10000 ? 50 : 100;

      // Get random sample rows via SQL
      try {
        const sampleRes = db.exec(`SELECT * FROM [${name}] ORDER BY RANDOM() LIMIT ${sampleSize}`);
        if (sampleRes.length) {
          const sampleCols = sampleRes[0].columns;
          const header = sampleCols.join(' | ');
          const maxCellLen = 100;
          let sampleText = `Sample data (${sampleRes[0].values.length} random rows):\n${header}\n`;
          for (const row of sampleRes[0].values) {
            const line = row.map(v => {
              const s = String(v ?? '');
              return s.length > maxCellLen ? s.substring(0, maxCellLen) + '...' : s;
            }).join(' | ');
            // Check budget
            if (info.length + sampleText.length + line.length + 1 > budgetPerTable) break;
            sampleText += line + '\n';
          }
          info += sampleText;
        }
      } catch {}

      parts.push(info);
    }
    return parts.join('\n---\n');
  }

  // Rich block system prompt instructions
  const _richBlockPrompt = `RICH OUTPUT — IMPORTANT:
When the user asks for a chart, visualization, formatted table, or PDF/report, you MUST output special code blocks. These blocks are rendered automatically — do NOT describe or explain the JSON, just output the block.

CHART — use the language tag "chart" (NOT "json"):
\`\`\`chart
{"type":"bar","data":{"labels":["A","B","C"],"datasets":[{"label":"Sales","data":[10,20,30]}]}}
\`\`\`
Supported chart types: bar, line, pie, doughnut, radar, polarArea, scatter, bubble. The JSON must be a valid Chart.js configuration object with "type" and "data" keys. Always include real data from your SQL query results. Do NOT use JavaScript functions or callbacks anywhere in the config — only plain JSON values (strings, numbers, booleans, arrays, objects, null).

TABLE — use the language tag "table" (NOT "json"):
\`\`\`table
{"columns":["Name","Value"],"rows":[{"Name":"Alice","Value":"100"},{"Name":"Bob","Value":"200"}]}
\`\`\`
The JSON must have "columns" (array of strings) and "rows" (array of objects keyed by column name). Use this for nicely formatted data tables.

PDF — use the language tag "pdf" (NOT "json"):
\`\`\`pdf
{"filename":"report.pdf","title":"My Report","content":[{"type":"heading","value":"Summary"},{"type":"text","value":"Total revenue was $1.2M."},{"type":"table","columns":["Product","Revenue"],"rows":[{"Product":"A","Revenue":"$500K"}]}]}
\`\`\`
Content block types: "heading", "text", "table", "chart", and "image". For chart blocks, include a "chart" key with a Chart.js config object. For image blocks, include a "name" key matching an uploaded image filename (e.g. {"type":"image","name":"logo.png"}). This generates a downloadable PDF file.

CRITICAL RULES:
- The code fence language MUST be chart, table, or pdf — never use json, javascript, or other languages for these blocks.
- Output ONLY valid JSON inside the block — no comments, no markdown, no extra text.
- Query the data with SQL first, then use the actual results to populate the rich block.
- You can combine rich blocks with regular text explanations.`;

  // Build image context for system prompt
  function _aiImageContext() {
    const names = Object.keys(_aiImages);
    if (names.length === 0) return '';
    return `\nIMAGES — The user has uploaded these images: ${names.join(', ')}
To include an image in a PDF, add a content block: {"type":"image","name":"filename.png"}
You can optionally set "width" (in PDF points, max is page width). Reference images by their exact filename.`;
  }

  // Lazy script loader
  const _loadedScripts = {};
  function loadScript(url) {
    if (_loadedScripts[url]) return _loadedScripts[url];
    _loadedScripts[url] = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = url;
      s.onload = resolve;
      s.onerror = () => reject(new Error('Failed to load ' + url));
      document.head.appendChild(s);
    });
    return _loadedScripts[url];
  }

  async function ensureChartJs() {
    if (typeof Chart === 'undefined') {
      await loadScript('lib/chart.umd.js');
    }
  }

  async function ensureJsPDF() {
    if (typeof jspdf === 'undefined') {
      await loadScript('lib/jspdf.umd.min.js');
      await loadScript('lib/jspdf.plugin.autotable.min.js');
    }
  }

  // Format AI response text with basic markdown
  function formatAIResponse(text) {
    let html = escHtml(text);
    // Code blocks: ```...``` — tag rich blocks with CSS classes for post-processing
    html = html.replace(/```(\w*)\n?([\s\S]*?)```/g, (_, lang, code) => {
      const cls = ['chart','table','pdf'].includes(lang) ? ` class="ai-block-${lang}"` : '';
      return `<pre${cls}>` + code.trim() + '</pre>';
    });
    // Inline code
    html = html.replace(/`([^`]+)`/g, '<code>$1</code>');
    // Bold
    html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
    // Newlines (but not inside pre tags)
    html = html.replace(/\n/g, '<br>');
    // Clean up <br> inside <pre>
    html = html.replace(/<pre([^>]*)>([\s\S]*?)<\/pre>/g, (_, attrs, code) => '<pre' + attrs + '>' + code.replace(/<br>/g, '\n') + '</pre>');
    return html;
  }

  // Apply dark theme defaults to Chart.js config
  function applyChartDefaults(config) {
    const colors = {
      text: '#e0e0f0',
      grid: '#444466',
      bg: '#2a2a3d',
    };
    if (!config.options) config.options = {};
    config.options.responsive = true;
    config.options.maintainAspectRatio = false;
    if (!config.options.plugins) config.options.plugins = {};
    if (!config.options.plugins.legend) config.options.plugins.legend = {};
    if (!config.options.plugins.legend.labels) config.options.plugins.legend.labels = {};
    config.options.plugins.legend.labels.color = colors.text;
    if (!config.options.plugins.title) config.options.plugins.title = {};
    config.options.plugins.title.color = colors.text;
    if (!config.options.scales) config.options.scales = {};
    for (const axis of ['x', 'y']) {
      if (!config.options.scales[axis]) config.options.scales[axis] = {};
      if (!config.options.scales[axis].ticks) config.options.scales[axis].ticks = {};
      config.options.scales[axis].ticks.color = colors.text;
      if (!config.options.scales[axis].grid) config.options.scales[axis].grid = {};
      config.options.scales[axis].grid.color = colors.grid;
    }
    // Default color palette if not specified
    if (config.data && config.data.datasets) {
      const palette = ['#7c6ff7','#55cc88','#e05577','#f0a050','#50b0f0','#cc77dd','#77ddaa','#ddaa55'];
      config.data.datasets.forEach((ds, i) => {
        if (!ds.backgroundColor) ds.backgroundColor = palette[i % palette.length];
        if (!ds.borderColor) ds.borderColor = palette[i % palette.length];
      });
    }
    return config;
  }

  // Detect rich block type from JSON structure
  function detectRichBlockType(obj) {
    if (obj && obj.type && obj.data && obj.data.datasets) return 'chart';
    if (obj && obj.columns && obj.rows && Array.isArray(obj.columns)) return 'table';
    if (obj && obj.content && Array.isArray(obj.content)) return 'pdf';
    return null;
  }

  // Extract JSON object from text that may have leading/trailing non-JSON content
  function extractJSON(text) {
    text = text.trim();
    const start = text.indexOf('{');
    if (start < 0) return null;
    // Find matching closing brace, respecting strings
    let depth = 0;
    let inStr = false;
    let esc = false;
    for (let i = start; i < text.length; i++) {
      const ch = text[i];
      if (esc) { esc = false; continue; }
      if (ch === '\\' && inStr) { esc = true; continue; }
      if (ch === '"') { inStr = !inStr; continue; }
      if (inStr) continue;
      if (ch === '{') depth++;
      else if (ch === '}') { depth--; if (depth === 0) return text.slice(start, i + 1); }
    }
    return null;
  }

  // Strip JS function expressions from a JSON-like string so JSON.parse can handle it
  function sanitizeChartJSON(text) {
    // Find each "function" keyword and replace through its matching closing brace with null
    let result = '';
    let i = 0;
    while (i < text.length) {
      const funcIdx = text.indexOf('function', i);
      if (funcIdx < 0) { result += text.slice(i); break; }
      result += text.slice(i, funcIdx);
      // Find opening brace of function body
      let j = text.indexOf('{', funcIdx);
      if (j < 0) { result += text.slice(funcIdx); break; }
      // Match braces to find end
      let depth = 0;
      let k = j;
      for (; k < text.length; k++) {
        if (text[k] === '{') depth++;
        else if (text[k] === '}') { depth--; if (depth === 0) break; }
      }
      result += 'null';
      i = k + 1;
    }
    return result;
  }

  // Show error inline when a rich block fails to render
  function showBlockError(pre, type, err) {
    console.warn(`AI ${type} block error:`, err);
    const msg = document.createElement('div');
    msg.style.cssText = 'color:var(--danger);font-size:11px;margin:4px 0;';
    msg.textContent = `Failed to render ${type}: ${err.message}`;
    pre.after(msg);
  }

  // Post-process AI rich blocks (chart, table, pdf) after streaming completes
  async function postProcessAIBlocks(bubbleEl) {
    // Fallback: check unclassed pre blocks for JSON that looks like rich content
    for (const pre of bubbleEl.querySelectorAll('pre:not([class])')) {
      try {
        const jsonStr = extractJSON(pre.textContent);
        if (!jsonStr) continue;
        const obj = JSON.parse(jsonStr);
        const type = detectRichBlockType(obj);
        if (type) pre.classList.add('ai-block-' + type);
      } catch {}
    }

    // Process table blocks
    for (const pre of [...bubbleEl.querySelectorAll('pre.ai-block-table')]) {
      try {
        const jsonStr = extractJSON(pre.textContent);
        if (!jsonStr) continue;
        const data = JSON.parse(jsonStr);
        if (!data.columns || !data.rows) continue;
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
      } catch (e) { showBlockError(pre, 'table', e); }
    }

    // Process chart blocks
    for (const pre of [...bubbleEl.querySelectorAll('pre.ai-block-chart')]) {
      try {
        const jsonStr = extractJSON(pre.textContent);
        if (!jsonStr) continue;
        const config = JSON.parse(sanitizeChartJSON(jsonStr));
        await ensureChartJs();
        applyChartDefaults(config);
        const container = document.createElement('div');
        container.className = 'ai-chart-container';
        const canvas = document.createElement('canvas');
        container.appendChild(canvas);
        pre.replaceWith(container);
        new Chart(canvas, config);
      } catch (e) { showBlockError(pre, 'chart', e); }
    }

    // Process PDF blocks
    for (const pre of [...bubbleEl.querySelectorAll('pre.ai-block-pdf')]) {
      try {
        const jsonStr = extractJSON(pre.textContent);
        if (!jsonStr) continue;
        const spec = JSON.parse(jsonStr);
        await ensureJsPDF();
        const doc = new jspdf.jsPDF();
        let y = 20;
        if (spec.title) {
          doc.setFontSize(18);
          doc.text(spec.title, 14, y);
          y += 12;
        }
        if (spec.content) {
          for (const block of spec.content) {
            if (block.type === 'text') {
              doc.setFontSize(block.fontSize || 12);
              const lines = doc.splitTextToSize(block.value || '', 180);
              doc.text(lines, 14, y);
              y += lines.length * (block.fontSize || 12) * 0.5 + 4;
            } else if (block.type === 'heading') {
              doc.setFontSize(block.fontSize || 14);
              doc.text(block.value || '', 14, y);
              y += 10;
            } else if (block.type === 'table' && block.columns && block.rows) {
              doc.autoTable({
                startY: y,
                head: [block.columns],
                body: block.rows.map(r => block.columns.map(c => String(r[c] ?? ''))),
                theme: 'grid',
                headStyles: { fillColor: [124, 111, 247] },
                styles: { fontSize: 9 },
              });
              y = doc.lastAutoTable.finalY + 8;
            } else if (block.type === 'chart' && block.chart) {
              await ensureChartJs();
              const chartConfig = JSON.parse(sanitizeChartJSON(
                typeof block.chart === 'string' ? block.chart : JSON.stringify(block.chart)
              ));
              // Render chart on offscreen canvas
              const offCanvas = document.createElement('canvas');
              offCanvas.width = block.width || 600;
              offCanvas.height = block.height || 300;
              // White background for PDF readability
              const ctx2d = offCanvas.getContext('2d');
              ctx2d.fillStyle = '#ffffff';
              ctx2d.fillRect(0, 0, offCanvas.width, offCanvas.height);
              // Override colors for light-background PDF
              if (!chartConfig.options) chartConfig.options = {};
              chartConfig.options.responsive = false;
              chartConfig.options.animation = false;
              const pdfColors = { text: '#333333', grid: '#cccccc' };
              if (!chartConfig.options.scales) chartConfig.options.scales = {};
              for (const axis of ['x', 'y']) {
                if (!chartConfig.options.scales[axis]) chartConfig.options.scales[axis] = {};
                if (!chartConfig.options.scales[axis].ticks) chartConfig.options.scales[axis].ticks = {};
                chartConfig.options.scales[axis].ticks.color = pdfColors.text;
                if (!chartConfig.options.scales[axis].grid) chartConfig.options.scales[axis].grid = {};
                chartConfig.options.scales[axis].grid.color = pdfColors.grid;
              }
              if (!chartConfig.options.plugins) chartConfig.options.plugins = {};
              if (!chartConfig.options.plugins.legend) chartConfig.options.plugins.legend = {};
              if (!chartConfig.options.plugins.legend.labels) chartConfig.options.plugins.legend.labels = {};
              chartConfig.options.plugins.legend.labels.color = pdfColors.text;
              const chart = new Chart(offCanvas, chartConfig);
              const imgData = offCanvas.toDataURL('image/png');
              chart.destroy();
              // Fit chart image to page width
              const pageW = doc.internal.pageSize.getWidth() - 28;
              const aspect = offCanvas.height / offCanvas.width;
              const imgH = pageW * aspect;
              if (y + imgH > doc.internal.pageSize.getHeight() - 20) {
                doc.addPage();
                y = 20;
              }
              doc.addImage(imgData, 'PNG', 14, y, pageW, imgH);
              y += imgH + 8;
            } else if (block.type === 'image' && block.name) {
              const dataUrl = _aiImages[block.name];
              if (dataUrl) {
                // Load image to get dimensions
                const img = await new Promise((res, rej) => {
                  const im = new Image();
                  im.onload = () => res(im);
                  im.onerror = () => rej(new Error('Failed to load image: ' + block.name));
                  im.src = dataUrl;
                });
                const format = dataUrl.startsWith('data:image/png') ? 'PNG' : 'JPEG';
                const pageW = doc.internal.pageSize.getWidth() - 28;
                const maxW = block.width || pageW;
                const w = Math.min(maxW, pageW);
                const aspect = img.naturalHeight / img.naturalWidth;
                const h = w * aspect;
                if (y + h > doc.internal.pageSize.getHeight() - 20) {
                  doc.addPage();
                  y = 20;
                }
                doc.addImage(dataUrl, format, 14, y, w, h);
                y += h + 8;
              }
            }
          }
        }
        const blob = doc.output('blob');
        const url = URL.createObjectURL(blob);
        const filename = spec.filename || 'report.pdf';
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        link.className = 'ai-pdf-download';
        link.textContent = 'Download ' + filename;
        pre.replaceWith(link);
      } catch (e) { showBlockError(pre, 'PDF', e); }
    }
  }

  // Run AI analysis
  async function runAI() {
    const sel = document.getElementById('ai-table-select');
    const input = document.getElementById('ai-input');
    const respDiv = document.getElementById('ai-response');
    if (!sel || !input || !respDiv) return;

    let selectedTables = [...sel.selectedOptions].map(o => o.value);
    const prompt = input.value.trim();

    // Default to all open tables if none selected
    if (selectedTables.length === 0) selectedTables = Object.keys(tables);
    if (!prompt) {
      setAIStatus('Enter a prompt.', 'error');
      return;
    }
    if (!_aiProvider) {
      setAIStatus('No AI provider available. Install Ollama or use Chrome with WebGPU.', 'error');
      showSetupHelp();
      return;
    }

    if (selectedTables.length > 0) flushAllSyncs();
    _aiConversation.push({ role: 'user', content: prompt });

    // Cap conversation history to limit token usage
    const MAX_AI_HISTORY = 20;
    if (_aiConversation.length > MAX_AI_HISTORY) {
      _aiConversation = _aiConversation.slice(-MAX_AI_HISTORY);
    }

    let systemPrompt;
    let systemPromptShort; // lighter prompt for SQL follow-up rounds (omits data context)
    if (selectedTables.length > 0) {
      const dataContext = buildDataContext(selectedTables);
      const tableList = selectedTables.map(n => {
        const t = tables[n];
        return `[${n}] (${t ? t.columns.join(', ') : 'unknown columns'})`;
      }).join(', ');
      const coreInstructions = `You are a data analyst. You have full access to a SQLite database containing the user's data.

IMPORTANT: To answer questions, you MUST write SQL queries. Write them in \`\`\`sql code blocks and they will be executed automatically. The results will be returned to you. Then use the results to answer the user's question.

Available tables: ${tableList}

Rules:
- ALWAYS query the data — never guess or estimate from the summary alone.
- Table and column names must be wrapped in square brackets, e.g. SELECT [column] FROM [table].
- You can run multiple queries, one per \`\`\`sql block.
- Use COUNT, SUM, AVG, GROUP BY, ORDER BY, JOINs, subqueries — any valid SQLite SQL.
- After receiving query results, give a clear, concise answer to the user.

Example — if the user asks "what are the top 5 products by revenue?", write:
\`\`\`sql
SELECT [product], SUM([revenue]) as total FROM [sales] GROUP BY [product] ORDER BY total DESC LIMIT 5
\`\`\`

${_richBlockPrompt}
${_aiImageContext()}`;

      systemPrompt = `${coreInstructions}

${dataContext}`;
      systemPromptShort = coreInstructions;
    } else {
      systemPrompt = `You are a helpful assistant. No data tables are currently loaded. Answer the user's question to the best of your ability. If the user asks about data analysis, let them know they can open CSV, Excel, or other data files to analyze.

${_richBlockPrompt}
${_aiImageContext()}`;
      systemPromptShort = systemPrompt;
    }

    // Append user message bubble
    const userMsg = document.createElement('div');
    userMsg.className = 'ai-msg ai-msg-user';
    userMsg.innerHTML = '<div class="ai-msg-bubble">' + escHtml(prompt) + '</div>';
    respDiv.appendChild(userMsg);

    // Save to history and clear input
    pushHistory(_aiHistory, prompt);
    _aiHistoryIdx = _aiHistory.length;
    _aiHistoryDraft = '';
    input.value = '';

    const t0 = performance.now();
    setAIStatus('Generating response... 0s', 'working');

    // Cancel any previous request
    if (_aiAbort) _aiAbort.abort();
    const abort = new AbortController();
    _aiAbort = abort;
    const signal = abort.signal;

    showInterruptButton(true, () => abort.abort());

    // Elapsed timer
    const aiTimer = setInterval(() => {
      const elapsed = ((performance.now() - t0) / 1000).toFixed(1);
      setAIStatus(`Generating response... ${elapsed}s`, 'working');
    }, 100);

    const MAX_SQL_ROUNDS = 5;

    try {
      for (let round = 0; round <= MAX_SQL_ROUNDS; round++) {
        const messages = [
          { role: 'system', content: round === 0 ? systemPrompt : systemPromptShort },
          ..._aiConversation,
        ];

        // Create AI response bubble
        const aiMsg = document.createElement('div');
        aiMsg.className = 'ai-msg ai-msg-ai';
        const aiBubble = document.createElement('div');
        aiBubble.className = 'ai-msg-bubble';
        aiMsg.appendChild(aiBubble);
        respDiv.appendChild(aiMsg);
        respDiv.scrollTop = respDiv.scrollHeight;

        let fullText = '';
        const onChunk = (chunk) => {
          const nearBottom = respDiv.scrollHeight - respDiv.scrollTop - respDiv.clientHeight < 40;
          fullText += chunk;
          aiBubble.innerHTML = formatAIResponse(fullText);
          if (nearBottom) respDiv.scrollTop = respDiv.scrollHeight;
        };

        if (_aiProvider === 'ollama') {
          await generateOllama(messages, onChunk, signal);
        } else if (_aiProvider === 'webllm') {
          await generateWebLLM(messages, onChunk, signal);
        } else if (_aiProvider === 'claude') {
          await generateClaude(messages, onChunk, signal);
        } else if (_aiProvider === 'openai') {
          await generateOpenAI(messages, onChunk, signal);
        }

        _aiConversation.push({ role: 'assistant', content: fullText });

        // Post-process rich blocks (charts, tables, PDFs)
        await postProcessAIBlocks(aiBubble);

        // Extract SQL blocks and execute them
        const sqlBlocks = [];
        fullText.replace(/```sql\n([\s\S]*?)```/g, (_, sql) => { sqlBlocks.push(sql.trim()); });

        if (sqlBlocks.length === 0 || round === MAX_SQL_ROUNDS) break;

        // Execute SQL queries and collect results
        let resultsText = '';
        for (const sql of sqlBlocks) {
          try {
            const results = db.exec(sql);
            if (results.length === 0) {
              resultsText += `Query: ${sql}\nResult: (no rows returned)\n\n`;
            } else {
              for (const r of results) {
                const header = r.columns.join(' | ');
                const rows = r.values.map(row => row.map(v => String(v ?? 'NULL')).join(' | '));
                // Limit to 50 result rows to stay within token budget
                const shown = rows.slice(0, 50);
                resultsText += `Query: ${sql}\n${header}\n${shown.join('\n')}`;
                if (rows.length > 50) resultsText += `\n... (${rows.length - 50} more rows)`;
                resultsText += '\n\n';
              }
            }
          } catch (e) {
            resultsText += `Query: ${sql}\nError: ${e.message}\n\n`;
          }
        }

        // Show query results as a system bubble
        const resultsMsg = document.createElement('div');
        resultsMsg.className = 'ai-msg ai-msg-ai';
        resultsMsg.innerHTML = '<div class="ai-msg-bubble"><pre style="margin:0;white-space:pre-wrap;font-size:12px;">' + escHtml(resultsText.trim()) + '</pre></div>';
        respDiv.appendChild(resultsMsg);
        const nearBottom = respDiv.scrollHeight - respDiv.scrollTop - respDiv.clientHeight < 40;
        if (nearBottom) respDiv.scrollTop = respDiv.scrollHeight;

        // Add results to conversation for next round
        _aiConversation.push({ role: 'user', content: 'SQL query results:\n\n' + resultsText });
      }

      clearInterval(aiTimer);
      showInterruptButton(false);
      const elapsed = performance.now() - t0;
      setAIStatus(`Done in ${formatElapsed(elapsed)}`, 'success');
    } catch (e) {
      clearInterval(aiTimer);
      showInterruptButton(false);
      // Remove trailing messages with no paired assistant response to avoid
      // consecutive same-role messages that some providers reject
      while (_aiConversation.length > 0 && _aiConversation[_aiConversation.length - 1].role !== 'assistant') {
        _aiConversation.pop();
      }
      if (e.name === 'AbortError') {
        setAIStatus('Cancelled', '');
      } else {
        setAIStatus(`Error: ${e.message}`, 'error');
        // Show error in a bubble
        const errMsg = document.createElement('div');
        errMsg.className = 'ai-msg ai-msg-ai';
        errMsg.innerHTML = '<div class="ai-msg-bubble"><span style="color:var(--danger)">' + escHtml(e.message) + '</span></div>';
        respDiv.appendChild(errMsg);
      }
    }
    if (_aiAbort === abort) _aiAbort = null;
  }

  // Ollama streaming
  let _ollamaReader = null;
  async function generateOllama(messages, onChunk, signal) {
    // Cancel any previous Ollama stream still running server-side
    if (_ollamaReader) {
      try { _ollamaReader.cancel(); } catch (_) {}
      _ollamaReader = null;
    }
    const r = await fetch(aiSettings.ollamaUrl + '/api/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ model: aiSettings.model, messages, stream: true }),
      signal,
    });
    if (!r.ok) throw new Error(`Ollama error: ${r.status} ${r.statusText}`);
    const reader = r.body.getReader();
    _ollamaReader = reader;
    const decoder = new TextDecoder();
    let buf = '';
    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buf += decoder.decode(value, { stream: true });
        const lines = buf.split('\n');
        buf = lines.pop();
        for (const line of lines) {
          if (!line.trim()) continue;
          try {
            const data = JSON.parse(line);
            if (data.message && data.message.content) onChunk(data.message.content);
          } catch {}
        }
      }
      // Process remaining buffer
      if (buf.trim()) {
        try {
          const data = JSON.parse(buf);
          if (data.message && data.message.content) onChunk(data.message.content);
        } catch {}
      }
    } finally {
      _ollamaReader = null;
    }
  }

  // WebLLM generation
  async function generateWebLLM(messages, onChunk, signal) {
    function checkAborted() {
      if (signal.aborted) throw new DOMException('Aborted', 'AbortError');
    }
    if (!_webllmEngine) {
      setAIStatus('Loading WebLLM engine (first time may download model)...', 'working');
      const webllm = await import('./lib/web-llm.js');
      checkAborted();
      const model = aiSettings.model || 'Qwen3-8B-q4f16_1-MLC';
      _webllmEngine = await webllm.CreateMLCEngine(model, {
        initProgressCallback: (progress) => {
          const pct = progress.progress != null ? Math.round(progress.progress * 100) + '%' : '';
          setAIStatus(`Loading model: ${progress.text || pct}`, 'working');
        },
      }, {
        context_window_size: 16384,
        sliding_window_size: -1,
      });
    }
    checkAborted();
    const response = await _webllmEngine.chat.completions.create({
      messages,
      stream: true,
      max_tokens: 4096,
    });
    const iterator = response[Symbol.asyncIterator]();
    try {
      while (true) {
        checkAborted();
        const { done, value } = await Promise.race([
          iterator.next(),
          new Promise((_, reject) => {
            if (signal.aborted) reject(new DOMException('Aborted', 'AbortError'));
            signal.addEventListener('abort', () => reject(new DOMException('Aborted', 'AbortError')), { once: true });
          }),
        ]);
        if (done) break;
        const delta = value.choices[0] && value.choices[0].delta && value.choices[0].delta.content;
        if (delta) onChunk(delta);
      }
    } catch (e) {
      if (e.name === 'AbortError') {
        // Destroy engine entirely — interruptGenerate/resetChat aren't reliable
        if (_webllmEngine) {
          try { _webllmEngine.unload(); } catch (_) {}
        }
        _webllmEngine = null;
      }
      throw e;
    }
  }

  // Claude streaming
  async function generateClaude(messages, onChunk, signal) {
    // Extract system message into top-level field (Anthropic format)
    let system = '';
    const filtered = [];
    for (const msg of messages) {
      if (msg.role === 'system') system = msg.content;
      else filtered.push(msg);
    }
    const r = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': aiSettings.claudeApiKey,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true',
      },
      body: JSON.stringify({
        model: aiSettings.model,
        max_tokens: 4096,
        system,
        messages: filtered,
        stream: true,
      }),
      signal,
    });
    if (!r.ok) {
      const errText = await r.text().catch(() => r.statusText);
      throw new Error(`Claude error: ${r.status} ${errText}`);
    }
    const reader = r.body.getReader();
    const decoder = new TextDecoder();
    let buf = '';
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buf += decoder.decode(value, { stream: true });
      const lines = buf.split('\n');
      buf = lines.pop();
      for (const line of lines) {
        if (!line.startsWith('data: ')) continue;
        try {
          const data = JSON.parse(line.slice(6));
          if (data.type === 'content_block_delta' && data.delta && data.delta.text) {
            onChunk(data.delta.text);
          }
        } catch {}
      }
    }
  }

  // OpenAI streaming
  async function generateOpenAI(messages, onChunk, signal) {
    const r = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + aiSettings.openaiApiKey,
      },
      body: JSON.stringify({
        model: aiSettings.model,
        messages,
        stream: true,
      }),
      signal,
    });
    if (!r.ok) {
      const errText = await r.text().catch(() => r.statusText);
      throw new Error(`OpenAI error: ${r.status} ${errText}`);
    }
    const reader = r.body.getReader();
    const decoder = new TextDecoder();
    let buf = '';
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buf += decoder.decode(value, { stream: true });
      const lines = buf.split('\n');
      buf = lines.pop();
      for (const line of lines) {
        if (!line.startsWith('data: ')) continue;
        const payload = line.slice(6).trim();
        if (payload === '[DONE]') return;
        try {
          const data = JSON.parse(payload);
          const content = data.choices && data.choices[0] && data.choices[0].delta && data.choices[0].delta.content;
          if (content) onChunk(content);
        } catch {}
      }
    }
  }

  // AI Settings modal
  function showAISettings() {
    const overlay = document.createElement('div');
    overlay.className = 'modal-overlay';

    const modelList = aiSettings.model ? aiSettings.model : '';
    const inputStyle = 'width:100%;box-sizing:border-box;background:var(--bg);border:1px solid var(--border);color:var(--text);padding:6px 10px;border-radius:4px;font-family:inherit;font-size:13px;margin-bottom:10px;outline:none;';
    overlay.innerHTML = `
      <div class="modal" style="min-width:380px">
        <h3>AI Settings</h3>
        <label style="font-size:12px;display:block;margin-bottom:4px;">Provider</label>
        <select id="ai-set-provider" style="${inputStyle}">
          <option value="auto" ${aiSettings.provider === 'auto' ? 'selected' : ''}>Auto-detect</option>
          <option value="claude" ${aiSettings.provider === 'claude' ? 'selected' : ''}>Claude (Anthropic)</option>
          <option value="openai" ${aiSettings.provider === 'openai' ? 'selected' : ''}>OpenAI</option>
          <option value="ollama" ${aiSettings.provider === 'ollama' ? 'selected' : ''}>Ollama</option>
          <option value="webllm" ${aiSettings.provider === 'webllm' ? 'selected' : ''}>WebLLM (in-browser)</option>
        </select>
        <div id="ai-set-ollama-fields">
          <label style="font-size:12px;display:block;margin-bottom:4px;">Ollama URL</label>
          <input type="text" id="ai-set-url" value="${escHtml(aiSettings.ollamaUrl)}" style="${inputStyle}">
        </div>
        <div id="ai-set-claude-fields">
          <label style="font-size:12px;display:block;margin-bottom:4px;">Claude API Key</label>
          <div style="display:flex;gap:6px;margin-bottom:10px;">
            <input type="password" id="ai-set-claude-key" value="${escHtml(aiSettings.claudeApiKey)}" placeholder="sk-ant-..." style="flex:1;background:var(--bg);border:1px solid var(--border);color:var(--text);padding:6px 10px;border-radius:4px;font-family:inherit;font-size:13px;outline:none;">
            <button class="ai-key-toggle" data-target="ai-set-claude-key" style="background:var(--surface2);border:1px solid var(--border);color:var(--text);padding:6px 8px;border-radius:4px;cursor:pointer;font-size:12px;">Show</button>
          </div>
        </div>
        <div id="ai-set-openai-fields">
          <label style="font-size:12px;display:block;margin-bottom:4px;">OpenAI API Key</label>
          <div style="display:flex;gap:6px;margin-bottom:10px;">
            <input type="password" id="ai-set-openai-key" value="${escHtml(aiSettings.openaiApiKey)}" placeholder="sk-..." style="flex:1;background:var(--bg);border:1px solid var(--border);color:var(--text);padding:6px 10px;border-radius:4px;font-family:inherit;font-size:13px;outline:none;">
            <button class="ai-key-toggle" data-target="ai-set-openai-key" style="background:var(--surface2);border:1px solid var(--border);color:var(--text);padding:6px 8px;border-radius:4px;cursor:pointer;font-size:12px;">Show</button>
          </div>
        </div>
        <label style="font-size:12px;display:block;margin-bottom:4px;">Model</label>
        <div style="display:flex;gap:6px;margin-bottom:10px;">
          <select id="ai-set-model" style="flex:1;background:var(--bg);border:1px solid var(--border);color:var(--text);padding:6px 10px;border-radius:4px;font-family:inherit;font-size:13px;outline:none;">
            ${modelList ? `<option value="${escHtml(modelList)}" selected>${escHtml(modelList)}</option>` : '<option value="">Loading...</option>'}
          </select>
          <button id="ai-set-refresh" style="background:var(--surface2);border:1px solid var(--border);color:var(--text);padding:6px 10px;border-radius:4px;cursor:pointer;font-family:inherit;font-size:12px;">Refresh</button>
        </div>
        <div class="modal-buttons">
          <button class="cancel">Cancel</button>
          <button class="primary ok">Save</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    const providerSel = overlay.querySelector('#ai-set-provider');
    const urlInput = overlay.querySelector('#ai-set-url');
    const modelSel = overlay.querySelector('#ai-set-model');
    const refreshBtn = overlay.querySelector('#ai-set-refresh');
    const ollamaFields = overlay.querySelector('#ai-set-ollama-fields');
    const claudeFields = overlay.querySelector('#ai-set-claude-fields');
    const openaiFields = overlay.querySelector('#ai-set-openai-fields');

    // Show/hide toggle for API key fields
    overlay.querySelectorAll('.ai-key-toggle').forEach(btn => {
      btn.addEventListener('click', () => {
        const inp = overlay.querySelector('#' + btn.dataset.target);
        if (inp.type === 'password') { inp.type = 'text'; btn.textContent = 'Hide'; }
        else { inp.type = 'password'; btn.textContent = 'Show'; }
      });
    });

    function updateFieldVisibility() {
      const p = providerSel.value;
      ollamaFields.style.display = (p === 'ollama' || p === 'auto') ? '' : 'none';
      claudeFields.style.display = (p === 'claude') ? '' : 'none';
      openaiFields.style.display = (p === 'openai') ? '' : 'none';
    }

    async function refreshModels() {
      const provider = providerSel.value;
      updateFieldVisibility();
      modelSel.innerHTML = '<option value="">Loading...</option>';

      if (provider === 'claude') {
        const claudeModels = [
          'claude-opus-4-20250514',
          'claude-sonnet-4-20250514',
          'claude-haiku-4-20250414',
        ];
        modelSel.innerHTML = claudeModels.map(m =>
          `<option value="${escHtml(m)}" ${m === aiSettings.model ? 'selected' : ''}>${escHtml(m)}</option>`
        ).join('');
        return;
      }
      if (provider === 'openai') {
        const openaiModels = [
          'o3', 'o3-mini',
          'o4-mini',
          'gpt-4.1', 'gpt-4.1-mini', 'gpt-4.1-nano',
          'gpt-4o', 'gpt-4o-mini',
        ];
        modelSel.innerHTML = openaiModels.map(m =>
          `<option value="${escHtml(m)}" ${m === aiSettings.model ? 'selected' : ''}>${escHtml(m)}</option>`
        ).join('');
        return;
      }
      if (provider === 'ollama' || provider === 'auto') {
        try {
          const r = await fetch(urlInput.value + '/api/tags', { signal: AbortSignal.timeout(3000) });
          if (r.ok) {
            const data = await r.json();
            const models = (data.models || []).map(m => m.name);
            modelSel.innerHTML = models.map(m =>
              `<option value="${escHtml(m)}" ${m === aiSettings.model ? 'selected' : ''}>${escHtml(m)}</option>`
            ).join('') || '<option value="">No models found</option>';
            return;
          }
        } catch {}
      }
      if (provider === 'webllm' || provider === 'auto') {
        const webllmModels = [
          'Qwen3-8B-q4f16_1-MLC',
          'Qwen3-4B-q4f16_1-MLC',
          'Qwen3-1.7B-q4f16_1-MLC',
          'Qwen3-0.6B-q4f16_1-MLC',
          'Llama-3.1-8B-Instruct-q4f16_1-MLC',
          'Llama-3.2-3B-Instruct-q4f16_1-MLC',
          'Llama-3.2-1B-Instruct-q4f16_1-MLC',
          'DeepSeek-R1-Distill-Llama-8B-q4f16_1-MLC',
          'DeepSeek-R1-Distill-Qwen-7B-q4f16_1-MLC',
          'Phi-3.5-mini-instruct-q4f16_1-MLC',
          'Mistral-7B-Instruct-v0.3-q4f16_1-MLC',
          'Hermes-3-Llama-3.1-8B-q4f16_1-MLC',
          'gemma-2-9b-it-q4f16_1-MLC',
          'gemma-2-2b-it-q4f16_1-MLC',
          'Qwen2.5-7B-Instruct-q4f16_1-MLC',
          'Qwen2.5-3B-Instruct-q4f16_1-MLC',
          'Qwen2.5-1.5B-Instruct-q4f16_1-MLC',
          'Qwen2.5-Coder-7B-Instruct-q4f16_1-MLC',
          'SmolLM2-1.7B-Instruct-q4f16_1-MLC',
          'SmolLM2-360M-Instruct-q4f16_1-MLC',
        ];
        modelSel.innerHTML = webllmModels.map(m =>
          `<option value="${escHtml(m)}" ${m === aiSettings.model ? 'selected' : ''}>${escHtml(m)}</option>`
        ).join('');
        return;
      }
      modelSel.innerHTML = '<option value="">No provider available</option>';
    }

    refreshBtn.addEventListener('click', refreshModels);
    providerSel.addEventListener('change', () => {
      // Reset model when switching provider type
      aiSettings.model = '';
      refreshModels();
    });
    refreshModels();

    const close = (save) => {
      if (save) {
        aiSettings.provider = providerSel.value;
        aiSettings.ollamaUrl = urlInput.value.replace(/\/+$/, '');
        aiSettings.claudeApiKey = (overlay.querySelector('#ai-set-claude-key').value || '').trim();
        aiSettings.openaiApiKey = (overlay.querySelector('#ai-set-openai-key').value || '').trim();
        aiSettings.model = modelSel.value;
        saveAISettings();
        _webllmEngine = null; // Reset engine on settings change
        detectAIProvider();
      }
      overlay.remove();
    };
    overlay.querySelector('.cancel').addEventListener('click', () => close(false));
    overlay.querySelector('.ok').addEventListener('click', () => close(true));
    overlay.addEventListener('click', (e) => { if (e.target === overlay) close(false); });
  }

  // Render image thumbnails in the AI input area
  function renderImageThumbs() {
    let strip = document.getElementById('ai-image-strip');
    if (!strip) {
      strip = document.createElement('div');
      strip.id = 'ai-image-strip';
      const inputRow = document.getElementById('ai-input-row');
      inputRow.parentNode.insertBefore(strip, inputRow);
    }
    const names = Object.keys(_aiImages);
    if (names.length === 0) { strip.innerHTML = ''; strip.style.display = 'none'; return; }
    strip.style.display = '';
    strip.innerHTML = '';
    for (const name of names) {
      const thumb = document.createElement('div');
      thumb.className = 'ai-image-thumb';
      thumb.innerHTML = `<img src="${_aiImages[name]}" alt="${escHtml(name)}" title="${escHtml(name)}"><span class="ai-image-name">${escHtml(name)}</span><button class="ai-image-remove" title="Remove">&times;</button>`;
      thumb.querySelector('.ai-image-remove').addEventListener('click', () => {
        delete _aiImages[name];
        renderImageThumbs();
      });
      strip.appendChild(thumb);
    }
  }

  // Read an image file and add to _aiImages
  function addAIImage(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['png', 'jpg', 'jpeg', 'gif', 'webp', 'svg', 'bmp'].includes(ext)) {
      setAIStatus('Unsupported image format: .' + ext, 'error');
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      _aiImages[file.name] = reader.result;
      renderImageThumbs();
      setAIStatus(`Image "${file.name}" added`, 'success');
    };
    reader.readAsDataURL(file);
  }

  function setupAI() {
    setupConsoleTabs();
    document.getElementById('ai-settings-btn').addEventListener('click', showAISettings);
  }

  // ---- Expression Language ----

  function exprTokenize(src) {
    const tokens = [];
    let i = 0;
    while (i < src.length) {
      if (src[i] === ' ' || src[i] === '\t' || src[i] === '\r' || src[i] === '\n') { i++; continue; }
      if (src[i] === "'" || src[i] === '"') {
        const q = src[i]; let s = ''; i++;
        while (i < src.length && src[i] !== q) {
          if (src[i] === '\\' && i + 1 < src.length) {
            i++;
            if (src[i] === 'n') s += '\n'; else if (src[i] === 't') s += '\t'; else s += src[i];
            i++;
          } else { s += src[i]; i++; }
        }
        if (i < src.length) i++;
        tokens.push({ type: 'string', value: s }); continue;
      }
      if ((src[i] >= '0' && src[i] <= '9') || (src[i] === '.' && i + 1 < src.length && src[i + 1] >= '0' && src[i + 1] <= '9')) {
        let n = '';
        while (i < src.length && ((src[i] >= '0' && src[i] <= '9') || src[i] === '.')) { n += src[i]; i++; }
        tokens.push({ type: 'number', value: parseFloat(n) }); continue;
      }
      if ((src[i] >= 'a' && src[i] <= 'z') || (src[i] >= 'A' && src[i] <= 'Z') || src[i] === '_') {
        let id = '';
        while (i < src.length && ((src[i] >= 'a' && src[i] <= 'z') || (src[i] >= 'A' && src[i] <= 'Z') || (src[i] >= '0' && src[i] <= '9') || src[i] === '_')) { id += src[i]; i++; }
        if (id === 'true') tokens.push({ type: 'boolean', value: true });
        else if (id === 'false') tokens.push({ type: 'boolean', value: false });
        else if (id === 'null') tokens.push({ type: 'null', value: null });
        else tokens.push({ type: 'ident', value: id });
        continue;
      }
      const two = src.substring(i, i + 2);
      if (two === '==' || two === '!=' || two === '<=' || two === '>=' || two === '&&' || two === '||') {
        tokens.push({ type: 'op', value: two }); i += 2; continue;
      }
      const ch = src[i];
      if ('+-*/%!<>?.,:()'.includes(ch)) {
        tokens.push({ type: 'op', value: ch }); i++; continue;
      }
      throw new Error(`Unexpected character '${ch}' at position ${i}`);
    }
    tokens.push({ type: 'eof' });
    return tokens;
  }

  function exprParse(tokens) {
    let pos = 0;
    function peek() { return tokens[pos]; }
    function eat(type, value) {
      const t = tokens[pos];
      if (type && t.type !== type) throw new Error(`Expected ${type} but got ${t.type} '${t.value}'`);
      if (value !== undefined && t.value !== value) throw new Error(`Expected '${value}' but got '${t.value}'`);
      pos++; return t;
    }
    function match(type, value) {
      const t = tokens[pos];
      if (t.type === type && (value === undefined || t.value === value)) { pos++; return t; }
      return null;
    }

    function parseExpr() { return parseTernary(); }

    function parseTernary() {
      let node = parseOr();
      if (match('op', '?')) {
        const then = parseExpr();
        eat('op', ':');
        const els = parseExpr();
        node = { type: 'ternary', cond: node, then, els };
      }
      return node;
    }

    function parseOr() {
      let node = parseAnd();
      while (match('op', '||')) { node = { type: 'binary', op: '||', left: node, right: parseAnd() }; }
      return node;
    }

    function parseAnd() {
      let node = parseEquality();
      while (match('op', '&&')) { node = { type: 'binary', op: '&&', left: node, right: parseEquality() }; }
      return node;
    }

    function parseEquality() {
      let node = parseComparison();
      while (peek().type === 'op' && (peek().value === '==' || peek().value === '!=')) {
        const op = eat('op').value;
        node = { type: 'binary', op, left: node, right: parseComparison() };
      }
      return node;
    }

    function parseComparison() {
      let node = parseAddition();
      if (peek().type === 'op' && (peek().value === '<' || peek().value === '>' || peek().value === '<=' || peek().value === '>=')) {
        const op = eat('op').value;
        node = { type: 'binary', op, left: node, right: parseAddition() };
      }
      return node;
    }

    function parseAddition() {
      let node = parseMultiplication();
      while (peek().type === 'op' && (peek().value === '+' || peek().value === '-')) {
        const op = eat('op').value;
        node = { type: 'binary', op, left: node, right: parseMultiplication() };
      }
      return node;
    }

    function parseMultiplication() {
      let node = parseUnary();
      while (peek().type === 'op' && (peek().value === '*' || peek().value === '/' || peek().value === '%')) {
        const op = eat('op').value;
        node = { type: 'binary', op, left: node, right: parseUnary() };
      }
      return node;
    }

    function parseUnary() {
      if (match('op', '!')) return { type: 'unary', op: '!', operand: parseUnary() };
      if (match('op', '-')) return { type: 'unary', op: '-', operand: parseUnary() };
      return parsePostfix();
    }

    function parsePostfix() {
      let node = parseAtom();
      while (match('op', '.')) {
        const prop = eat('ident').value;
        node = { type: 'member', object: node, property: prop };
      }
      return node;
    }

    function parseAtom() {
      const t = peek();
      if (t.type === 'number') { eat('number'); return { type: 'literal', value: t.value }; }
      if (t.type === 'string') { eat('string'); return { type: 'literal', value: t.value }; }
      if (t.type === 'boolean') { eat('boolean'); return { type: 'literal', value: t.value }; }
      if (t.type === 'null') { eat('null'); return { type: 'literal', value: null }; }
      if (t.type === 'ident') {
        eat('ident');
        if (match('op', '(')) {
          const args = [];
          if (peek().value !== ')') {
            args.push(parseExpr());
            while (match('op', ',')) args.push(parseExpr());
          }
          eat('op', ')');
          return { type: 'call', name: t.value, args };
        }
        return { type: 'var', name: t.value };
      }
      if (match('op', '(')) {
        const node = parseExpr();
        eat('op', ')');
        return node;
      }
      throw new Error(`Unexpected token '${t.value}' at position ${pos}`);
    }

    const ast = parseExpr();
    if (peek().type !== 'eof') throw new Error(`Unexpected token '${peek().value}' after expression`);
    return ast;
  }

  function exprCompile(src) {
    return exprParse(exprTokenize(src));
  }

  const _exprFunctions = {
    upper(s) { return String(s ?? '').toUpperCase(); },
    lower(s) { return String(s ?? '').toLowerCase(); },
    trim(s) { return String(s ?? '').trim(); },
    len(s) { return String(s ?? '').length; },
    substr(s, start, length) {
      s = String(s ?? '');
      return length === undefined ? s.substring(start) : s.substring(start, start + length);
    },
    replace(s, search, repl) { return String(s ?? '').replace(String(search), String(repl ?? '')); },
    replaceAll(s, search, repl) { return String(s ?? '').split(String(search)).join(String(repl ?? '')); },
    startsWith(s, prefix) { return String(s ?? '').startsWith(String(prefix ?? '')); },
    endsWith(s, suffix) { return String(s ?? '').endsWith(String(suffix ?? '')); },
    contains(s, sub) { return String(s ?? '').includes(String(sub ?? '')); },
    padLeft(s, width, ch) { return String(s ?? '').padStart(width, String(ch ?? ' ')); },
    padRight(s, width, ch) { return String(s ?? '').padEnd(width, String(ch ?? ' ')); },
    concat(...args) { return args.map(a => String(a ?? '')).join(''); },
    repeat(s, count) { return String(s ?? '').repeat(Math.max(0, Math.floor(count) || 0)); },

    num(s) { const n = Number(s); return isNaN(n) ? null : n; },
    fixed(n, d) { n = Number(n); return isNaN(n) ? '' : n.toFixed(d ?? 0); },
    round(n, d) { n = Number(n); if (isNaN(n)) return null; if (d === undefined) return Math.round(n); const f = Math.pow(10, d); return Math.round(n * f) / f; },
    floor(n) { n = Number(n); return isNaN(n) ? null : Math.floor(n); },
    ceil(n) { n = Number(n); return isNaN(n) ? null : Math.ceil(n); },
    abs(n) { n = Number(n); return isNaN(n) ? null : Math.abs(n); },
    min(a, b) { return Math.min(Number(a), Number(b)); },
    max(a, b) { return Math.max(Number(a), Number(b)); },
    commas(n) { n = Number(n); return isNaN(n) ? '' : n.toLocaleString(); },

    date(s, fmt) {
      if (!s) return '';
      const sStr = String(s);
      let d, nanoFrac = '';
      d = new Date(sStr);
      if (isNaN(d.getTime())) {
        if (/^\d+\.\d+$/.test(sStr)) {
          const dotIdx = sStr.indexOf('.');
          const secPart = sStr.slice(0, dotIdx);
          const fracPart = sStr.slice(dotIdx + 1).padEnd(9, '0').slice(0, 9);
          nanoFrac = fracPart;
          d = new Date(parseInt(secPart) * 1000 + parseInt(fracPart.slice(0, 3)));
        } else if (/^\d+$/.test(sStr)) {
          d = new Date(Number(sStr) * 1000);
        }
      }
      if (isNaN(d.getTime())) return sStr;
      if (fmt === 'full') {
        const y = String(d.getFullYear());
        const mo = String(d.getMonth() + 1).padStart(2, '0');
        const dy = String(d.getDate()).padStart(2, '0');
        const h = String(d.getHours()).padStart(2, '0');
        const mi = String(d.getMinutes()).padStart(2, '0');
        const sc = String(d.getSeconds()).padStart(2, '0');
        const frac = nanoFrac || String(d.getMilliseconds()).padStart(3, '0') + '000000';
        return y + '-' + mo + '-' + dy + ' ' + h + ':' + mi + ':' + sc + '.' + frac;
      }
      if (!fmt || fmt === 'locale') return d.toLocaleDateString();
      if (fmt === 'iso') return d.toISOString().substring(0, 10);
      if (fmt === 'time') return d.toLocaleTimeString();
      if (fmt === 'datetime') return d.toLocaleString();
      const yyyy = String(d.getFullYear());
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      const dd = String(d.getDate()).padStart(2, '0');
      return fmt.replace('YYYY', yyyy).replace('MM', mm).replace('DD', dd);
    },

    'if': function(cond, then, els) { return cond ? then : els; },
    choose(val, ...pairs) {
      const sv = String(val ?? '');
      const hasDefault = pairs.length % 2 === 1;
      const limit = hasDefault ? pairs.length - 1 : pairs.length;
      for (let i = 0; i < limit; i += 2) {
        if (sv === String(pairs[i] ?? '')) return pairs[i + 1];
      }
      return hasDefault ? pairs[pairs.length - 1] : val;
    },
    coalesce(...args) {
      for (const a of args) { if (a !== null && a !== undefined && a !== '') return a; }
      return args.length ? args[args.length - 1] : null;
    },
    isEmpty(s) { return s === null || s === undefined || s === ''; },
    isNum(s) { return s !== null && s !== undefined && s !== '' && !isNaN(Number(s)); },
  };

  const _exprBlockedProps = new Set(['constructor', '__proto__', 'prototype', '__defineGetter__', '__defineSetter__', '__lookupGetter__', '__lookupSetter__']);

  function exprEval(node, env) {
    switch (node.type) {
      case 'literal': return node.value;
      case 'var':
        if (node.name === 'value') return env.value;
        if (node.name === 'column') return env.column;
        if (node.name === 'table') return env.table;
        if (node.name === 'row') return env.row;
        throw new Error(`Unknown variable '${node.name}'`);
      case 'member': {
        const obj = exprEval(node.object, env);
        if (obj === null || obj === undefined) return null;
        if (_exprBlockedProps.has(node.property)) return null;
        if (typeof obj === 'object' && Object.prototype.hasOwnProperty.call(obj, node.property)) return obj[node.property];
        return null;
      }
      case 'unary':
        if (node.op === '!') { const v = exprEval(node.operand, env); return !v; }
        if (node.op === '-') { return -Number(exprEval(node.operand, env)); }
        return null;
      case 'binary': {
        if (node.op === '||') { const l = exprEval(node.left, env); return l ? l : exprEval(node.right, env); }
        if (node.op === '&&') { const l = exprEval(node.left, env); return l ? exprEval(node.right, env) : l; }
        const left = exprEval(node.left, env);
        const right = exprEval(node.right, env);
        switch (node.op) {
          case '+': {
            if (typeof left === 'number' && typeof right === 'number') return left + right;
            return String(left ?? '') + String(right ?? '');
          }
          case '-': return Number(left) - Number(right);
          case '*': return Number(left) * Number(right);
          case '/': { const d = Number(right); return d === 0 ? null : Number(left) / d; }
          case '%': { const d = Number(right); return d === 0 ? null : Number(left) % d; }
          case '==': return String(left ?? '') === String(right ?? '');
          case '!=': return String(left ?? '') !== String(right ?? '');
          case '<': {
            const nl = Number(left), nr = Number(right);
            if (!isNaN(nl) && !isNaN(nr)) return nl < nr;
            return String(left ?? '') < String(right ?? '');
          }
          case '>': {
            const nl = Number(left), nr = Number(right);
            if (!isNaN(nl) && !isNaN(nr)) return nl > nr;
            return String(left ?? '') > String(right ?? '');
          }
          case '<=': {
            const nl = Number(left), nr = Number(right);
            if (!isNaN(nl) && !isNaN(nr)) return nl <= nr;
            return String(left ?? '') <= String(right ?? '');
          }
          case '>=': {
            const nl = Number(left), nr = Number(right);
            if (!isNaN(nl) && !isNaN(nr)) return nl >= nr;
            return String(left ?? '') >= String(right ?? '');
          }
        }
        return null;
      }
      case 'ternary': return exprEval(node.cond, env) ? exprEval(node.then, env) : exprEval(node.els, env);
      case 'call': {
        const fn = _exprFunctions[node.name];
        if (!fn) throw new Error(`Unknown function '${node.name}'`);
        const args = node.args.map(a => exprEval(a, env));
        return fn(...args);
      }
    }
    return null;
  }

  function exprEvalToString(ast, env) {
    try {
      const result = exprEval(ast, env);
      if (result === null || result === undefined) return '';
      if (typeof result === 'boolean') return result ? 'true' : 'false';
      return String(result);
    } catch (e) {
      console.warn('Plugin expression error:', e.message);
      return String(env.value ?? '');
    }
  }

  // ---- Plugin System ----

  function validatePlugin(config) {
    const errors = [];
    if (!config || typeof config !== 'object') { errors.push('Plugin must be a JSON object'); return errors; }
    if (typeof config.name !== 'string' || !config.name.trim()) errors.push('Plugin "name" is required');
    if (typeof config.table !== 'string') errors.push('Plugin "table" regex is required');
    else { try { new RegExp(config.table); } catch (e) { errors.push(`Invalid table regex: ${e.message}`); } }
    if (!Array.isArray(config.columns) || config.columns.length === 0) errors.push('Plugin "columns" must be a non-empty array');
    else {
      config.columns.forEach((col, i) => {
        if (typeof col.match !== 'string') errors.push(`columns[${i}].match is required`);
        else { try { new RegExp(col.match); } catch (e) { errors.push(`columns[${i}].match invalid regex: ${e.message}`); } }
        if (typeof col.display !== 'string') errors.push(`columns[${i}].display expression is required`);
      });
    }
    if ('version' in config && typeof config.version !== 'string') errors.push('"version" must be a string');
    if ('author' in config && typeof config.author !== 'string') errors.push('"author" must be a string');
    if ('created' in config && typeof config.created !== 'string') errors.push('"created" must be a string');
    return errors;
  }

  function compilePlugin(config) {
    let tableRe;
    try { tableRe = new RegExp('^(?:' + config.table + ')$'); } catch (e) { return null; }
    const columns = [];
    for (const col of config.columns) {
      try {
        const matchRe = new RegExp('^(?:' + col.match + ')$');
        const displayAst = exprCompile(col.display);
        columns.push({ matchRe, displayAst });
      } catch (e) {
        console.warn(`Plugin "${config.name}": skipping column rule "${col.match}": ${e.message}`);
      }
    }
    return { tableRe, columns };
  }

  function rebuildTransformCache() {
    _columnTransformCache = {};
    for (const tableName of Object.keys(tables)) {
      rebuildTransformCacheForTable(tableName);
    }
  }

  function rebuildTransformCacheForTable(tableName) {
    const t = tables[tableName];
    if (!t) { delete _columnTransformCache[tableName]; return; }
    const colMap = {};
    for (let pi = 0; pi < plugins.length; pi++) {
      const compiled = plugins[pi]._compiled;
      if (!compiled || !compiled.tableRe.test(tableName)) continue;
      for (const colName of t.columns) {
        for (const rule of compiled.columns) {
          if (rule.matchRe.test(colName)) {
            colMap[colName] = { displayAst: rule.displayAst, pluginIdx: pi };
            break;
          }
        }
      }
    }
    if (Object.keys(colMap).length > 0) _columnTransformCache[tableName] = colMap;
    else delete _columnTransformCache[tableName];
  }

  function getDisplayValue(tableName, column, row) {
    const tableCache = _columnTransformCache[tableName];
    if (!tableCache) return row[column] ?? '';
    const entry = tableCache[column];
    if (!entry) return row[column] ?? '';
    return exprEvalToString(entry.displayAst, { value: row[column] ?? '', column, row, table: tableName });
  }

  function hasDisplayTransform(tableName, column) {
    const tableCache = _columnTransformCache[tableName];
    return !!(tableCache && tableCache[column]);
  }

  async function loadPluginFile(file) {
    try {
      const text = await file.text();
      const config = JSON.parse(text);
      const errors = validatePlugin(config);
      if (errors.length > 0) { showToast('Plugin error: ' + errors.join('; '), 'error'); return; }
      const compiled = compilePlugin(config);
      if (!compiled || compiled.columns.length === 0) { showToast('Plugin has no valid column rules.', 'error'); return; }
      config._compiled = compiled;
      config._filename = file.name;
      plugins.push(config);
      rebuildTransformCache();
      rerenderAllWindows();
      persistPlugins();
      updatePluginMenu();
      showToast('Plugin "' + (config.name || file.name) + '" loaded');
    } catch (e) {
      showToast('Failed to load plugin: ' + e.message, 'error');
    }
  }

  async function loadPluginFromFile() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json';
    input.addEventListener('change', () => {
      if (input.files && input.files[0]) loadPluginFile(input.files[0]);
    });
    input.click();
  }

  function unloadPlugin(index) {
    if (index < 0 || index >= plugins.length) return;
    const pluginName = plugins[index].name || 'Plugin ' + (index + 1);
    closePluginPopover();
    plugins.splice(index, 1);
    rebuildTransformCache();
    rerenderAllWindows();
    persistPlugins();
    updatePluginMenu();
    showToast('Plugin "' + pluginName + '" unloaded');
  }

  function rerenderAllWindows() {
    for (const win of windows) {
      if (win.tableName && tables[win.tableName]) {
        rebuildTable(win);
      }
    }
  }

  function persistPlugins() {
    const configs = plugins.map(p => {
      const { _compiled, ...rest } = p;
      return rest;
    });
    localStorage.setItem('csvsql_plugins', JSON.stringify(configs));
  }

  function loadPersistedPlugins() {
    try {
      const data = JSON.parse(localStorage.getItem('csvsql_plugins') || 'null');
      if (!Array.isArray(data)) return;
      for (const config of data) {
        const errors = validatePlugin(config);
        if (errors.length > 0) continue;
        const compiled = compilePlugin(config);
        if (!compiled || compiled.columns.length === 0) continue;
        config._compiled = compiled;
        plugins.push(config);
      }
      if (plugins.length > 0) rebuildTransformCache();
    } catch (e) {
      console.warn('Failed to load persisted plugins:', e.message);
    }
  }

  function updatePluginMenu() {
    const area = document.getElementById('plugin-list-area');
    if (!area) return;
    if (plugins.length === 0) {
      area.innerHTML = '<span class="menu-hint">No plugins loaded</span>';
      return;
    }
    area.innerHTML = '';
    plugins.forEach((p, i) => {
      const entry = document.createElement('div');
      entry.className = 'plugin-entry';

      const unloadSpan = document.createElement('span');
      unloadSpan.className = 'plugin-unload';
      unloadSpan.textContent = '✕';
      unloadSpan.title = 'Unload plugin';
      unloadSpan.addEventListener('click', (e) => {
        e.stopPropagation();
        unloadPlugin(i);
      });

      const nameSpan = document.createElement('span');
      nameSpan.className = 'plugin-name';
      nameSpan.textContent = p.name || p._filename || 'Plugin ' + (i + 1);
      nameSpan.addEventListener('click', (e) => {
        e.stopPropagation();
        e.preventDefault();
        showPluginAbout(p, i);
      });

      entry.appendChild(unloadSpan);
      entry.appendChild(nameSpan);
      area.appendChild(entry);
    });
  }

  function closePluginPopover() {
    if (!_activePluginPopover) return;
    if (_activePluginPopover._onEscape) document.removeEventListener('keydown', _activePluginPopover._onEscape);
    _activePluginPopover.remove();
    _activePluginPopover = null;
  }

  function showPluginAbout(plugin, pluginIndex) {
    closePluginPopover();

    const name = escHtml(plugin.name || 'Untitled Plugin');
    const version = plugin.version ? escHtml(plugin.version) : null;
    const author = plugin.author ? escHtml(plugin.author) : null;
    const created = plugin.created ? escHtml(plugin.created) : null;
    const description = plugin.description ? escHtml(plugin.description) : null;
    const tablePattern = escHtml(plugin.table || '');

    let html = '<button class="modal-close">✕</button>';
    html += '<h3>' + name + '</h3>';

    const metaLines = [];
    if (version) metaLines.push('<strong>Version:</strong> ' + version);
    if (author) metaLines.push('<strong>Author:</strong> ' + author);
    if (created) metaLines.push('<strong>Created:</strong> ' + created);
    if (metaLines.length > 0) {
      html += '<div class="plugin-popover-meta">' + metaLines.join('<br>') + '</div>';
    }

    if (description) {
      html += '<div class="plugin-popover-desc">' + description + '</div>';
    }

    html += '<div class="plugin-popover-section"><strong>Table pattern:</strong> <code>' + tablePattern + '</code></div>';

    html += '<div class="plugin-popover-section"><strong>Column rules:</strong></div>';
    html += '<div class="plugin-popover-rules">';
    for (const col of plugin.columns) {
      html += '<div class="plugin-popover-rule">';
      html += '<span class="plugin-rule-match"><code>' + escHtml(col.match) + '</code></span>';
      html += '<span class="plugin-rule-arrow">→</span>';
      html += '<span class="plugin-rule-display"><code>' + escHtml(col.display) + '</code></span>';
      html += '</div>';
    }
    html += '</div>';

    html += '<div class="modal-buttons"><button class="modal-unload">Unload Plugin</button><button class="modal-cancel">Close</button></div>';

    const overlay = document.createElement('div');
    overlay.className = 'modal-overlay';
    const modal = document.createElement('div');
    modal.className = 'modal';
    modal.innerHTML = html;

    modal.querySelector('.modal-close').addEventListener('click', closePluginPopover);
    modal.querySelector('.modal-cancel').addEventListener('click', closePluginPopover);
    modal.querySelector('.modal-unload').addEventListener('click', () => {
      closePluginPopover();
      unloadPlugin(pluginIndex);
    });
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) closePluginPopover();
    });
    overlay._onEscape = (e) => { if (e.key === 'Escape') closePluginPopover(); };
    document.addEventListener('keydown', overlay._onEscape);

    overlay.appendChild(modal);
    document.body.appendChild(overlay);
    _activePluginPopover = overlay;
  }

  function showExpressionReference() {
    showHelpWindow('Plugin Expression Reference', `
<h4>Overview</h4>
<p>Plugin display expressions use the CSVSQL expression language &mdash; a safe, sandboxed language with no access to JavaScript globals, the DOM, or browser APIs.</p>

<h4>Variables</h4>
<table>
<tr><th>Variable</th><th>Type</th><th>Description</th></tr>
<tr><td><code>value</code></td><td>String</td><td>The raw cell value</td></tr>
<tr><td><code>column</code></td><td>String</td><td>The column name</td></tr>
<tr><td><code>table</code></td><td>String</td><td>The table name</td></tr>
<tr><td><code>row</code></td><td>Object</td><td>The full row &mdash; access fields with <code>row.fieldname</code></td></tr>
</table>

<h4>Data Types</h4>
<p>String (<code>'hello'</code>, <code>"world"</code>), Number (<code>42</code>, <code>3.14</code>), Boolean (<code>true</code>, <code>false</code>), Null (<code>null</code>). All values coerce to string for final output.</p>

<h4>Operators (by precedence, lowest first)</h4>
<table>
<tr><th>Operator</th><th>Description</th><th>Example</th></tr>
<tr><td><code>? :</code></td><td>Ternary conditional</td><td><code>value == '' ? 'N/A' : value</code></td></tr>
<tr><td><code>||</code></td><td>Logical OR (short-circuit, returns first truthy)</td><td><code>value || 'default'</code></td></tr>
<tr><td><code>&&</code></td><td>Logical AND (short-circuit)</td><td><code>value && upper(value)</code></td></tr>
<tr><td><code>== !=</code></td><td>Equality (string comparison)</td><td><code>value == 'yes'</code></td></tr>
<tr><td><code>&lt; &gt; &lt;= &gt;=</code></td><td>Comparison (numeric if both parse as numbers, else string)</td><td><code>num(value) &gt; 100</code></td></tr>
<tr><td><code>+ -</code></td><td>Add/concatenate, subtract</td><td><code>'$' + value</code></td></tr>
<tr><td><code>* / %</code></td><td>Multiply, divide, modulo</td><td><code>num(value) * 100</code></td></tr>
<tr><td><code>! -</code></td><td>Logical NOT, numeric negation</td><td><code>!isEmpty(value)</code></td></tr>
<tr><td><code>.</code></td><td>Property access</td><td><code>row.last_name</code></td></tr>
</table>
<p><strong><code>+</code></strong>: concatenation if either operand is a string, addition if both are numbers. <strong><code>||</code></strong>: falsy values are <code>null</code>, <code>false</code>, <code>''</code>, <code>0</code>.</p>

<h4>String Functions</h4>
<table>
<tr><th>Function</th><th>Description</th><th>Example</th></tr>
<tr><td><code>upper(s)</code></td><td>Uppercase</td><td><code>upper('hello')</code> &rarr; <code>'HELLO'</code></td></tr>
<tr><td><code>lower(s)</code></td><td>Lowercase</td><td><code>lower('Hello')</code> &rarr; <code>'hello'</code></td></tr>
<tr><td><code>trim(s)</code></td><td>Strip whitespace</td><td><code>trim(' hi ')</code> &rarr; <code>'hi'</code></td></tr>
<tr><td><code>len(s)</code></td><td>String length</td><td><code>len('abc')</code> &rarr; <code>3</code></td></tr>
<tr><td><code>substr(s, start)</code></td><td>Substring from start</td><td><code>substr('hello', 2)</code> &rarr; <code>'llo'</code></td></tr>
<tr><td><code>substr(s, start, len)</code></td><td>Substring with length</td><td><code>substr('hello', 1, 3)</code> &rarr; <code>'ell'</code></td></tr>
<tr><td><code>replace(s, search, repl)</code></td><td>Replace first occurrence</td><td><code>replace('a-b-c', '-', '/')</code> &rarr; <code>'a/b-c'</code></td></tr>
<tr><td><code>replaceAll(s, search, repl)</code></td><td>Replace all occurrences</td><td><code>replaceAll('a-b-c', '-', '/')</code> &rarr; <code>'a/b/c'</code></td></tr>
<tr><td><code>startsWith(s, prefix)</code></td><td>Test prefix</td><td><code>startsWith('hello', 'he')</code> &rarr; <code>true</code></td></tr>
<tr><td><code>endsWith(s, suffix)</code></td><td>Test suffix</td><td><code>endsWith('hello', 'lo')</code> &rarr; <code>true</code></td></tr>
<tr><td><code>contains(s, sub)</code></td><td>Test substring</td><td><code>contains('hello', 'ell')</code> &rarr; <code>true</code></td></tr>
<tr><td><code>padLeft(s, width, char)</code></td><td>Left-pad to width</td><td><code>padLeft('42', 5, '0')</code> &rarr; <code>'00042'</code></td></tr>
<tr><td><code>padRight(s, width, char)</code></td><td>Right-pad to width</td><td><code>padRight('hi', 5, '.')</code> &rarr; <code>'hi...'</code></td></tr>
<tr><td><code>concat(s1, s2, ...)</code></td><td>Concatenate</td><td><code>concat('a', 'b', 'c')</code> &rarr; <code>'abc'</code></td></tr>
<tr><td><code>repeat(s, n)</code></td><td>Repeat string</td><td><code>repeat('*', 3)</code> &rarr; <code>'***'</code></td></tr>
</table>

<h4>Number Functions</h4>
<table>
<tr><th>Function</th><th>Description</th><th>Example</th></tr>
<tr><td><code>num(s)</code></td><td>Parse to number (null if NaN)</td><td><code>num('42.5')</code> &rarr; <code>42.5</code></td></tr>
<tr><td><code>fixed(n, decimals)</code></td><td>Format with fixed decimals</td><td><code>fixed(3.1, 2)</code> &rarr; <code>'3.10'</code></td></tr>
<tr><td><code>round(n)</code> / <code>round(n, d)</code></td><td>Round</td><td><code>round(3.456, 2)</code> &rarr; <code>3.46</code></td></tr>
<tr><td><code>floor(n)</code></td><td>Round down</td><td><code>floor(3.9)</code> &rarr; <code>3</code></td></tr>
<tr><td><code>ceil(n)</code></td><td>Round up</td><td><code>ceil(3.1)</code> &rarr; <code>4</code></td></tr>
<tr><td><code>abs(n)</code></td><td>Absolute value</td><td><code>abs(-5)</code> &rarr; <code>5</code></td></tr>
<tr><td><code>min(a, b)</code></td><td>Minimum</td><td><code>min(3, 7)</code> &rarr; <code>3</code></td></tr>
<tr><td><code>max(a, b)</code></td><td>Maximum</td><td><code>max(3, 7)</code> &rarr; <code>7</code></td></tr>
<tr><td><code>commas(n)</code></td><td>Format with thousand separators</td><td><code>commas(1234567)</code> &rarr; <code>'1,234,567'</code></td></tr>
</table>

<h4>Date Functions</h4>
<table>
<tr><th>Function</th><th>Description</th><th>Example</th></tr>
<tr><td><code>date(s, format)</code></td><td>Parse &amp; format date</td><td><code>date('2024-01-15', 'locale')</code></td></tr>
</table>
<p>Format strings: <code>'locale'</code>, <code>'iso'</code>, <code>'time'</code>, <code>'datetime'</code>, <code>'full'</code>, or a pattern with <code>YYYY</code>, <code>MM</code>, <code>DD</code> (e.g. <code>'YYYY/MM/DD'</code>). Returns the original value if unparseable.</p>
<p><code>'full'</code> shows local date and time with sub-second precision: <code>2026-06-07 13:07:06.123456789</code>. Decimal-seconds timestamps (e.g. <code>1780862826.123456789</code>) are auto-detected and the fractional part is preserved at full precision (up to 9 digits). Integer timestamps (Unix epoch seconds) are also supported.</p>

<h4>Logic / Utility Functions</h4>
<table>
<tr><th>Function</th><th>Description</th><th>Example</th></tr>
<tr><td><code>if(cond, then, else)</code></td><td>Conditional</td><td><code>if(value == '1', 'Yes', 'No')</code></td></tr>
<tr><td><code>choose(val, k1, v1, ...)</code></td><td>Map value via key-value pairs</td><td><code>choose(value, 'A', 'Active', 'I', 'Inactive')</code></td></tr>
<tr><td><code>coalesce(a, b, ...)</code></td><td>First non-empty argument</td><td><code>coalesce(row.nickname, row.name, 'Anon')</code></td></tr>
<tr><td><code>isEmpty(s)</code></td><td>True if null or <code>''</code></td><td><code>isEmpty(value) ? 'N/A' : value</code></td></tr>
<tr><td><code>isNum(s)</code></td><td>True if parseable as number</td><td><code>isNum(value) ? fixed(num(value), 2) : value</code></td></tr>
</table>

<h4>choose() Details</h4>
<p><code>choose(val, key1, result1, key2, result2, ..., default?)</code> &mdash; compares val against each key (string comparison). Returns the matching result. If no match: returns the trailing default if argument count is even (val + odd args), otherwise returns val unchanged.</p>
<pre>choose(value, 'M', 'Male', 'F', 'Female', 'Other')
  'M' &rarr; 'Male', 'F' &rarr; 'Female', 'X' &rarr; 'Other' (default)

choose(value, 'A', 'Active', 'I', 'Inactive')
  'A' &rarr; 'Active', 'I' &rarr; 'Inactive', 'X' &rarr; 'X' (no default)</pre>

<h4>Error Handling</h4>
<ul>
<li><strong>Parse errors</strong> are caught when a plugin is loaded. Invalid column rules are skipped.</li>
<li><strong>Runtime errors</strong> fall back to displaying the raw value silently.</li>
<li><code>num()</code> on non-numeric strings returns <code>null</code>.</li>
<li>Division by zero returns <code>null</code> (displays as empty).</li>
<li>Property access on null returns <code>null</code>.</li>
</ul>
    `);
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
    cancelQuery,
    clearConsole,
    runConsole,
    layoutTileH,
    layoutTileV,
    layoutGrid,
    layoutCascade,
    minimizeAll,
    restoreAll,
    showAbout,
    showManual,
    loadPluginFromFile,
    showExpressionReference,
    ...(new URLSearchParams(location.search).has('test') ? {
      _test: {
        sanitizeTableName, sanitizeColumnName, sanitizeColumns,
        getUniqueTableName, extractIntoClause,
        get tables() { return tables; },
        get windows() { return windows; },
        get db() { return db; },
        set _shiftOpen(v) { _shiftOpen = v; },
        get plugins() { return plugins; },
        get _columnTransformCache() { return _columnTransformCache; },
        exprCompile, exprEval, exprEvalToString,
        validatePlugin, compilePlugin, loadPersistedPlugins,
        rebuildTransformCache, rebuildTransformCacheForTable,
        getDisplayValue, hasDisplayTransform,
        loadPluginFile, unloadPlugin, persistPlugins, updatePluginMenu,
        showToast, showPluginAbout, closePluginPopover,
        rebuildTable, rerenderAllWindows,
        undoTable, redoTable,
      }
    } : {}),
  };
})();

document.addEventListener('DOMContentLoaded', app.init);
