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

  // Virtual scrolling constants
  const ROW_HEIGHT = 26;
  const OVERSCAN = 10;

  // Debounced sync timers
  const syncTimers = {};

  // Sort optimization
  const collator = new Intl.Collator(undefined, { sensitivity: 'base' });

  // Zip group counter — tables from the same zip share a zipGroupId
  let nextZipGroupId = 1;

  // Excel group counter — tables from the same workbook share an excelGroupId
  let nextExcelGroupId = 1;

  // ---- Init ----
  async function init() {
    const SQL = await initSqlJs({
      locateFile: file => `https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.10.3/${file}`
    });
    db = new SQL.Database();
    db.create_function('regexp', (pattern, value) => {
      try { return new RegExp(pattern, 'i').test(value) ? 1 : 0; } catch (_) { return 0; }
    });
    setupConsoleResize();
    setupFileInput();
    setupDragAndDrop();
    setupKeyboard();
    setupMenuClose();
    setupAI();
    setupBrowserResize();
    fixShortcutLabels();
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

  function openFileByType(file, fileHandle, compression) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (COMPRESSION_EXTS.has(ext)) {
      decompressAndOpen(file, fileHandle);
    } else if (ext === 'xlsx' || ext === 'xls') {
      loadExcelFile(file, fileHandle, compression);
    } else {
      loadDelimitedFile(file, compression ? compression.fileHandle : fileHandle, compression);
    }
  }

  async function decompressAndOpen(file, fileHandle) {
    const ext = file.name.split('.').pop().toLowerCase();
    setStatus(`Decompressing ${file.name}...`, 'working');
    try {
      if (ext === 'gz') {
        await decompressGzip(file, fileHandle);
      } else if (ext === 'zip') {
        await decompressZip(file, fileHandle);
      } else {
        setStatus(`Unsupported compression format: .${ext} — please decompress the file first and open the decompressed file`, 'error');
      }
    } catch (e) {
      setStatus(`Error decompressing ${file.name}: ${e.message}`, 'error');
    }
  }

  async function decompressGzip(file, fileHandle) {
    const ds = new DecompressionStream('gzip');
    const decompressed = file.stream().pipeThrough(ds);
    const blob = await new Response(decompressed).blob();
    // Inner filename: strip .gz
    const innerName = file.name.replace(/\.gz$/i, '') || 'decompressed.csv';
    const innerFile = new File([blob], innerName, { type: 'application/octet-stream' });
    openFileByType(innerFile, null, { type: 'gz', compressedFilename: file.name, fileHandle: fileHandle || null });
  }

  async function decompressZip(file, fileHandle) {
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
      openFileByType(innerFile, null, { type: 'zip', compressedFilename: file.name, fileHandle: fileHandle || null, zipGroupId, innerName: name, zipOriginalCount });
    }
  }

  function delimiterForExt(filename) {
    const ext = filename.split('.').pop().toLowerCase();
    if (ext === 'tsv') return '\t';
    if (ext === 'psv') return '|';
    return undefined; // let Papa auto-detect (handles csv, txt)
  }

  function loadDelimitedFile(file, fileHandle, compression) {
    setStatus(`Loading ${file.name}...`, 'working');
    const t0 = performance.now();
    const delimiter = delimiterForExt(file.name);
    const allRows = [];
    let detectedDelimiter = delimiter || ',';
    let rawColumns = null;
    let columns = null;
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      dynamicTyping: false,
      delimiter,
      chunk(results) {
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

  function loadExcelFile(file, fileHandle, compression) {
    setStatus(`Loading ${file.name}...`, 'working');
    const t0 = performance.now();
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const workbook = XLSX.read(e.target.result, { type: 'array' });
        const excelGroupId = nextExcelGroupId++;
        const nonEmptySheets = workbook.SheetNames.filter(sn => {
          const s = workbook.Sheets[sn];
          return XLSX.utils.sheet_to_json(s, { defval: '' }).length > 0;
        });
        const excelOriginalCount = nonEmptySheets.length;
        for (const sheetName of nonEmptySheets) {
          const sheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: '' });
          const name = sanitizeTableName(sheetName);
          const uniqueName = getUniqueTableName(name);
          const rawColumns = Object.keys(jsonData[0]);
          const columns = sanitizeColumns(rawColumns);
          const rows = jsonData.map((row, i) => {
            const r = { _rownum: i + 1 };
            rawColumns.forEach((raw, j) => { r[columns[j]] = row[raw] != null ? String(row[raw]) : ''; });
            return r;
          });
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
    if (win.tableName && tables[win.tableName]) {
      const t = tables[win.tableName];
      if (!win.isQuery && t.modified) {
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
      const dx = e.clientX - startX, dy = e.clientY - startY;
      if (Math.abs(dx) < 3 && Math.abs(dy) < 3) return;
      win.el.style.left = (origX + dx) + 'px';
      win.el.style.top = (origY + dy) + 'px';
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
      <input type="text" class="filter-input" placeholder="WHERE clause, e.g. age > 30 AND name LIKE '%Smith%'" value="${escHtml(win.filterText)}">
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

    columns.forEach((col, colIdx) => {
      const th = document.createElement('th');
      th.textContent = col;
      th.dataset.colIdx = colIdx;
      const sortIdx = win.sortCols.findIndex(s => s.col === col);
      if (sortIdx !== -1) {
        th.classList.add('sorted');
        const arrow = document.createElement('span');
        arrow.className = 'sort-arrow';
        const dir = win.sortCols[sortIdx].dir;
        arrow.textContent = dir === 'asc' ? '\u25B2' : '\u25BC';
        if (win.sortCols.length > 1) arrow.textContent += (sortIdx + 1);
        th.appendChild(arrow);
      }
      // Ctrl/Cmd+drag to reorder columns, Ctrl/Cmd+click to rename
      th.addEventListener('mousedown', (e) => {
        if (!(e.ctrlKey || e.metaKey) || e.button !== 0) return;
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
                if (ue.clientX >= mid && dropIdx < columns.length - 1) dropIdx++;
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
      th.addEventListener('click', (e) => {
        if (th._renaming) return;
        if (e.ctrlKey || e.metaKey) {
          e.stopPropagation();
          if (th._didDrag) { th._didDrag = false; return; }
          startColumnRename(win, th, col);
          return;
        }
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

    const savedScrollLeft = container.scrollLeft;
    buildTableHTML(win, container, t);

    container.scrollLeft = savedScrollLeft;
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

  async function addColumn(tableName, colName) {
    const t = tables[tableName];
    if (t.columns.includes(colName)) return;
    t.columns.push(colName);
    t.rows.forEach(r => { r[colName] = ''; });
    markModified(tableName);
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
    markModified(tableName);
    try { db.run(`ALTER TABLE [${tableName}] RENAME COLUMN [${oldCol}] TO [${newCol}]`); } catch (_) {}
    rebuildTable(win);
  }

  async function reorderColumn(win, fromIdx, toIdx) {
    const t = tables[win.tableName];
    if (!t) return;
    const col = t.columns.splice(fromIdx, 1)[0];
    if (toIdx > fromIdx) toIdx--;
    t.columns.splice(toIdx, 0, col);
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
    th.textContent = '';
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
      if (!(e.ctrlKey || e.metaKey)) return;
      switch (e.key) {
        case 's': e.preventDefault(); saveActiveTable(); break;
        case 'o': e.preventDefault(); openFile(); break;
        case 'n': e.preventDefault(); newTable(); break;
        case 'w': e.preventDefault(); if (activeWinId) closeWindow(activeWinId); break;
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

  // Worker-based query execution for interruptibility and live timer
  let _queryWorker = null;
  let _queryTimer = null;

  function makeQueryWorker() {
    const src = `
      importScripts('https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.10.3/sql-wasm.js');
      let SQL = null;
      let db = null;
      initSqlJs({ locateFile: f => 'https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.10.3/' + f }).then((sql) => {
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

  function showInterruptButton(show) {
    let btn = document.getElementById('btn-interrupt');
    if (!btn) {
      btn = document.createElement('button');
      btn.id = 'btn-interrupt';
      btn.textContent = 'Interrupt';
      btn.onclick = cancelQuery;
      document.getElementById('console-actions').appendChild(btn);
    }
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
      showQueryResult(sql, rows);
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
    tables[tableName] = { columns, rows, filename: null, modified: true };

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
    } else {
      document.getElementById('sql-input').value = '';
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
    }

    function openItem(item) {
      allItems.forEach(m => m.classList.remove('open'));
      item.classList.add('open');
      updateMenuState();
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
      <p>Version 0.8.1 &mdash; &copy; 2026 Mark Kim</p>
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
<p>Use <strong>File &rarr; Open</strong> (<code>Ctrl+O</code> / <code>&#8984;O</code>), <strong>File &rarr; Open URL</strong>, or drag and drop files onto the window.</p>

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
<li><strong>Edit cells:</strong> Click any cell to edit. Press <code>Tab</code>/<code>Shift+Tab</code> to move between cells, <code>Enter</code> to move down, <code>Escape</code> to cancel.</li>
<li><strong>Add rows:</strong> Click <code>+ Row</code> in the toolbar, or right-click a row number to insert above.</li>
<li><strong>Delete rows:</strong> Right-click a row number and choose Delete Row.</li>
<li><strong>Add columns:</strong> Click <code>+ Col</code> in the toolbar.</li>
<li><strong>Rename columns:</strong> <code>Ctrl</code>/<code>&#8984;</code>+click a column header.</li>
<li><strong>Reorder columns:</strong> <code>Ctrl</code>/<code>&#8984;</code>+drag a column header to a new position.</li>
<li><strong>Rename tables:</strong> <code>Ctrl</code>/<code>&#8984;</code>+click the window title.</li>
</ul>

<h4>Sorting &amp; Filtering</h4>
<ul>
<li><strong>Sort:</strong> Click a column header to sort ascending &rarr; descending &rarr; unsorted.</li>
<li><strong>Multi-column sort:</strong> Shift+click additional column headers. Numbers next to arrows indicate sort priority.</li>
<li><strong>Filter:</strong> Type a SQL <code>WHERE</code> clause in the filter bar (without the <code>WHERE</code> keyword). For example: <code>age > 30 AND name LIKE '%Smith%'</code></li>
<li>The filter supports all SQLite expressions including <code>REGEXP</code> (see below).</li>
</ul>

<h4>SQL Console</h4>
<p>The SQL Console at the bottom runs queries against all open tables using SQLite syntax. Press <code>Ctrl+Enter</code> / <code>&#8984;+Enter</code> to execute.</p>
<p>Tables are referenced by the name shown in their window title bar. Names are sanitized to <code>[a-zA-Z0-9_]</code> characters.</p>

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
<li><strong>Move:</strong> Drag the title bar.</li>
<li><strong>Resize:</strong> Drag any edge or corner.</li>
<li><strong>Maximize/Restore:</strong> Double-click the title bar, or click the maximize button.</li>
<li><strong>Minimize:</strong> Click the minimize button. Restore from the Windows menu.</li>
<li><strong>Close:</strong> Click the close button. <code>Ctrl</code>/<code>&#8984;</code>+click closes all windows.</li>
<li><strong>Layout:</strong> Use the View menu to tile, grid, or cascade all windows.</li>
<li><strong>Proportional scaling:</strong> Windows reposition and resize proportionally when the browser window or console panel is resized.</li>
</ul>

<h4>Keyboard Shortcuts</h4>
<table>
<tr><th>Shortcut</th><th>Action</th></tr>
<tr><td><code>Ctrl+O</code> / <code>&#8984;O</code></td><td>Open file</td></tr>
<tr><td><code>Ctrl+S</code> / <code>&#8984;S</code></td><td>Save table</td></tr>
<tr><td><code>Ctrl+N</code> / <code>&#8984;N</code></td><td>New table</td></tr>
<tr><td><code>Ctrl+W</code> / <code>&#8984;W</code></td><td>Close window</td></tr>
<tr><td><code>Ctrl+Enter</code></td><td>Execute SQL query</td></tr>
<tr><td><code>Enter</code></td><td>Send AI prompt</td></tr>
<tr><td><code>Shift+Enter</code></td><td>Newline in AI prompt</td></tr>
<tr><td><code>Up</code> / <code>Down</code></td><td>AI prompt history</td></tr>
</table>

<h4>AI Analysis <em>(experimental)</em></h4>
<p>The AI tab in the console panel lets you analyze your data using natural language. The AI has full SQL access to your data &mdash; it writes and executes queries automatically to answer your questions with exact results, regardless of dataset size.</p>
<p><strong>Four provider options:</strong></p>
<ul>
<li><strong>WebLLM (default):</strong> Runs entirely in the browser via WebGPU. Requires Chrome/Edge 113+. No install, no API key, no data leaves your machine.</li>
<li><strong>Ollama:</strong> Local AI server. Install from <a href="https://ollama.com">ollama.com</a>, then run <code>ollama pull llama3.2</code>. Larger models than WebLLM, still fully local.</li>
<li><strong>Claude (Anthropic):</strong> Cloud provider. Requires an API key from <a href="https://console.anthropic.com">console.anthropic.com</a>. Best reasoning quality.</li>
<li><strong>OpenAI:</strong> Cloud provider. Requires an API key from <a href="https://platform.openai.com">platform.openai.com</a>.</li>
</ul>
<p><strong>Usage:</strong> Switch to the AI tab, select one or more tables, type your question, and press <code>Enter</code> or click Run. Use <code>Shift+Enter</code> for multiline prompts. Press <code>Up</code>/<code>Down</code> arrow to recall previous prompts.</p>
<p><strong>How it works:</strong> The AI receives column statistics and sample rows for context, then writes SQL queries in <code>\`\`\`sql</code> code blocks. These queries are executed automatically against the full dataset, and the results are fed back to the AI for analysis. This loop repeats (up to 5 rounds) until the AI has enough data to answer.</p>
<p>Click the gear icon &#9881; to configure the provider, model, and API keys.</p>
    `);
  }

  // ---- AI Analysis ----
  let _aiProvider = null;
  let _aiAbort = null;
  let _webllmEngine = null;
  let _aiConversation = []; // accumulated user/assistant message history

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
    if (_aiProvider === 'webllm') return 6000; // 4K token context; leave room for system prompt, conversation, and response
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
            // 10-bucket histogram
            if (mn != null && mx != null) {
              const minVal = Number(mn), maxVal = Number(mx);
              if (isFinite(minVal) && isFinite(maxVal) && maxVal > minVal) {
                const bucketWidth = (maxVal - minVal) / 10;
                const histRes = db.exec(`SELECT CASE WHEN CAST(((CAST([${col}] AS REAL) - ${minVal}) / ${bucketWidth}) AS INTEGER) >= 10 THEN 9 ELSE CAST(((CAST([${col}] AS REAL) - ${minVal}) / ${bucketWidth}) AS INTEGER) END AS bucket, COUNT(*) FROM [${tableName}] WHERE [${col}] GLOB '*[0-9]*' GROUP BY bucket ORDER BY bucket`);
                if (histRes.length) {
                  const buckets = histRes[0].values.map(r => {
                    const b = r[0];
                    const lo = (minVal + b * bucketWidth).toFixed(2);
                    const hi = (minVal + (b + 1) * bucketWidth).toFixed(2);
                    return `[${lo}-${hi}]: ${r[1]}`;
                  });
                  line += `\n    Distribution: ${buckets.join(' | ')}`;
                }
              }
            }
          } catch {}
        } else if (distinct <= 1000 && distinct > 0) {
          // Categorical: all value counts (≤1000 distinct fits easily in budget)
          try {
            const topRes = db.exec(`SELECT [${col}], COUNT(*) as cnt FROM [${tableName}] WHERE [${col}] IS NOT NULL AND [${col}] != '' GROUP BY [${col}] ORDER BY cnt DESC`);
            if (topRes.length) {
              const vals = topRes[0].values.map(r => `${r[0]} (${r[1]})`).join(', ');
              line += `\n    Value counts: ${vals}`;
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
    const budgetPerTable = Math.floor(getDataCharBudget() / tableNames.length);

    for (const name of tableNames) {
      const t = tables[name];
      if (!t || t.columns.length === 0) continue;

      let info = `Table: ${name}\nColumns: ${t.columns.join(', ')}\nTotal rows: ${t.rows.length}\n\n`;

      // For small tables (<=200 rows), include all rows directly
      if (t.rows.length <= 200) {
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

  // Format AI response text with basic markdown
  function formatAIResponse(text) {
    let html = escHtml(text);
    // Code blocks: ```...```
    html = html.replace(/```(\w*)\n?([\s\S]*?)```/g, (_, lang, code) => '<pre>' + code.trim() + '</pre>');
    // Inline code
    html = html.replace(/`([^`]+)`/g, '<code>$1</code>');
    // Bold
    html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
    // Newlines (but not inside pre tags)
    html = html.replace(/\n/g, '<br>');
    // Clean up <br> inside <pre>
    html = html.replace(/<pre>([\s\S]*?)<\/pre>/g, (_, code) => '<pre>' + code.replace(/<br>/g, '\n') + '</pre>');
    return html;
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
    if (selectedTables.length === 0) {
      setAIStatus('Open a CSV file first to analyze data.', 'error');
      return;
    }
    if (!prompt) {
      setAIStatus('Enter a prompt to analyze the data.', 'error');
      return;
    }
    if (!_aiProvider) {
      setAIStatus('No AI provider available. Install Ollama or use Chrome with WebGPU.', 'error');
      showSetupHelp();
      return;
    }

    flushAllSyncs();
    const dataContext = buildDataContext(selectedTables);
    _aiConversation.push({ role: 'user', content: prompt });

    const tableList = selectedTables.map(n => {
      const t = tables[n];
      return `[${n}] (${t ? t.columns.join(', ') : 'unknown columns'})`;
    }).join(', ');
    const systemPrompt = `You are a data analyst. You have full access to a SQLite database containing the user's data.

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

${dataContext}`;

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
    _aiAbort = new AbortController();
    const signal = _aiAbort.signal;

    // Elapsed timer
    const aiTimer = setInterval(() => {
      const elapsed = ((performance.now() - t0) / 1000).toFixed(1);
      setAIStatus(`Generating response... ${elapsed}s`, 'working');
    }, 100);

    const MAX_SQL_ROUNDS = 5;

    try {
      for (let round = 0; round <= MAX_SQL_ROUNDS; round++) {
        const messages = [
          { role: 'system', content: systemPrompt },
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
                // Limit to 200 result rows to stay within budget
                const shown = rows.slice(0, 200);
                resultsText += `Query: ${sql}\n${header}\n${shown.join('\n')}`;
                if (rows.length > 200) resultsText += `\n... (${rows.length - 200} more rows)`;
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
      const elapsed = performance.now() - t0;
      setAIStatus(`Done in ${formatElapsed(elapsed)}`, 'success');
    } catch (e) {
      clearInterval(aiTimer);
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
    _aiAbort = null;
  }

  // Ollama streaming
  async function generateOllama(messages, onChunk, signal) {
    const r = await fetch(aiSettings.ollamaUrl + '/api/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ model: aiSettings.model, messages, stream: true }),
      signal,
    });
    if (!r.ok) throw new Error(`Ollama error: ${r.status} ${r.statusText}`);
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
  }

  // WebLLM generation
  async function generateWebLLM(messages, onChunk, signal) {
    if (!_webllmEngine) {
      setAIStatus('Loading WebLLM engine (first time may download model)...', 'working');
      const webllm = await import('https://esm.run/@mlc-ai/web-llm');
      const model = aiSettings.model || 'Qwen3-8B-q4f16_1-MLC';
      _webllmEngine = await webllm.CreateMLCEngine(model, {
        initProgressCallback: (progress) => {
          const pct = progress.progress != null ? Math.round(progress.progress * 100) + '%' : '';
          setAIStatus(`Loading model: ${progress.text || pct}`, 'working');
        },
      }, {
        context_window_size: 4096,
        sliding_window_size: -1,
      });
    }
    const response = await _webllmEngine.chat.completions.create({
      messages,
      stream: true,
    });
    for await (const chunk of response) {
      if (signal.aborted) {
        if (_webllmEngine.interruptGenerate) _webllmEngine.interruptGenerate();
        throw new DOMException('Aborted', 'AbortError');
      }
      const delta = chunk.choices[0] && chunk.choices[0].delta && chunk.choices[0].delta.content;
      if (delta) onChunk(delta);
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

  function setupAI() {
    setupConsoleTabs();
    document.getElementById('ai-settings-btn').addEventListener('click', showAISettings);
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
    ...(new URLSearchParams(location.search).has('test') ? {
      _test: {
        sanitizeTableName, sanitizeColumnName, sanitizeColumns,
        getUniqueTableName, extractIntoClause,
        get tables() { return tables; },
        get windows() { return windows; },
        get db() { return db; },
      }
    } : {}),
  };
})();

document.addEventListener('DOMContentLoaded', app.init);
