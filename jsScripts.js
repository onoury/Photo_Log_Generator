/* Photo Log Builder v1.4
 * Local static browser app. Original photos remain browser File references.
 * Pyodide/Pillow processes one source photo at a time. Processed JPEGs can be
 * previewed in a photo-log layout, then Python packages the Word template into
 * one finished DOCX that the browser downloads normally.
 */
'use strict';
(() => {
  const $ = id => document.getElementById(id);
  const PYODIDE = 'https://cdn.jsdelivr.net/pyodide/v0.29.3/full/';
  const TEMPLATE_FILENAME = 'Word Photo Log Organs - Clean.zip';
  const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  const encoder = new TextEncoder();
  const decoder = new TextDecoder('utf-8');
  const PHOTO_EXTENSIONS = new Set(['jpg','jpeg','png','webp','bmp','tif','tiff']);
  const PREVIEW_ROWS_PER_PAGE = 3;
  const PREVIEW_SMALL_MAX_W = 480;
  const PREVIEW_SMALL_MAX_H = 360;

  const state = {
    ready: false,
    options: null,
    worker: null,
    workerURL: null,
    nextId: 0,
    pending: new Map(),
    photos: [],
    photoIndex: new Map(),
    rows: [],
    templateBuffer: null,
    templateName: TEMPLATE_FILENAME,
    templateReady: false,
    working: false,
    cancelRequested: false,
    validation: null,
    result: null,
    processCache: null,
    previewVisible: false,
  };

  // ---------- General UI helpers ----------
  const fmtInt = n => Number(n || 0).toLocaleString();
  function fmtBytes(bytes) {
    const n = Number(bytes || 0);
    if (!Number.isFinite(n) || n <= 0) return '0 B';
    const units = ['B','KB','MB','GB','TB'];
    const i = Math.min(Math.floor(Math.log(n) / Math.log(1024)), units.length - 1);
    const value = n / (1024 ** i);
    return `${value.toFixed(i < 2 ? 0 : value >= 100 ? 0 : value >= 10 ? 1 : 2)} ${units[i]}`;
  }
  function esc(value) {
    return String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  }
  function message(text, kind='error') {
    const el = $('message');
    el.textContent = text || '';
    el.hidden = !text;
    el.className = `message${kind === 'success' ? ' success' : kind === 'info' ? ' info' : ''}`;
  }
  function setRuntime(text, error=false) {
    const el = $('runtime-status');
    el.textContent = text;
    el.className = `status-pill${error ? ' error' : ''}`;
  }
  function setTemplateStatus(text, error=false) {
    const el = $('template-status');
    el.textContent = text;
    el.className = `status-pill neutral${error ? ' error' : ''}`;
  }
  function sanitizeOutputName(value) {
    const cleaned = String(value || 'Photo_Log').trim().replace(/[<>:"/\\|?*\x00-\x1f]/g, '_').replace(/\.docx$/i, '');
    return cleaned || 'Photo_Log';
  }
  function photoStem(name) {
    return String(name || '').replace(/\.[^.]+$/, '').replace(/[_-]+/g, ' ').trim();
  }
  function isLikelyImage(file) {
    const ext = String(file?.name || '').split('.').pop().toLocaleLowerCase();
    return PHOTO_EXTENSIONS.has(ext);
  }
  function sentenceCase(text, fallback='') {
    const value = String(text || '').trim() || String(fallback || '').trim();
    if (!value) return '';
    return value.slice(0, 1).toUpperCase() + value.slice(1).toLowerCase();
  }

  // ---------- Python worker ----------
  function invoke(payload, transfer=[]) {
    if (!state.worker) return Promise.reject(new Error('Python worker is not available.'));
    const id = ++state.nextId;
    return new Promise((resolve, reject) => {
      state.pending.set(id, {resolve, reject});
      state.worker.postMessage({id, payload}, transfer);
    });
  }
  function rejectPending(text) {
    for (const task of state.pending.values()) task.reject(new Error(text));
    state.pending.clear();
  }
  function friendlyError(raw) {
    const text = String(raw || 'Unknown error');
    const lines = text.trim().split('\n');
    const useful = [...lines].reverse().find(line => /^(ValueError|RuntimeError|FileNotFoundError|UnidentifiedImageError|OSError|Error|TypeError):/.test(line.trim()));
    return (useful || text).replace(/^[A-Za-z]+Error:\s*/, '');
  }

  async function bootPython() {
    state.ready = false;
    updateActions();
    setRuntime('Loading Python…');
    try {
      if (location.protocol === 'file:') {
        throw new Error('Open this app through a local web server. Run “python -m http.server 8000” in the app folder, then open http://localhost:8000.');
      }
      if (state.worker) state.worker.terminate();
      rejectPending('Python runtime restarted.');
      if (state.workerURL) URL.revokeObjectURL(state.workerURL);

      const sourceResponse = await fetch(new URL('pyScripts.py', document.baseURI));
      if (!sourceResponse.ok) throw new Error('pyScripts.py could not be loaded. Keep it beside index.html.');
      const source = await sourceResponse.text();

      const workerCode = `
        let py;
        let queue = Promise.resolve();
        function clean(path){ try { if (py?.FS?.analyzePath(path).exists) py.FS.unlink(path); } catch(_){} }
        self.onmessage = event => {
          queue = queue.then(async () => {
            const {id, payload} = event.data;
            try {
              if (payload.action === 'init') {
                self.postMessage({status:'Loading Python…'});
                importScripts(${JSON.stringify(PYODIDE + 'pyodide.js')});
                py = await loadPyodide({indexURL:${JSON.stringify(PYODIDE)}});
                self.postMessage({status:'Loading Pillow…'});
                await py.loadPackage('pillow');
                await py.runPythonAsync(payload.source);
                py.globals.set('_browser_request', JSON.stringify({action:'options'}));
                try {
                  const result = JSON.parse(py.runPython('browser_dispatch(_browser_request)'));
                  self.postMessage({id, result});
                } finally { py.globals.delete('_browser_request'); }
                return;
              }

              if (payload.action === 'process_image') {
                const safeExt = String(payload.filename || '').match(/\.[A-Za-z0-9]{1,8}$/)?.[0] || '.img';
                const inputPath = '/tmp/photo_log_input' + safeExt;
                const outputPath = '/tmp/photo_log_stage/' + payload.image_name;
                clean(inputPath);
                try {
                  py.FS.mkdirTree('/tmp/photo_log_stage');
                  py.FS.writeFile(inputPath, new Uint8Array(payload.buffer));
                  py.globals.set('_browser_input_path', inputPath);
                  py.globals.set('_browser_output_path', outputPath);
                  py.globals.set('_browser_settings_json', JSON.stringify(payload.settings || {}));
                  const result = JSON.parse(py.runPython('browser_process_image(_browser_input_path, _browser_output_path, _browser_settings_json)'));
                  if (payload.return_bytes) {
                    const wasmBytes = py.FS.readFile(outputPath);
                    const copy = new Uint8Array(wasmBytes.length);
                    copy.set(wasmBytes);
                    self.postMessage({id, result, buffer:copy.buffer}, [copy.buffer]);
                  } else {
                    self.postMessage({id, result});
                  }
                } finally {
                  try { py.globals.delete('_browser_input_path'); } catch(_){}
                  try { py.globals.delete('_browser_output_path'); } catch(_){}
                  try { py.globals.delete('_browser_settings_json'); } catch(_){}
                  clean(inputPath);
                }
                return;
              }

              if (payload.action === 'build_docx') {
                const templatePath = '/tmp/photo_log_template.zip';
                const outputPath = '/tmp/photo_log.docx';
                clean(templatePath); clean(outputPath);
                try {
                  py.FS.writeFile(templatePath, new Uint8Array(payload.template_buffer));
                  py.globals.set('_browser_template_path', templatePath);
                  py.globals.set('_browser_output_path', outputPath);
                  py.globals.set('_browser_entries_json', JSON.stringify(payload.entries || []));
                  py.globals.set('_browser_settings_json', JSON.stringify(payload.settings || {}));
                  const result = JSON.parse(py.runPython('browser_build_docx(_browser_template_path, _browser_output_path, _browser_entries_json, _browser_settings_json)'));
                  const wasmBytes = py.FS.readFile(outputPath);
                  const copy = new Uint8Array(wasmBytes.length);
                  copy.set(wasmBytes);
                  self.postMessage({id, result, buffer:copy.buffer}, [copy.buffer]);
                } finally {
                  try { py.globals.delete('_browser_template_path'); } catch(_){}
                  try { py.globals.delete('_browser_output_path'); } catch(_){}
                  try { py.globals.delete('_browser_entries_json'); } catch(_){}
                  try { py.globals.delete('_browser_settings_json'); } catch(_){}
                  clean(templatePath); clean(outputPath);
                }
                return;
              }

              if (payload.action === 'clear_stage') {
                const result = JSON.parse(py.runPython('browser_clear_stage()'));
                self.postMessage({id, result});
                return;
              }

              py.globals.set('_browser_request', JSON.stringify(payload));
              try {
                const result = JSON.parse(py.runPython('browser_dispatch(_browser_request)'));
                self.postMessage({id, result});
              } finally { py.globals.delete('_browser_request'); }
            } catch (error) {
              self.postMessage({id, error:String(error?.message || error)});
            }
          });
        };
      `;

      state.workerURL = URL.createObjectURL(new Blob([workerCode], {type:'text/javascript'}));
      state.worker = new Worker(state.workerURL);
      state.worker.onmessage = event => {
        if (event.data.status) { setRuntime(event.data.status); return; }
        const job = state.pending.get(event.data.id);
        if (!job) return;
        state.pending.delete(event.data.id);
        if (event.data.error) job.reject(new Error(friendlyError(event.data.error)));
        else if (event.data.buffer) job.resolve({result:event.data.result, buffer:event.data.buffer});
        else job.resolve(event.data.result);
      };
      state.worker.onerror = event => {
        state.ready = false;
        setRuntime('Python unavailable', true);
        rejectPending(event.message || 'Python worker failed.');
        message('The Python runtime could not start. Check your connection and reload the page.');
        updateActions();
      };

      const options = await invoke({action:'init', source});
      state.options = options;
      applyDefaults(options.defaults);
      state.ready = true;
      setRuntime('Python ready');
      updateActions();
    } catch (error) {
      state.ready = false;
      setRuntime('Python unavailable', true);
      message(error.message);
      updateActions();
    }
  }

  // ---------- Template ZIP reader (used only for streaming output) ----------
  class TemplateZipReader {
    constructor(arrayBuffer) {
      this.bytes = new Uint8Array(arrayBuffer);
      this.view = new DataView(arrayBuffer);
      this.entries = this.#parse();
    }
    #u16(offset){ return this.view.getUint16(offset, true); }
    #u32(offset){ return this.view.getUint32(offset, true); }
    #parse() {
      const b = this.bytes;
      const min = Math.max(0, b.length - 65557);
      let eocd = -1;
      for (let i = b.length - 22; i >= min; i--) {
        if (this.#u32(i) === 0x06054b50) { eocd = i; break; }
      }
      if (eocd < 0) throw new Error('The template ZIP does not contain a standard end-of-central-directory record.');
      const disk = this.#u16(eocd + 4), cdDisk = this.#u16(eocd + 6);
      const total = this.#u16(eocd + 10), cdOffset = this.#u32(eocd + 16);
      if (disk !== 0 || cdDisk !== 0 || total === 0xffff || cdOffset === 0xffffffff) {
        throw new Error('ZIP64 or multi-disk templates are not supported by streaming mode.');
      }
      const entries = [];
      let pos = cdOffset;
      for (let i = 0; i < total; i++) {
        if (this.#u32(pos) !== 0x02014b50) throw new Error('The template ZIP central directory is invalid.');
        const flags = this.#u16(pos + 8);
        const method = this.#u16(pos + 10);
        const compressedSize = this.#u32(pos + 20);
        const uncompressedSize = this.#u32(pos + 24);
        const nameLen = this.#u16(pos + 28);
        const extraLen = this.#u16(pos + 30);
        const commentLen = this.#u16(pos + 32);
        const localOffset = this.#u32(pos + 42);
        const name = decoder.decode(b.subarray(pos + 46, pos + 46 + nameLen));
        if (flags & 0x0001) throw new Error(`Encrypted template entry is not supported: ${name}`);
        entries.push({name, method, compressedSize, uncompressedSize, localOffset});
        pos += 46 + nameLen + extraLen + commentLen;
      }
      return entries;
    }
    async getData(entry) {
      const pos = entry.localOffset;
      if (this.view.getUint32(pos, true) !== 0x04034b50) throw new Error(`Invalid local ZIP header for ${entry.name}.`);
      const nameLen = this.view.getUint16(pos + 26, true);
      const extraLen = this.view.getUint16(pos + 28, true);
      const start = pos + 30 + nameLen + extraLen;
      const compressed = this.bytes.slice(start, start + entry.compressedSize);
      if (entry.method === 0) return compressed;
      if (entry.method !== 8) throw new Error(`Unsupported compression method ${entry.method} in ${entry.name}.`);
      if (typeof DecompressionStream === 'undefined') throw new Error('Streaming mode requires a current Chrome or Edge browser.');
      const stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
      const out = new Uint8Array(await new Response(stream).arrayBuffer());
      if (entry.uncompressedSize && out.length !== entry.uncompressedSize) throw new Error(`Template entry size check failed for ${entry.name}.`);
      return out;
    }
  }

  // ---------- Word template ----------
  async function acceptTemplateBuffer(buffer, displayName) {
    if (!(buffer instanceof ArrayBuffer) || buffer.byteLength < 4) throw new Error('The selected template is empty or unreadable.');
    const sig = new Uint8Array(buffer, 0, Math.min(4, buffer.byteLength));
    if (sig[0] !== 0x50 || sig[1] !== 0x4b) throw new Error('The selected template is not a ZIP/DOCX package.');
    state.templateBuffer = buffer;
    state.templateName = displayName;
    state.templateReady = true;
    $('template-name').textContent = displayName;
    $('template-name').title = displayName;
    setTemplateStatus('Template ready');
    updateActions();
  }

  async function loadBundledTemplate() {
    state.templateReady = false;
    updateActions();
    setTemplateStatus('Loading template…');
    try {
      const response = await fetch(new URL(TEMPLATE_FILENAME, document.baseURI));
      if (!response.ok) throw new Error(`Could not load ${TEMPLATE_FILENAME}.`);
      await acceptTemplateBuffer(await response.arrayBuffer(), TEMPLATE_FILENAME);
    } catch (error) {
      state.templateReady = false;
      setTemplateStatus('Template needed', true);
      message(`${error.message}\nKeep “${TEMPLATE_FILENAME}” beside index.html and reload the app.`);
      updateActions();
    }
  }

  // ---------- CSV ----------
  function parseCSV(text) {
    text = String(text || '').replace(/^\ufeff/, '');
    const rows = [];
    let row = [], field = '', quoted = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (quoted) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; }
          else quoted = false;
        } else field += c;
      } else {
        if (c === '"') quoted = true;
        else if (c === ',') { row.push(field); field = ''; }
        else if (c === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
        else if (c !== '\r') field += c;
      }
    }
    if (quoted) throw new Error('CSV contains an unterminated quoted field.');
    if (field.length || row.length) { row.push(field); rows.push(row); }
    return rows.filter(r => r.some(v => String(v).trim() !== ''));
  }
  function csvTextToRows(text) {
    const matrix = parseCSV(text);
    if (!matrix.length) throw new Error('The CSV is empty.');
    const headers = matrix[0].map(v => String(v).trim().toLowerCase());
    const fileIndex = headers.findIndex(h => ['file_name','file name','filename','file'].includes(h));
    const descIndex = headers.findIndex(h => ['photo_description','photo description','description','caption'].includes(h));
    if (fileIndex < 0 || descIndex < 0) throw new Error('CSV headers must include File_Name and Photo_Description.');
    return matrix.slice(1).map((r, index) => ({
      fileName: String(r[fileIndex] ?? '').trim(),
      description: String(r[descIndex] ?? '').trim(),
      sourceLine: index + 2,
    })).filter(r => r.fileName || r.description);
  }
  function quoteCSV(value) { return `"${String(value ?? '').replace(/"/g, '""')}"`; }
  function saveCSV() {
    const text = ['Photo_Description,File_Name', ...state.rows.map(r => `${quoteCSV(r.description)},${quoteCSV(r.fileName)}`)].join('\r\n') + '\r\n';
    downloadBlob(new Blob(['\ufeff' + text], {type:'text/csv;charset=utf-8'}), 'Photos_and_Description.csv');
  }

  // ---------- Files and validation ----------
  function setPhotos(files) {
    state.photos = [...files];
    state.photoIndex = new Map();
    for (const file of state.photos) {
      const key = file.name.toLocaleLowerCase();
      if (!state.photoIndex.has(key)) state.photoIndex.set(key, []);
      state.photoIndex.get(key).push(file);
    }
    const firstPath = state.photos.find(f => f.webkitRelativePath)?.webkitRelativePath || '';
    const folder = firstPath.includes('/') ? firstPath.split('/')[0] : '';
    $('photos-action').textContent = 'Replace photo folder';
    $('photos-name').textContent = folder ? `${folder} · ${fmtInt(state.photos.length)} files` : `${fmtInt(state.photos.length)} files selected`;
    $('photos-name').title = $('photos-name').textContent;
    invalidateProcessedArtifacts();
    updateAll();
  }

  function validateRows() {
    const seen = new Set();
    const statuses = [];
    let matched = 0, missing = 0, duplicates = 0, ambiguous = 0, blanks = 0;
    state.rows.forEach((row, index) => {
      const fileName = String(row.fileName || '').trim();
      const key = fileName.toLocaleLowerCase();
      let status, file = null;
      if (!fileName) {
        status = 'blank'; blanks++;
      } else if (seen.has(key)) {
        status = 'duplicate'; duplicates++;
      } else {
        seen.add(key);
        const candidates = state.photoIndex.get(key) || [];
        if (!state.photos.length) status = 'waiting';
        else if (!candidates.length) { status = 'missing'; missing++; }
        else if (candidates.length > 1) { status = 'ambiguous'; ambiguous++; }
        else { status = 'matched'; file = candidates[0]; matched++; }
      }
      statuses.push({index, row, status, file});
    });
    const folderDuplicateNames = [...state.photoIndex.entries()].filter(([, files]) => files.length > 1).map(([name]) => name);
    state.validation = {statuses, matched, missing, duplicates, ambiguous, blanks, folderDuplicateNames};
    return state.validation;
  }

  function getMatchedTargets(v=validateRows()) {
    return v.statuses.filter(s => s.status === 'matched');
  }

  function statusLabel(status) {
    return ({matched:'✓ Matched', missing:'Missing', duplicate:'Duplicate', ambiguous:'Ambiguous', waiting:'Need folder', blank:'Blank'})[status] || status;
  }
  function statusClass(status) {
    if (status === 'matched') return 'ok';
    if (status === 'missing' || status === 'ambiguous') return 'bad';
    if (status === 'duplicate' || status === 'blank') return 'warn';
    return '';
  }
  function rowClass(status) {
    if (status === 'matched') return 'row-ready';
    if (status === 'missing') return 'row-missing';
    if (status === 'duplicate') return 'row-duplicate';
    if (status === 'ambiguous') return 'row-ambiguous';
    return '';
  }

  function renderTable() {
    const validation = validateRows();
    const statusByIndex = new Map(validation.statuses.map(s => [s.index, s]));
    const frag = document.createDocumentFragment();
    state.rows.forEach((row, index) => {
      const info = statusByIndex.get(index);
      const tr = document.createElement('tr');
      tr.dataset.index = index;
      tr.className = rowClass(info.status);
      tr.innerHTML = `
        <td>${index + 1}</td>
        <td><input data-key="fileName" data-index="${index}" value="${esc(row.fileName)}" aria-label="Row ${index + 1} file name"></td>
        <td><input data-key="description" data-index="${index}" value="${esc(row.description)}" aria-label="Row ${index + 1} description"></td>
        <td><span class="row-status ${statusClass(info.status)}">${esc(statusLabel(info.status))}</span></td>
        <td><div class="row-actions">
          <button class="row-action" data-move="up" data-index="${index}" ${index === 0 ? 'disabled' : ''} title="Move up" aria-label="Move row ${index + 1} up">↑</button>
          <button class="row-action" data-move="down" data-index="${index}" ${index === state.rows.length - 1 ? 'disabled' : ''} title="Move down" aria-label="Move row ${index + 1} down">↓</button>
          <button class="row-action delete" data-remove="${index}" title="Remove" aria-label="Remove row ${index + 1}">×</button>
        </div></td>`;
      frag.append(tr);
    });
    $('schedule-rows').replaceChildren(frag);
    $('schedule-empty').hidden = !!state.rows.length;
    $('row-count').textContent = `${fmtInt(state.rows.length)} row${state.rows.length === 1 ? '' : 's'}`;
    renderValidation(validation);
  }

  function refreshTableStatuses() {
    const v = state.validation || validateRows();
    for (const tr of $('schedule-rows').children) {
      const info = v.statuses[Number(tr.dataset.index)];
      if (!info) continue;
      tr.className = rowClass(info.status);
      const statusEl = tr.querySelector('.row-status');
      if (statusEl) {
        statusEl.className = `row-status ${statusClass(info.status)}`;
        statusEl.textContent = statusLabel(info.status);
      }
    }
  }

  function renderValidation(v=validateRows()) {
    const set = (id, value) => { $(id).textContent = state.rows.length ? fmtInt(value) : '—'; };
    set('metric-matched', v.matched);
    set('metric-missing', v.missing);
    set('metric-duplicates', v.duplicates);
    set('metric-ambiguous', v.ambiguous);
    const badge = $('validation-badge');
    badge.className = 'badge';
    if (!state.rows.length || !state.photos.length) badge.textContent = 'Waiting for files';
    else if (v.matched === 0) { badge.textContent = 'Needs attention'; badge.classList.add('error'); }
    else if (v.missing || v.duplicates || v.ambiguous || v.blanks) { badge.textContent = 'Ready with warnings'; badge.classList.add('warn'); }
    else { badge.textContent = 'Ready'; badge.classList.add('ready'); }

    const notes = [];
    if (!state.photos.length) notes.push('<span class="warn">Select a photo folder to match filenames.</span>');
    if (!state.rows.length) notes.push('Upload the CSV or add schedule rows.');
    if (v.matched) notes.push(`<span class="good">${fmtInt(v.matched)} row${v.matched === 1 ? '' : 's'} will be processed in schedule order.</span>`);
    const missingNames = v.statuses.filter(s => s.status === 'missing').slice(0, 8).map(s => s.row.fileName);
    if (missingNames.length) notes.push(`<span class="bad">Missing:</span> ${missingNames.map(esc).join(', ')}${v.missing > missingNames.length ? `, +${v.missing - missingNames.length} more` : ''}`);
    if (v.duplicates) notes.push(`<span class="warn">${fmtInt(v.duplicates)} later duplicate CSV row${v.duplicates === 1 ? '' : 's'} will be skipped, matching the Python script’s first-occurrence behavior.</span>`);
    if (v.ambiguous) notes.push(`<span class="bad">${fmtInt(v.ambiguous)} CSV filename${v.ambiguous === 1 ? '' : 's'} match more than one selected file. Rename or remove duplicate filenames in subfolders.</span>`);
    if (v.blanks) notes.push(`${fmtInt(v.blanks)} blank filename row${v.blanks === 1 ? '' : 's'} will be ignored.`);
    $('validation-details').innerHTML = notes.length ? notes.map(n => `<div>${n}</div>`).join('') : 'Select a photo folder and upload a CSV to validate the schedule.';
    updateActions();
  }

  function updateSourceSummary() {
    const totalBytes = state.photos.reduce((sum, file) => sum + file.size, 0);
    $('source-summary').innerHTML = `<span><b>Photos/files</b> ${state.photos.length ? fmtInt(state.photos.length) : '—'}</span><span><b>Folder size</b> ${state.photos.length ? fmtBytes(totalBytes) : '—'}</span><span><b>CSV rows</b> ${state.rows.length ? fmtInt(state.rows.length) : '—'}</span>`;
  }
  function updateAll() {
    updateSourceSummary();
    renderTable();
    updateActions();
  }

  // ---------- Settings ----------
  function applyDefaults(defaults) {
    if (!defaults) return;
    $('width-cm').value = defaults.width_cm;
    $('height-cm').value = defaults.height_cm;
    $('dpi').value = defaults.dpi;
    $('jpeg-quality').value = defaults.jpeg_quality;
    $('first-image-number').value = defaults.first_image_number;
    $('first-rel-number').value = defaults.first_rel_id_number;
    for (const id of ['width-cm','height-cm','dpi','jpeg-quality','first-image-number','first-rel-number']) $(id).disabled = false;
    $('reset-settings').disabled = false;
    updateSettingsNote();
  }
  function settings() {
    return {
      width_cm: $('width-cm').value,
      height_cm: $('height-cm').value,
      dpi: $('dpi').value,
      jpeg_quality: $('jpeg-quality').value,
      first_image_number: $('first-image-number').value,
      first_rel_id_number: $('first-rel-number').value,
    };
  }
  function updateSettingsNote() {
    const w = Number($('width-cm').value), h = Number($('height-cm').value), dpi = Number($('dpi').value);
    if (w > 0 && h > 0 && dpi > 0) {
      const pxW = Math.round(w / 2.54 * dpi), pxH = Math.round(h / 2.54 * dpi);
      $('settings-note').textContent = `Centered crop · ${pxW.toLocaleString()} × ${pxH.toLocaleString()} px · EXIF orientation correction · RGB JPEG output.`;
    }
  }

  function updateOutputModeLabel() {
    const streaming = $('stream-output').checked;
    const supported = 'showSaveFilePicker' in window;
    $('output-mode-label').textContent = streaming
      ? (supported ? 'Streaming to selected file' : 'Streaming unavailable')
      : 'Normal download';
    $('generate-note').textContent = streaming
      ? (supported ? 'The final DOCX will be written directly to the file location you choose.' : 'Streaming requires a current Chrome or Edge browser.')
      : 'The final DOCX will be built in memory and downloaded normally.';
  }

  function updateActions() {
    const v = state.validation || validateRows();
    const canRun = state.ready && state.templateReady && v.matched > 0 && !state.working;
    $('save-csv').disabled = !state.rows.length || state.working;
    $('folder-to-rows').disabled = !state.photos.length || state.working;
    $('add-row').disabled = state.working;
    $('preview').disabled = !canRun;
    $('generate').disabled = !canRun;
    for (const id of ['width-cm','height-cm','dpi','jpeg-quality','first-image-number','first-rel-number','output-name','preview-mode','stream-output']) {
      const el = $(id);
      if (el) el.disabled = (id !== 'output-name' && id !== 'preview-mode' && id !== 'stream-output' && !state.ready) || state.working;
    }
    $('reset-settings').disabled = !state.ready || state.working;
    document.body.classList.toggle('busy', state.working);
    updateOutputModeLabel();
  }

  // ---------- Download ----------
  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }

  // ---------- Streaming DOCX writer (optional output mode) ----------
  const CRC_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
      table[n] = c >>> 0;
    }
    return table;
  })();
  function crcUpdate(crc, bytes) {
    let c = crc >>> 0;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    return c >>> 0;
  }
  function le16(n) { const a = new Uint8Array(2); new DataView(a.buffer).setUint16(0, n & 0xffff, true); return a; }
  function le32(n) { const a = new Uint8Array(4); new DataView(a.buffer).setUint32(0, n >>> 0, true); return a; }
  function joinBytes(parts) {
    const length = parts.reduce((n, part) => n + part.length, 0);
    const out = new Uint8Array(length); let offset = 0;
    for (const part of parts) { out.set(part, offset); offset += part.length; }
    return out;
  }
  function dosDateTime(date=new Date()) {
    const year = Math.max(1980, date.getFullYear());
    const time = (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2);
    const day = ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate();
    return {time, date:day};
  }
  async function* sourceChunks(source) {
    if (source instanceof Uint8Array) { yield source; return; }
    if (source instanceof ArrayBuffer) { yield new Uint8Array(source); return; }
    if (source instanceof Blob) {
      const reader = source.stream().getReader();
      try {
        while (true) {
          const {value, done} = await reader.read();
          if (done) break;
          if (value?.length) yield value;
        }
      } finally { reader.releaseLock(); }
      return;
    }
    throw new Error('Unsupported ZIP entry source.');
  }
  class WritableFileSink {
    constructor(writable){ this.writable = writable; }
    async write(bytes){ await this.writable.write(bytes); }
    async close(){ await this.writable.close(); }
    async abort(){ try { await this.writable.abort(); } catch (_) {} }
  }
  class StoreZipWriter {
    constructor(sink) { this.sink = sink; this.offset = 0; this.central = []; }
    async #write(bytes) {
      if (!(bytes instanceof Uint8Array)) bytes = new Uint8Array(bytes);
      await this.sink.write(bytes);
      this.offset += bytes.length;
      if (this.offset > 0xffffffff) throw new Error('The generated DOCX exceeded the 4 GB limit of streaming mode.');
    }
    async add(name, source) {
      if (this.central.length >= 65535) throw new Error('The generated DOCX has too many ZIP entries.');
      const nameBytes = encoder.encode(name.replace(/\\/g, '/'));
      const localOffset = this.offset;
      const dt = dosDateTime();
      const flags = 0x0808;
      await this.#write(joinBytes([
        le32(0x04034b50), le16(20), le16(flags), le16(0), le16(dt.time), le16(dt.date),
        le32(0), le32(0), le32(0), le16(nameBytes.length), le16(0), nameBytes,
      ]));
      let crc = 0xffffffff, size = 0;
      for await (const chunk of sourceChunks(source)) {
        crc = crcUpdate(crc, chunk); size += chunk.length;
        if (size > 0xffffffff) throw new Error(`ZIP entry is larger than 4 GB: ${name}`);
        await this.#write(chunk);
      }
      crc = (crc ^ 0xffffffff) >>> 0;
      await this.#write(joinBytes([le32(0x08074b50), le32(crc), le32(size), le32(size)]));
      this.central.push({nameBytes, flags, dt, crc, size, localOffset});
    }
    async finish() {
      const cdOffset = this.offset;
      for (const e of this.central) {
        await this.#write(joinBytes([
          le32(0x02014b50), le16(20), le16(20), le16(e.flags), le16(0), le16(e.dt.time), le16(e.dt.date),
          le32(e.crc), le32(e.size), le32(e.size), le16(e.nameBytes.length), le16(0), le16(0), le16(0), le16(0), le32(0),
          le32(e.localOffset), e.nameBytes,
        ]));
      }
      const cdSize = this.offset - cdOffset;
      const count = this.central.length;
      await this.#write(joinBytes([
        le32(0x06054b50), le16(0), le16(0), le16(count), le16(count), le32(cdSize), le32(cdOffset), le16(0),
      ]));
      await this.sink.close();
    }
  }

  async function streamDocxToHandle(saveHandle, xmlResult, processedEntries) {
    const reader = new TemplateZipReader(state.templateBuffer.slice(0));
    const generatedMedia = new Set(processedEntries.map(e => `word/media/${e.image_name}`));
    const skip = new Set(['word/document.xml', 'word/_rels/document.xml.rels', ...generatedMedia]);
    const sink = new WritableFileSink(await saveHandle.createWritable());
    const writer = new StoreZipWriter(sink);
    try {
      for (const entry of reader.entries) {
        if (entry.name.endsWith('/') || skip.has(entry.name)) continue;
        await writer.add(entry.name, await reader.getData(entry));
      }
      await writer.add('word/document.xml', encoder.encode(xmlResult.document_xml));
      await writer.add('word/_rels/document.xml.rels', encoder.encode(xmlResult.document_rels_xml));
      for (const entry of processedEntries) {
        await writer.add(`word/media/${entry.image_name}`, entry.blob);
      }
      await writer.finish();
    } catch (error) {
      await sink.abort();
      throw error;
    }
  }

  // ---------- Preview / processing cache ----------
  async function clearStageSilently() {
    if (!state.ready || !state.worker) return;
    try { await invoke({action:'clear_stage'}); } catch (_) {}
  }

  function disposeProcessCache() {
    if (!state.processCache?.items?.length) {
      state.processCache = null;
      return;
    }
    for (const item of state.processCache.items) {
      if (item.fullUrl) URL.revokeObjectURL(item.fullUrl);
      if (item.smallUrl) URL.revokeObjectURL(item.smallUrl);
    }
    state.processCache = null;
  }

  function hidePreviewPanel() {
    state.previewVisible = false;
    $('hide-preview').hidden = true;
    $('preview-pages').innerHTML = '<div class="preview-empty"><strong>Photo log preview</strong><p>The preview will appear here and scroll independently from the photo schedule.</p></div>';
    $('preview-meta').textContent = 'Generate a preview to review the processed crop, order, and labels before creating the DOCX.';
    $('preview-badge').className = 'badge';
    $('preview-badge').textContent = 'Not generated';
  }

  function invalidateProcessedArtifacts() {
    disposeProcessCache();
    state.result = null;
    $('result-summary').hidden = true;
    hidePreviewPanel();
    clearStageSilently();
    resetProgress('Ready when files are validated');
    updateActions();
  }

  async function ensureSmallPreview(item) {
    if (item.smallUrl) return item.smallUrl;
    const image = await createImageBitmap(item.blob);
    const scale = Math.min(PREVIEW_SMALL_MAX_W / image.width, PREVIEW_SMALL_MAX_H / image.height, 1);
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(image.width * scale));
    canvas.height = Math.max(1, Math.round(image.height * scale));
    const ctx = canvas.getContext('2d', {alpha:false});
    ctx.drawImage(image, 0, 0, canvas.width, canvas.height);
    image.close();
    const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', 0.8));
    item.smallUrl = URL.createObjectURL(blob || item.blob);
    return item.smallUrl;
  }

  function processingSignature(validatedSettings, targets) {
    return JSON.stringify({
      width_cm: validatedSettings.width_cm,
      height_cm: validatedSettings.height_cm,
      dpi: validatedSettings.dpi,
      jpeg_quality: validatedSettings.jpeg_quality,
      first_image_number: validatedSettings.first_image_number,
      files: targets.map(t => ({
        name: t.file.name,
        size: t.file.size,
        lastModified: t.file.lastModified,
      })),
    });
  }

  function syncCachedEntryMetadata(items, targets, validatedSettings) {
    items.forEach((item, i) => {
      const target = targets[i];
      item.source_name = target.file.name;
      item.description = target.row.description;
      item.image_name = `image${validatedSettings.first_image_number + i}.jpg`;
      item.rel_id = `rId${validatedSettings.first_rel_id_number + i}`;
      item.index = i + 1;
    });
  }

  function resetProgress(text='Ready when files are validated') {
    $('progress-count').textContent = '0 / 0';
    $('progress-bar').style.width = '0%';
    $('progress-file').textContent = text;
    $('progress-size').textContent = '';
  }

  function showProgress(current, total, filename, bytes=0, phase='Processing photos') {
    $('progress-count').textContent = `${fmtInt(current)} / ${fmtInt(total)}`;
    $('progress-bar').style.width = `${total ? Math.min(100, current / total * 100) : 0}%`;
    $('progress-file').textContent = filename ? `${phase} · ${filename}` : phase;
    $('progress-size').textContent = bytes ? fmtBytes(bytes) : '';
  }

  async function ensureProcessedImages(reasonLabel='Preview') {
    const v = validateRows();
    const targets = getMatchedTargets(v);
    if (!targets.length) throw new Error('No matched photos are available to process.');

    const validated = await invoke({action:'validate_settings', settings:settings()});
    const s = validated.settings;
    const signature = processingSignature(s, targets);

    if (state.processCache && state.processCache.signature === signature && state.processCache.items.length === targets.length) {
      syncCachedEntryMetadata(state.processCache.items, targets, s);
      return {settings:s, validation:v, processed:state.processCache.items, reused:true};
    }

    disposeProcessCache();
    await clearStageSilently();

    const processed = [];
    const failed = [];
    let processedBytes = 0;
    showProgress(0, targets.length, 'Preparing temporary processed images…', 0, `${reasonLabel}: processing photos`);

    for (let i = 0; i < targets.length; i++) {
      if (state.cancelRequested) break;
      const target = targets[i];
      const file = target.file;
      const imageName = `image${s.first_image_number + processed.length}.jpg`;
      const relId = `rId${s.first_rel_id_number + processed.length}`;
      showProgress(i, targets.length, file.name, file.size, `${reasonLabel}: processing photos`);
      try {
        const inputBuffer = await file.arrayBuffer();
        const response = await invoke({
          action:'process_image',
          filename:file.name,
          image_name:imageName,
          buffer:inputBuffer,
          settings:s,
          return_bytes:true,
        }, [inputBuffer]);
        const meta = response.result;
        const blob = new Blob([response.buffer], {type:'image/jpeg'});
        const fullUrl = URL.createObjectURL(blob);
        processedBytes += meta.output_bytes || 0;
        processed.push({
          index: processed.length + 1,
          source_name: file.name,
          description: target.row.description,
          image_name: imageName,
          rel_id: relId,
          output_bytes: meta.output_bytes || 0,
          width_px: meta.width_px,
          height_px: meta.height_px,
          blob,
          fullUrl,
          smallUrl: null,
        });
      } catch (error) {
        failed.push({name:file.name, error:error.message});
      }
      showProgress(i + 1, targets.length, file.name, file.size, `${reasonLabel}: processing photos`);
      await new Promise(requestAnimationFrame);
    }

    if (state.cancelRequested) {
      disposeProcessCache();
      await clearStageSilently();
      return {settings:s, validation:v, processed:[], failed, cancelled:true};
    }
    if (!processed.length) throw new Error('No photos were successfully processed.');

    state.processCache = {
      signature,
      settings:s,
      items:processed,
      failed,
      processedBytes,
    };
    return {settings:s, validation:v, processed:processed, failed, reused:false};
  }

  function splitIntoPairs(items) {
    const pairs = [];
    for (let i = 0; i < items.length; i += 2) pairs.push([items[i], items[i + 1] || null]);
    return pairs;
  }

  async function renderPreviewFromCache() {
    if (!state.processCache?.items?.length) {
      hidePreviewPanel();
      return;
    }
    const mode = $('preview-mode').value;
    const items = state.processCache.items;
    $('preview-badge').className = 'badge';
    $('preview-badge').textContent = 'Rendering…';
    $('preview-pages').replaceChildren();
    $('hide-preview').hidden = false;

    const metaText = [
      `${fmtInt(items.length)} processed photo${items.length === 1 ? '' : 's'}`,
      mode === 'full' ? 'full-resolution preview' : 'smaller preview',
      `${items[0].width_px.toLocaleString()} × ${items[0].height_px.toLocaleString()} px processed image size`,
      'approximate photo-table pagination (3 rows per preview page)',
    ];
    $('preview-meta').textContent = metaText.join(' · ');

    const pairs = splitIntoPairs(items);
    const pages = [];
    for (let i = 0; i < pairs.length; i += PREVIEW_ROWS_PER_PAGE) pages.push(pairs.slice(i, i + PREVIEW_ROWS_PER_PAGE));

    const container = $('preview-pages');
    let photoCounter = 0;
    for (let pageIndex = 0; pageIndex < pages.length; pageIndex++) {
      const pagePairs = pages[pageIndex];
      const sheet = document.createElement('article');
      sheet.className = 'preview-sheet';
      const header = document.createElement('header');
      header.className = 'preview-sheet-header';
      header.innerHTML = `<div class="preview-sheet-title">Photo Table: Multiple Photos</div><div class="preview-sheet-page">Page ${pageIndex + 1}</div>`;
      const body = document.createElement('div');
      body.className = 'preview-sheet-body';

      for (const pair of pagePairs) {
        const row = document.createElement('div');
        row.className = 'preview-row';
        for (const item of pair) {
          if (!item) {
            const blank = document.createElement('div');
            blank.className = 'preview-cell blank';
            blank.innerHTML = `<div class="preview-photo blank"></div><div class="preview-caption blank"></div>`;
            row.append(blank);
            continue;
          }
          photoCounter += 1;
          const src = mode === 'small' ? await ensureSmallPreview(item) : item.fullUrl;
          const caption = sentenceCase(item.description, photoStem(item.source_name));
          const cell = document.createElement('div');
          cell.className = 'preview-cell';
          cell.innerHTML = `
            <div class="preview-photo"><img src="${src}" alt="Preview photo ${photoCounter}" loading="lazy" decoding="async"></div>
            <div class="preview-caption">Photo ${photoCounter} ${esc(caption)}</div>`;
          row.append(cell);
        }
        body.append(row);
      }

      sheet.append(header, body);
      container.append(sheet);
    }
    state.previewVisible = true;
    $('preview-badge').className = 'badge ready';
    $('preview-badge').textContent = 'Preview ready';
  }

  async function previewPhotoLog() {
    if ($('preview').disabled) return;
    state.result = null;
    state.cancelRequested = false;
    state.working = true;
    $('cancel').hidden = false;
    $('cancel').disabled = false;
    $('cancel').textContent = 'Cancel after current photo';
    $('result-summary').hidden = true;
    message('');
    updateActions();

    try {
      const {processed, cancelled, reused} = await ensureProcessedImages('Preview');
      if (cancelled) {
        message('Preview generation cancelled after the current photo. Processed preview images were discarded.', 'info');
        hidePreviewPanel();
        return;
      }
      showProgress(processed.length, processed.length, 'Rendering preview…', state.processCache?.processedBytes || 0, 'Preview');
      await renderPreviewFromCache();
      message(reused ? 'Preview refreshed from the existing processed-photo cache.' : 'Preview generated successfully.', 'success');
    } catch (error) {
      message(error.message);
      $('progress-file').textContent = 'Preview stopped';
    } finally {
      state.working = false;
      state.cancelRequested = false;
      $('cancel').hidden = true;
      updateActions();
    }
  }

  // ---------- Generation ----------
  async function generateDocx() {
    if ($('generate').disabled) return;
    state.result = null;
    state.cancelRequested = false;

    const outputName = sanitizeOutputName($('output-name').value) + '.docx';
    $('output-name').value = outputName.replace(/\.docx$/i, '');
    const streaming = $('stream-output').checked;
    let saveHandle = null;

    // Streaming mode must ask for the file while this click still has user activation.
    if (streaming) {
      if (!('showSaveFilePicker' in window)) {
        message('Streaming output is not supported by this browser. Switch Output mode back to Normal download or use a current Chrome/Edge browser.');
        return;
      }
      try {
        saveHandle = await window.showSaveFilePicker({
          suggestedName: outputName,
          types: [{description:'Word document', accept:{[DOCX_MIME]:['.docx']}}],
        });
      } catch (error) {
        if (error?.name === 'AbortError') return;
        message(`Could not open the streaming save picker: ${error.message}`);
        return;
      }
    }

    state.working = true;
    $('cancel').hidden = false;
    $('cancel').disabled = false;
    $('cancel').textContent = 'Cancel after current photo';
    $('result-summary').hidden = true;
    message('');
    updateActions();

    try {
      const {settings:s, validation:v, processed, cancelled, failed, reused} = await ensureProcessedImages('Create Word Doc');
      if (cancelled) {
        message('Generation cancelled after the current photo. Processed images were discarded.', 'info');
        return;
      }
      if (!processed.length) throw new Error('No photos were successfully processed. The Word document was not created.');

      const entries = processed.map(item => ({
        source_name: item.source_name,
        description: item.description,
        image_name: item.image_name,
        rel_id: item.rel_id,
        output_bytes: item.output_bytes,
      }));

      let outputBytes = null;
      let odd = false;
      if (streaming) {
        showProgress(processed.length, processed.length, 'Building Word XML…', state.processCache?.processedBytes || 0, 'Create Word Doc');
        const xmlResult = await invoke({action:'build_xml', entries, settings:s});
        odd = xmlResult.odd_photo_count;
        showProgress(processed.length, processed.length, 'Streaming DOCX to selected file…', state.processCache?.processedBytes || 0, 'Create Word Doc');
        await streamDocxToHandle(saveHandle, xmlResult, processed);
      } else {
        showProgress(processed.length, processed.length, 'Building Word document…', state.processCache?.processedBytes || 0, 'Create Word Doc');
        const templateCopy = state.templateBuffer.slice(0);
        const built = await invoke({action:'build_docx', template_buffer:templateCopy, entries, settings:s}, [templateCopy]);
        const blob = new Blob([built.buffer], {type:DOCX_MIME});
        outputBytes = blob.size;
        odd = built.result.odd_photo_count;
        downloadBlob(blob, outputName);
      }

      $('progress-bar').style.width = '100%';
      $('progress-file').textContent = `${outputName} created`;
      $('progress-size').textContent = streaming ? 'streamed to selected file' : `${fmtBytes(outputBytes)} Word document`;
      renderResult({
        outputName,
        outputBytes,
        processed,
        failed: failed || [],
        validation: v,
        odd,
        reused,
        streaming,
      });
      message(`${outputName} created successfully.`, 'success');
    } catch (error) {
      message(error.message);
      $('progress-file').textContent = 'Generation stopped';
    } finally {
      state.working = false;
      state.cancelRequested = false;
      $('cancel').hidden = true;
      updateActions();
    }
  }

  function renderResult(result) {
    state.result = result;
    const skipped = result.validation.missing + result.validation.duplicates + result.validation.ambiguous + result.validation.blanks;
    const items = [];
    const sizeText = result.outputBytes ? ` (${fmtBytes(result.outputBytes)})` : '';
    items.push(`<span><strong>${fmtInt(result.processed.length)}</strong> photos · ${esc(result.outputName)}${sizeText}</span>`);
    items.push(`<span>Output: <strong>${result.streaming ? 'streamed directly to selected file' : 'normal browser download'}</strong></span>`);
    if (result.reused) items.push('<span>Processed-photo cache reused</span>');
    if (skipped) items.push(`<span>${fmtInt(skipped)} schedule row${skipped === 1 ? '' : 's'} skipped</span>`);
    if (result.failed.length) items.push(`<span>${fmtInt(result.failed.length)} matched file${result.failed.length === 1 ? '' : 's'} failed</span>`);
    if (result.odd) items.push('<span>Odd final photo: right-hand cell left blank</span>');
    $('result-summary').innerHTML = items.join(' <span class="result-sep">·</span> ');
    $('result-summary').hidden = false;
  }

  // ---------- Events ----------
  $('photos-input').addEventListener('change', e => {
    if (!state.working && e.target.files?.length) setPhotos(e.target.files);
    e.target.value = '';
  });

  $('csv-input').addEventListener('change', async e => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file || state.working) return;
    try {
      state.rows = csvTextToRows(await file.text());
      $('csv-action').textContent = 'Replace CSV';
      $('csv-name').textContent = file.name;
      $('csv-name').title = file.name;
      invalidateProcessedArtifacts();
      message('');
      updateAll();
    } catch (error) { message(error.message); }
  });


  for (const [id, handler] of [
    ['photos-drop', files => setPhotos(files)],
    ['csv-drop', async files => {
      const file = files[0];
      if (!file) return;
      state.rows = csvTextToRows(await file.text());
      $('csv-action').textContent = 'Replace CSV';
      $('csv-name').textContent = file.name;
      $('csv-name').title = file.name;
      invalidateProcessedArtifacts();
      updateAll();
    }],
  ]) {
    const el = $(id);
    el.addEventListener('dragover', e => { e.preventDefault(); el.classList.add('dragover'); });
    el.addEventListener('dragleave', () => el.classList.remove('dragover'));
    el.addEventListener('drop', async e => {
      e.preventDefault();
      el.classList.remove('dragover');
      try {
        if (!state.working && e.dataTransfer.files?.length) await handler([...e.dataTransfer.files]);
      } catch (error) {
        message(error.message);
      }
    });
  }

  $('schedule-rows').addEventListener('input', e => {
    const input = e.target.closest('input[data-key]');
    if (!input || state.working) return;
    const row = state.rows[Number(input.dataset.index)];
    row[input.dataset.key] = input.value;
    updateSourceSummary();
    renderValidation();
    refreshTableStatuses();
    if (input.dataset.key === 'fileName') {
      invalidateProcessedArtifacts();
    } else if (state.processCache) {
      syncCachedEntryMetadata(state.processCache.items, getMatchedTargets(), state.processCache.settings);
      if (state.previewVisible) renderPreviewFromCache().catch(err => message(err.message));
    }
  });

  $('schedule-rows').addEventListener('click', e => {
    if (state.working) return;
    const remove = e.target.closest('[data-remove]');
    if (remove) {
      state.rows.splice(Number(remove.dataset.remove), 1);
      invalidateProcessedArtifacts();
      updateAll();
      return;
    }
    const move = e.target.closest('[data-move]');
    if (move) {
      const i = Number(move.dataset.index), j = move.dataset.move === 'up' ? i - 1 : i + 1;
      if (j >= 0 && j < state.rows.length) [state.rows[i], state.rows[j]] = [state.rows[j], state.rows[i]];
      invalidateProcessedArtifacts();
      updateAll();
    }
  });

  $('add-row').addEventListener('click', () => {
    state.rows.push({fileName:'', description:'', sourceLine:null});
    invalidateProcessedArtifacts();
    updateAll();
    $('schedule-rows').lastElementChild?.querySelector('[data-key="fileName"]')?.focus();
  });

  $('folder-to-rows').addEventListener('click', () => {
    if (!state.photos.length) return;
    if (state.rows.length && !confirm('Replace the current schedule with the selected folder filenames?')) return;
    const candidates = state.photos.filter(isLikelyImage);
    state.rows = candidates.sort((a,b) => (a.webkitRelativePath || a.name).localeCompare(b.webkitRelativePath || b.name, undefined, {numeric:true})).map(file => ({fileName:file.name, description:photoStem(file.name), sourceLine:null}));
    invalidateProcessedArtifacts();
    updateAll();
  });

  $('save-csv').addEventListener('click', saveCSV);

  for (const id of ['width-cm','height-cm','dpi','jpeg-quality','first-image-number','first-rel-number']) {
    $(id).addEventListener('input', () => {
      updateSettingsNote();
      invalidateProcessedArtifacts();
    });
  }
  $('reset-settings').addEventListener('click', () => {
    applyDefaults(state.options?.defaults);
    invalidateProcessedArtifacts();
  });

  $('stream-output').addEventListener('change', updateOutputModeLabel);

  $('preview-mode').addEventListener('change', () => {
    if (state.previewVisible && state.processCache) renderPreviewFromCache().catch(err => message(err.message));
  });
  $('preview').addEventListener('click', previewPhotoLog);
  $('hide-preview').addEventListener('click', hidePreviewPanel);
  $('generate').addEventListener('click', generateDocx);
  $('cancel').addEventListener('click', () => {
    state.cancelRequested = true;
    $('cancel').disabled = true;
    $('cancel').textContent = 'Cancelling…';
  });


  renderTable();
  updateSourceSummary();
  hidePreviewPanel();
  resetProgress();
  updateOutputModeLabel();
  updateActions();
  Promise.allSettled([bootPython(), loadBundledTemplate()]);
})();
