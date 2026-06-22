/* ===== DOM REFERENCES ===== */
const dropZone        = document.getElementById('dropZone');
const fileInput       = document.getElementById('fileInput');
const fileName        = document.getElementById('fileName');
const processingBar   = document.getElementById('processingBar');
const processingMsg   = document.getElementById('processingMsg');
const resultsSection  = document.getElementById('resultsSection');
const errorBanner     = document.getElementById('errorBanner');
const errorMessage    = document.getElementById('errorMessage');
const fileInfo        = document.getElementById('fileInfo');
const breakdownBody   = document.getElementById('breakdownBody');
const paginationInfo  = document.getElementById('paginationInfo');
const resetBtn        = document.getElementById('resetBtn');

/* Counter display elements */
const elCntPCombined       = document.getElementById('cntPCombined');
const elCntBRemarkContains = document.getElementById('cntBRemarkContains');
const elCntMTypeEquals     = document.getElementById('cntMTypeEquals');
const elCntCUnique         = document.getElementById('cntCUnique');

/* ===== DRAG & DROP ===== */
dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));

dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const files = Array.from(e.dataTransfer.files).filter(f => f.name.match(/\.(xlsx|xls)$/i));
  if (files.length) handleFiles(files);
  else showError('Please drop valid Excel files (.xlsx or .xls).');
});

dropZone.addEventListener('click', (e) => { if (e.target.tagName !== 'LABEL') fileInput.click(); });
fileInput.addEventListener('change', () => {
  const files = Array.from(fileInput.files);
  if (files.length) handleFiles(files);
});

resetBtn.addEventListener('click', resetAll);

/* ===== FILE HANDLER ===== */
function handleFiles(files) {
  hideError();
  resultsSection.classList.add('hidden');

  const invalid = files.filter(f => !f.name.match(/\.(xlsx|xls)$/i));
  if (invalid.length) {
    showError(`Some files are not valid Excel files: ${invalid.map(f => f.name).join(', ')}`);
    return;
  }

  fileName.textContent = files.length === 1
    ? files[0].name
    : `${files.length} files selected: ${files.map(f => f.name).join(', ')}`;

  showProcessing(`Reading ${files.length} file${files.length > 1 ? 's' : ''}…`);

  Promise.all(files.map(f => readFileAsArrayBuffer(f)))
    .then(buffers => {
      setTimeout(() => {
        try { processAllWorkbooks(buffers, files); }
        catch (err) {
          hideProcessing();
          showError('Could not read one or more files. Make sure they are valid Excel workbooks.');
          console.error(err);
        }
      }, 50);
    })
    .catch(err => { hideProcessing(); showError('Failed to read the file(s).'); console.error(err); });
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload  = (e) => resolve({ buffer: e.target.result, file });
    fr.onerror = () => reject(new Error(`Failed to read ${file.name}`));
    fr.readAsArrayBuffer(file);
  });
}

/* ===== PROCESS ALL WORKBOOKS ===== */
function processAllWorkbooks(fileData, files) {
  setProcessingMsg('Parsing workbooks…');

  // Per-file stats
  const fileStats = []; // { name, predictive, broadcast, manual, connected }
  let totalFilesInfo = [];

  // Cross-file dedup for connected accounts
  const seenAccountNos = new Map();

  // Grand totals
  let totalPredictive = 0;
  let totalBroadcast  = 0;
  let totalManual     = 0;
  let totalConnected  = 0;
  let totalRows       = 0;

  for (const { buffer, file } of fileData) {
    const workbook  = XLSX.read(buffer, { type: 'array', cellDates: false });
    const sheetName = workbook.SheetNames[0];
    const sheet     = workbook.Sheets[sheetName];
    const rows      = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });

    totalRows += rows.length;
    totalFilesInfo.push(`${file.name} (${rows.length.toLocaleString()} rows, sheet "${sheetName}")`);

    setProcessingMsg('Detecting columns…');

    const keys          = Object.keys(rows.find(r => Object.keys(r).length > 0) || rows[0] || {});
    const remarkKey     = keys.find(k => k.trim().toLowerCase() === 'remark');
    const remarkTypeKey = keys.find(k => k.trim().toLowerCase() === 'remark type');
    const callStatusKey = keys.find(k => k.trim().toLowerCase() === 'call status');
    const accountNoKey  = keys.find(k => k.trim().toLowerCase() === 'account no.' || k.trim().toLowerCase() === 'account no');

    if (!remarkKey)     { hideProcessing(); showError('Column "Remark" not found. Check your file headers.'); return; }
    if (!remarkTypeKey) { hideProcessing(); showError('Column "Remark Type" not found. Check your file headers.'); return; }

    let filePredictive = 0;
    let fileBroadcast  = 0;
    let fileManual     = 0;
    let fileConnected  = 0;

    // Per-file seen accounts (for per-file unique connected)
    const fileSeenAccounts = new Map();

    setProcessingMsg(`Scanning ${rows.length.toLocaleString()} rows in ${file.name}…`);

    rows.forEach(row => {
      const remarkRaw = String(row[remarkKey]     ?? '');
      const typeRaw   = String(row[remarkTypeKey] ?? '');
      const remarkLow = remarkRaw.toLowerCase();
      const typeLow   = typeRaw.trim().toLowerCase();

      const p_remark = remarkLow.includes('predictive');
      const p_type   = typeLow === 'predictive';
      const b_remark = remarkLow.includes('broadcast');
      const m_type   = typeLow === 'outgoing';

      if (p_remark || p_type) filePredictive++;
      if (b_remark)           fileBroadcast++;
      if (m_type)             fileManual++;

      /* Connected: unique per file, also track cross-file global unique */
      if (callStatusKey) {
        const statusRaw = String(row[callStatusKey] ?? '').trim().toUpperCase();
        if (statusRaw === 'CONNECTED') {
          const acctRaw = accountNoKey ? String(row[accountNoKey] ?? '').trim() : '';
          const acctKey = acctRaw.toLowerCase();

          if (!fileSeenAccounts.has(acctKey)) {
            fileSeenAccounts.set(acctKey, true);
            fileConnected++;
          }
          // Also track global cross-file unique
          if (!seenAccountNos.has(acctKey)) {
            seenAccountNos.set(acctKey, true);
          }
        }
      }
    });

    totalPredictive += filePredictive;
    totalBroadcast  += fileBroadcast;
    totalManual     += fileManual;
    totalConnected  += fileConnected;

    fileStats.push({
      name:       file.name,
      predictive: filePredictive,
      broadcast:  fileBroadcast,
      manual:     fileManual,
      connected:  fileConnected,
    });
  }

  // Global unique connected = cross-file deduplicated count
  const globalUniqueConnected = seenAccountNos.size;

  if (!totalRows) {
    hideProcessing();
    showError('All uploaded sheets appear to be empty or have no data rows.');
    return;
  }

  /* Animate summary counters */
  animateCount(elCntPCombined,       totalPredictive);
  animateCount(elCntBRemarkContains, totalBroadcast);
  animateCount(elCntMTypeEquals,     totalManual);
  animateCount(elCntCUnique,         globalUniqueConnected);

  /* File meta */
  if (files.length === 1) {
    const f = files[0];
    fileInfo.textContent =
      `${f.name}  ·  ${totalRows.toLocaleString()} rows scanned  ·  ${(f.size / 1024).toFixed(1)} KB`;
  } else {
    fileInfo.textContent =
      `${files.length} files merged  ·  ${totalRows.toLocaleString()} total rows  ·  ` +
      totalFilesInfo.join('  |  ');
  }

  /* Build per-file summary table */
  buildSummaryTable(fileStats, globalUniqueConnected, files.length);

  hideProcessing();
  resultsSection.classList.remove('hidden');
  resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/* ===== PER-FILE SUMMARY TABLE ===== */
function buildSummaryTable(fileStats, globalUniqueConnected, fileCount) {
  const frag = document.createDocumentFragment();
  breakdownBody.innerHTML = '';

  fileStats.forEach(stat => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="source-file-cell" title="${escapeHtml(stat.name)}">${escapeHtml(stat.name)}</td>
      <td class="col-predictive num-cell">${stat.predictive.toLocaleString()}</td>
      <td class="col-broadcast num-cell">${stat.broadcast.toLocaleString()}</td>
      <td class="col-manual num-cell">${stat.manual.toLocaleString()}</td>
      <td class="col-connected num-cell">${stat.connected.toLocaleString()}</td>
    `;
    frag.appendChild(tr);
  });

  /* Totals row (only when multiple files) */
  if (fileCount > 1) {
    const totals = fileStats.reduce((acc, s) => {
      acc.predictive += s.predictive;
      acc.broadcast  += s.broadcast;
      acc.manual     += s.manual;
      acc.connected  += s.connected;
      return acc;
    }, { predictive: 0, broadcast: 0, manual: 0, connected: 0 });

    const tr = document.createElement('tr');
    tr.classList.add('totals-row');
    tr.innerHTML = `
      <td class="totals-label">Total (${fileCount} files)</td>
      <td class="col-predictive num-cell">${totals.predictive.toLocaleString()}</td>
      <td class="col-broadcast num-cell">${totals.broadcast.toLocaleString()}</td>
      <td class="col-manual num-cell">${totals.manual.toLocaleString()}</td>
      <td class="col-connected num-cell">${globalUniqueConnected.toLocaleString()} <span class="unique-note">unique</span></td>
    `;
    frag.appendChild(tr);
  }

  breakdownBody.appendChild(frag);

  paginationInfo.textContent = fileCount === 1
    ? '1 file'
    : `${fileCount} files`;
}

/* ===== ANIMATE COUNTER ===== */
function animateCount(el, target) {
  if (!el) return;
  if (target === 0) { el.textContent = '0'; return; }
  const steps = Math.min(40, Math.max(1, target));
  const interval = 600 / steps;
  const inc = target / steps;
  let current = 0;
  const timer = setInterval(() => {
    current += inc;
    if (current >= target) { current = target; clearInterval(timer); }
    el.textContent = Math.round(current).toLocaleString();
  }, interval);
}

/* ===== PROCESSING ===== */
function showProcessing(msg)   { processingMsg.textContent = msg; processingBar.classList.remove('hidden'); }
function setProcessingMsg(msg) { processingMsg.textContent = msg; }
function hideProcessing()      { processingBar.classList.add('hidden'); }

/* ===== ERROR ===== */
function showError(msg) { errorMessage.textContent = msg; errorBanner.classList.remove('hidden'); }
function hideError()    { errorBanner.classList.add('hidden'); }

/* ===== RESET ===== */
function resetAll() {
  fileInput.value            = '';
  fileName.textContent       = 'No files selected';
  resultsSection.classList.add('hidden');
  hideError(); hideProcessing();
  breakdownBody.innerHTML    = '';
  fileInfo.textContent       = '';
  paginationInfo.textContent = '';

  [elCntPCombined, elCntBRemarkContains, elCntMTypeEquals, elCntCUnique]
    .forEach(el => { if (el) el.textContent = '—'; });

  window.scrollTo({ top: 0, behavior: 'smooth' });
}

/* ===== HELPERS ===== */
function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}