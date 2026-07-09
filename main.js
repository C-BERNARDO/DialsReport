/* ===== DOM REFERENCES ===== */
const dropZone        = document.getElementById('dropZone');
const fileInput       = document.getElementById('fileInput');
const fileName        = document.getElementById('fileName');
const processStatus    = document.getElementById('processStatus');
const processCurrent   = document.getElementById('processCurrent');
const processProgress  = document.getElementById('processProgress');
const processDoneList  = document.getElementById('processDoneList');
const processBarFill   = document.getElementById('processBarFill');
const resultsSection  = document.getElementById('resultsSection');
const emptyState      = document.getElementById('emptyState');
const errorBanner     = document.getElementById('errorBanner');
const errorMessage    = document.getElementById('errorMessage');
const fileInfoSummary = document.getElementById('fileInfoSummary');
const fileInfoDetail  = document.getElementById('fileInfoDetail');
const fileInfoToggle  = document.getElementById('fileInfoToggle');
const fileInfoBar     = document.getElementById('fileInfoBar');
const breakdownBody   = document.getElementById('breakdownBody');
const resetBtn        = document.getElementById('resetBtn');

/* Counter display elements */
const elCntPCombined       = document.getElementById('cntPCombined');
const elCntBRemarkContains = document.getElementById('cntBRemarkContains');
const elCntMTypeEquals     = document.getElementById('cntMTypeEquals');
const elCntCUnique         = document.getElementById('cntCUnique');
const elCntConnectedWithClient = document.getElementById('cntConnectedWithClient');
const elCntUniqueWorked    = document.getElementById('cntUniqueWorked');
const elCntAccountsDialed  = document.getElementById('cntAccountsDialed');
const elCntDropSystem      = document.getElementById('cntDropSystem');
const elCntDropClient      = document.getElementById('cntDropClient');
const elCntPU               = document.getElementById('cntPU');
const elCntPM               = document.getElementById('cntPM');

/* ===== DRAG & DROP ===== */
dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));

dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const files = Array.from(e.dataTransfer.files).filter(f => f.name.match(/\.(xlsx|xls|csv)$/i));
  if (files.length) handleFiles(files);
  else showError('Please drop valid Excel (.xlsx, .xls) or CSV (.csv) files.');
});

dropZone.addEventListener('click', (e) => { if (e.target.tagName !== 'LABEL') fileInput.click(); });
fileInput.addEventListener('change', () => {
  const files = Array.from(fileInput.files);
  if (files.length) handleFiles(files);
});

resetBtn.addEventListener('click', resetAll);

fileInfoToggle.addEventListener('click', () => {
  const isHidden = fileInfoDetail.classList.contains('hidden');
  fileInfoDetail.classList.toggle('hidden', !isHidden);
  fileInfoToggle.textContent = isHidden ? 'Hide files' : 'Show files';
});

/* ===== FILE HANDLER ===== */
function handleFiles(files) {
  hideError();
  resultsSection.classList.add('hidden');
  fileInfoBar.classList.add('hidden');

  const invalid = files.filter(f => !f.name.match(/\.(xlsx|xls|csv)$/i));
  if (invalid.length) {
    showError(`Some files are not valid Excel or CSV files: ${invalid.map(f => f.name).join(', ')}`);
    return;
  }

  fileName.textContent = files.length === 1
    ? files[0].name
    : `${files.length} files selected`;

  initProcessingStatus(files.length);

  Promise.all(files.map(f => readFileAsArrayBuffer(f)))
    .then(async buffers => {
      try { await processAllWorkbooks(buffers, files); }
      catch (err) {
        hideProcessing();
        showError('Could not read one or more files. Make sure they are valid Excel or CSV files.');
        console.error(err);
      }
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
async function processAllWorkbooks(fileData, files) {
  // Per-file stats
  const fileStats = []; // { name, predictive, broadcast, manual, connected }
  let totalFilesInfo = [];

  // Cross-file dedup for connected accounts
  const seenAccountNos = new Map();

  // Cross-file dedup for "Connected with Client" accounts
  const seenConnectedClientAccounts = new Map();

  // Cross-file dedup for "Unique Worked Accounts" (Status not NEW / ABORT / UNLOCKED / LOCKED)
  const seenWorkedAccounts = new Map();

  // Cross-file dedup for "Accounts Dialed" (unique accounts hit by
  // Predictive, Broadcast, or Manual dials — merged, not summed)
  const seenDialedAccounts = new Map();

  // Grand totals
  let totalPredictive = 0;
  let totalBroadcast  = 0;
  let totalManual     = 0;
  let totalConnected  = 0;
  let totalDropSystem = 0;
  let totalDropClient = 0;
  let totalPU = 0;
  let totalPM = 0;
  let totalRows       = 0;

  const totalFiles = fileData.length;

  for (let i = 0; i < fileData.length; i++) {
    const { buffer, file } = fileData[i];

    setCurrentProcessingFile(file.name, i, totalFiles);
    await yieldToUI(); // let the browser paint "Processing: <file.name>" before we block on parsing

    let workbook, sheetName, rows;

    if (file.name.match(/\.csv$/i)) {
      /* ── CSV path: read as text, parse with SheetJS ── */
      const text = new TextDecoder('utf-8').decode(buffer);
      workbook  = XLSX.read(text, { type: 'string', cellDates: false });
      sheetName = workbook.SheetNames[0];
      rows      = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '', raw: false });
    } else {
      /* ── Excel path (unchanged) ── */
      workbook  = XLSX.read(buffer, { type: 'array', cellDates: false });
      sheetName = workbook.SheetNames[0];
      rows      = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '', raw: false });
    }

    totalRows += rows.length;
    totalFilesInfo.push(`${file.name} (${rows.length.toLocaleString()} rows, sheet "${sheetName}")`);



    const keys          = Object.keys(rows.find(r => Object.keys(r).length > 0) || rows[0] || {});
    const remarkKey     = keys.find(k => k.trim().toLowerCase() === 'remark');
    const remarkTypeKey = keys.find(k => k.trim().toLowerCase() === 'remark type');
    const callDurationKey = keys.find(k => k.trim().toLowerCase() === 'call duration');
    const accountNoKey    = keys.find(k =>
      k.trim().toLowerCase() === 'account no.' || k.trim().toLowerCase() === 'account no'
    );
    const callStatusKey   = keys.find(k => k.trim().toLowerCase() === 'call status');
    const statusKey       = keys.find(k => k.trim().toLowerCase() === 'status');
    const remarkByKey     = keys.find(k => k.trim().toLowerCase() === 'remark by');

    if (!remarkKey)     { hideProcessing(); showError('Column "Remark" not found. Check your file headers.'); return; }
    if (!remarkTypeKey) { hideProcessing(); showError('Column "Remark Type" not found. Check your file headers.'); return; }

    let filePredictive = 0;
    let fileBroadcast  = 0;
    let fileManual     = 0;
    let fileConnected  = 0;
    let fileConnectedWithClient = 0;
    let fileUniqueWorked        = 0;
    let fileDropSystem = 0;
    let fileDropClient = 0;
    let filePU = 0;
    let filePM = 0;
    let fileAccountsDialed = 0;

    // Per-file seen accounts (for per-file unique connected)
    const fileSeenAccounts = new Map();

    // Per-file seen accounts (for per-file unique "Connected with Client")
    const fileSeenConnectedClientAccounts = new Map();

    // Per-file seen accounts (for per-file unique "Worked Accounts")
    const fileSeenWorkedAccounts = new Map();

    // Per-file seen accounts (for per-file unique "Accounts Dialed")
    const fileSeenDialedAccounts = new Map();


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

      /* ── Accounts Dialed: merge Predictive/Broadcast/Manual matches into
         one set, deduped by Account No. — NOT a sum of the three counts. ── */
      if (p_remark || p_type || b_remark || m_type) {
        const acctRaw = accountNoKey ? String(row[accountNoKey] ?? '').trim() : '';
        const acctKey = acctRaw.toLowerCase();

        if (!fileSeenDialedAccounts.has(acctKey)) {
          fileSeenDialedAccounts.set(acctKey, true);
          fileAccountsDialed++;
        }
        // Also track global cross-file unique
        if (!seenDialedAccounts.has(acctKey)) {
          seenDialedAccounts.set(acctKey, true);
        }
      }

      /* ── Dropped Calls: Call Status is "DROPPED", split by Remark By ── */
      if (callStatusKey) {
        const callStatusLow = String(row[callStatusKey] ?? '').trim().toLowerCase();
        const isDropped = callStatusLow === 'dropped';

        if (isDropped) {
          const remarkByLow = remarkByKey ? String(row[remarkByKey] ?? '').trim().toLowerCase() : '';
          const isSystem = remarkByLow === 'system';

          if (isSystem) fileDropSystem++;
          else          fileDropClient++;
        }
      }

      /* ── PU / PM: raw counts based on Status column, no dedup ── */
      if (statusKey) {
        const statusLow = String(row[statusKey] ?? '').trim().toLowerCase();

        if (statusLow === 'pu') filePU++;
        if (statusLow === 'pm') filePM++;
      }

      /* ── Connected: valid Call Duration (not blank/missing — "00:00:00" now counts), unique Account No. ── */
      if (callDurationKey) {
        const durRaw = String(row[callDurationKey] ?? '').trim();
        const hasValidDuration = durRaw !== '';

        if (hasValidDuration) {
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

      /* ── Connected with Client: unique Account No. where Call Status is "CONNECTED" ── */
      if (callStatusKey) {
        const statusRaw = String(row[callStatusKey] ?? '').trim().toLowerCase();
        const isConnected = statusRaw === 'connected';

        if (isConnected) {
          const acctRaw = accountNoKey ? String(row[accountNoKey] ?? '').trim() : '';
          const acctKey = acctRaw.toLowerCase();

          if (!fileSeenConnectedClientAccounts.has(acctKey)) {
            fileSeenConnectedClientAccounts.set(acctKey, true);
            fileConnectedWithClient++;
          }
          // Also track global cross-file unique
          if (!seenConnectedClientAccounts.has(acctKey)) {
            seenConnectedClientAccounts.set(acctKey, true);
          }
        }
      }

      /* ── Unique Worked Accounts: exclude Status "NEW", "ABORT", "UNLOCKED",
         or "LOCKED", dedupe by Account No. ── */
      if (statusKey) {
        const statusValRaw = String(row[statusKey] ?? '').trim().toLowerCase();
        const isExcluded =
          statusValRaw === 'new' ||
          statusValRaw === 'abort' ||
          statusValRaw === 'unlocked' ||
          statusValRaw === 'locked';

        if (!isExcluded) {
          const acctRaw = accountNoKey ? String(row[accountNoKey] ?? '').trim() : '';
          const acctKey = acctRaw.toLowerCase();

          if (!fileSeenWorkedAccounts.has(acctKey)) {
            fileSeenWorkedAccounts.set(acctKey, true);
            fileUniqueWorked++;
          }
          // Also track global cross-file unique
          if (!seenWorkedAccounts.has(acctKey)) {
            seenWorkedAccounts.set(acctKey, true);
          }
        }
      }
    });

    totalPredictive += filePredictive;
    totalBroadcast  += fileBroadcast;
    totalManual     += fileManual;
    totalConnected  += fileConnected;
    totalDropSystem += fileDropSystem;
    totalDropClient += fileDropClient;
    totalPU         += filePU;
    totalPM         += filePM;

    fileStats.push({
      name:               file.name,
      predictive:         filePredictive,
      broadcast:          fileBroadcast,
      manual:             fileManual,
      connected:          fileConnected,
      connectedWithClient: fileConnectedWithClient,
      uniqueWorked:        fileUniqueWorked,
      dropSystem:          fileDropSystem,
      dropClient:          fileDropClient,
      pu:                  filePU,
      pm:                  filePM,
      accountsDialed:      fileAccountsDialed,
    });

    markFileDone(file.name, rows.length, i, totalFiles);
    await yieldToUI(); // let the browser paint the completed-file update
  }

  // Global unique connected = cross-file deduplicated count
  const globalUniqueConnected = seenAccountNos.size;

  // Global unique "Connected with Client" = cross-file deduplicated count
  const globalUniqueConnectedWithClient = seenConnectedClientAccounts.size;

  // Global unique "Worked Accounts" = cross-file deduplicated count
  const globalUniqueWorked = seenWorkedAccounts.size;

  // Global unique "Accounts Dialed" = merged Predictive/Broadcast/Manual
  // matches, deduplicated by Account No. across all files (not summed)
  const globalAccountsDialed = seenDialedAccounts.size;

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
  animateCount(elCntConnectedWithClient, globalUniqueConnectedWithClient);
  animateCount(elCntUniqueWorked,    globalUniqueWorked);
  animateCount(elCntAccountsDialed,  globalAccountsDialed);
  animateCount(elCntDropSystem,      totalDropSystem);
  animateCount(elCntDropClient,      totalDropClient);
  animateCount(elCntPU,              totalPU);
  animateCount(elCntPM,              totalPM);

  /* File meta — summary line always visible; detailed list auto-hidden
     to reduce clutter, with a toggle to reveal it on demand. */
  fileInfoBar.classList.remove('hidden');
  if (files.length === 1) {
    const f = files[0];
    fileInfoSummary.textContent =
      `${f.name}  ·  ${totalRows.toLocaleString()} rows scanned  ·  ${(f.size / 1024).toFixed(1)} KB`;
    fileInfoToggle.classList.add('hidden');
    fileInfoDetail.classList.add('hidden');
    fileInfoDetail.textContent = '';
  } else {
    fileInfoSummary.textContent =
      `${files.length} files merged  ·  ${totalRows.toLocaleString()} total rows`;
    fileInfoDetail.innerHTML = totalFilesInfo.map(escapeHtml).join('<br>');
    fileInfoToggle.textContent = 'Show files';
    fileInfoToggle.classList.remove('hidden');
    fileInfoDetail.classList.add('hidden'); // auto-collapsed once processing succeeds
  }

  /* Build per-file summary table */
  buildSummaryTable(fileStats);

  hideProcessing();
  emptyState.classList.add('hidden');
  resultsSection.classList.remove('hidden');
  resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/* ===== PER-FILE SUMMARY TABLE ===== */
function buildSummaryTable(fileStats) {
  const frag = document.createDocumentFragment();
  breakdownBody.innerHTML = '';

  fileStats.forEach(stat => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="source-file-cell" title="${escapeHtml(stat.name)}">${escapeHtml(stat.name)}</td>
      <td class="col-worked num-cell">${stat.uniqueWorked.toLocaleString()}</td>
      <td class="col-dialed num-cell">${stat.accountsDialed.toLocaleString()}</td>
      <td class="col-predictive num-cell">${stat.predictive.toLocaleString()}</td>
      <td class="col-broadcast num-cell">${stat.broadcast.toLocaleString()}</td>
      <td class="col-manual num-cell">${stat.manual.toLocaleString()}</td>
      <td class="col-connected num-cell">${stat.connected.toLocaleString()}</td>
      <td class="col-cwc num-cell">${stat.connectedWithClient.toLocaleString()}</td>
      <td class="col-drop-system num-cell">${stat.dropSystem.toLocaleString()}</td>
      <td class="col-drop-client num-cell">${stat.dropClient.toLocaleString()}</td>
      <td class="col-pm num-cell">${stat.pm.toLocaleString()}</td>
      <td class="col-pu num-cell">${stat.pu.toLocaleString()}</td>
    `;
    frag.appendChild(tr);
  });

  breakdownBody.appendChild(frag);
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

/* ===== PROCESSING STATUS ===== */
function initProcessingStatus(totalFiles) {
  processDoneList.innerHTML = '';
  processCurrent.textContent = 'Preparing…';
  processProgress.textContent = `0 of ${totalFiles} file${totalFiles > 1 ? 's' : ''} processed`;
  processBarFill.style.width = '0%';
  processStatus.classList.remove('hidden');
}

function setCurrentProcessingFile(name, index, totalFiles) {
  processCurrent.textContent = `Processing: ${name}`;
  processProgress.textContent = `${index} of ${totalFiles} file${totalFiles > 1 ? 's' : ''} processed`;
  processBarFill.style.width = `${Math.round((index / totalFiles) * 100)}%`;
}

function markFileDone(name, rowCount, index, totalFiles) {
  const li = document.createElement('li');
  li.innerHTML = `
    <svg class="done-check" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
    <span class="done-name" title="${escapeHtml(name)}">${escapeHtml(name)} — ${rowCount.toLocaleString()} rows</span>
  `;
  processDoneList.appendChild(li);
  processDoneList.scrollTop = processDoneList.scrollHeight;

  const doneCount = index + 1;
  processProgress.textContent = `${doneCount} of ${totalFiles} file${totalFiles > 1 ? 's' : ''} processed`;
  processBarFill.style.width = `${Math.round((doneCount / totalFiles) * 100)}%`;
}

/* Yields control back to the browser so it can paint the status panel
   update before the (synchronous) parsing of the next file begins. */
function yieldToUI() {
  return new Promise(resolve => requestAnimationFrame(() => setTimeout(resolve, 0)));
}

function hideProcessing() { processStatus.classList.add('hidden'); }

/* ===== ERROR ===== */
function showError(msg) { errorMessage.textContent = msg; errorBanner.classList.remove('hidden'); }
function hideError()    { errorBanner.classList.add('hidden'); }

/* ===== RESET ===== */
function resetAll() {
  fileInput.value            = '';
  fileName.textContent       = 'No files selected';
  resultsSection.classList.add('hidden');
  emptyState.classList.remove('hidden');
  hideError(); hideProcessing();
  breakdownBody.innerHTML    = '';
  fileInfoSummary.textContent = '';
  fileInfoDetail.textContent  = '';
  fileInfoDetail.classList.add('hidden');
  fileInfoToggle.classList.add('hidden');
  fileInfoBar.classList.add('hidden');

  [elCntPCombined, elCntBRemarkContains, elCntMTypeEquals, elCntAccountsDialed, elCntCUnique, elCntConnectedWithClient, elCntUniqueWorked, elCntDropSystem, elCntDropClient, elCntPU, elCntPM]
    .forEach(el => { if (el) el.textContent = '—'; });

  window.scrollTo({ top: 0, behavior: 'smooth' });
}

/* ===== HELPERS ===== */
function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/* ===== COPY-TO-CLIPBOARD (summary metric cards) ===== */
document.querySelectorAll('.copy-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    const targetEl = document.getElementById(btn.dataset.copyTarget);
    if (!targetEl) return;

    const value = targetEl.textContent.trim();
    if (!value || value === '—') return; // nothing computed yet

    const showCopiedState = () => {
      clearTimeout(btn._copyTimeout);
      btn.classList.add('copied');
      btn._copyTimeout = setTimeout(() => btn.classList.remove('copied'), 1500);
    };

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(value).then(showCopiedState).catch(() => fallbackCopy(value, showCopiedState));
    } else {
      fallbackCopy(value, showCopiedState);
    }
  });
});

function fallbackCopy(text, onSuccess) {
  const ta = document.createElement('textarea');
  ta.value = text;
  ta.style.position = 'fixed';
  ta.style.opacity = '0';
  document.body.appendChild(ta);
  ta.select();
  try { document.execCommand('copy'); onSuccess(); } catch (err) { console.error(err); }
  document.body.removeChild(ta);
}