/**
 * Garment Inventory Verification App Engine
 * Handles Excel parsing, column mapping, barcode scanning (Camera/USB),
 * discrepancy tracking, IndexedDB session persistence, and Excel export.
 */

// Global State
const state = {
  currentSession: null,
  excelRawData: null, // Array of arrays from SheetJS
  excelHeaders: [],
  excelRows: [],
  zxingReader: null,
  isScanningCamera: false,
  activeScanItem: null,
  activeScanRowIndex: null,
  audioCtx: null
};

// Column Schema Definition
const COLUMN_FIELDS = {
  core: [
    { key: 'ean', label: 'EAN / Barcode', required: true, keywords: ['ean', 'barcode', 'upc', 'bar_code', 'code', 'item_code'] },
    { key: 'qty', label: 'Expected Quantity', required: true, keywords: ['qty', 'quantity', 'pcs', 'count', 'expected_qty', 'total_qty', 'units', 'stock'] },
    { key: 'mrp', label: 'MRP / Tag Price', required: false, keywords: ['mrp', 'price', 'tag_mrp', 'retail_price', 'rate_mrp', 'tag_price'] },
    { key: 'size', label: 'Size', required: false, keywords: ['size', 'sz', 'garment_size', 'item_size'] }
  ],
  optional: [
    { key: 'rate', label: 'Purchase Rate / Cost', keywords: ['purchase_rate', 'pur_rate', 'cost', 'buy_price', 'unit_rate', 'rate'] },
    { key: 'brand', label: 'Brand Name', keywords: ['brand', 'company', 'label', 'make'] },
    { key: 'pattern', label: 'Pattern / Print', keywords: ['pattern', 'print', 'style_pattern'] },
    { key: 'section', label: 'Section / Department', keywords: ['section', 'sec', 'department', 'dept', 'category'] },
    { key: 'season', label: 'Season / Collection', keywords: ['season', 'collection'] },
    { key: 'design', label: 'Design Number / Style', keywords: ['design_no', 'design', 'style', 'style_no', 'article_no', 'design_num'] },
    { key: 'nature', label: 'Nature / Item Type', keywords: ['nature', 'item_type', 'item', 'description', 'desc', 'garment_type'] }
  ]
};

// ==========================================
// 1. INDEXEDDB PERSISTENCE MODULE
// ==========================================
const DB = {
  dbName: 'GarmentVerifyDB',
  dbVersion: 1,

  open() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(this.dbName, this.dbVersion);
      request.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains('sessions')) {
          db.createObjectStore('sessions', { keyPath: 'id' });
        }
      };
      request.onsuccess = (e) => resolve(e.target.result);
      request.onerror = (e) => reject(e.target.error);
    });
  },

  async saveSession(session) {
    try {
      const db = await this.open();
      return new Promise((resolve, reject) => {
        const tx = db.transaction('sessions', 'readwrite');
        const store = tx.objectStore('sessions');
        session.updatedAt = new Date().toISOString();
        const req = store.put(session);
        req.onsuccess = () => resolve(session);
        req.onerror = (e) => reject(e.target.error);
      });
    } catch (err) {
      console.error('Failed to save session to IndexedDB:', err);
      showToast('Warning: Session auto-save failed', 'warning');
    }
  },

  async getSession(id) {
    const db = await this.open();
    return new Promise((resolve, reject) => {
      const tx = db.transaction('sessions', 'readonly');
      const store = tx.objectStore('sessions');
      const req = store.get(id);
      req.onsuccess = () => resolve(req.result);
      req.onerror = (e) => reject(e.target.error);
    });
  },

  async getAllSessions() {
    const db = await this.open();
    return new Promise((resolve, reject) => {
      const tx = db.transaction('sessions', 'readonly');
      const store = tx.objectStore('sessions');
      const req = store.getAll();
      req.onsuccess = () => resolve(req.result || []);
      req.onerror = (e) => reject(e.target.error);
    });
  },

  async deleteSession(id) {
    const db = await this.open();
    return new Promise((resolve, reject) => {
      const tx = db.transaction('sessions', 'readwrite');
      const store = tx.objectStore('sessions');
      const req = store.delete(id);
      req.onsuccess = () => resolve();
      req.onerror = (e) => reject(e.target.error);
    });
  }
};

// ==========================================
// 2. AUDIO & VIBRATION FEEDBACK MODULE
// ==========================================
function playBeep(type = 'success') {
  if (!state.currentSession?.settings?.soundEnabled) return;
  try {
    if (!state.audioCtx) {
      state.audioCtx = new (window.AudioContext || window.webkitAudioContext)();
    }
    const ctx = state.audioCtx;
    if (ctx.state === 'suspended') {
      ctx.resume();
    }
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.connect(gain);
    gain.connect(ctx.destination);

    if (type === 'success') {
      osc.frequency.setValueAtTime(880, ctx.currentTime); // A5 tone
      gain.gain.setValueAtTime(0.15, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.12);
      osc.start();
      osc.stop(ctx.currentTime + 0.12);
    } else if (type === 'warning') {
      osc.frequency.setValueAtTime(587.33, ctx.currentTime); // D5 tone
      osc.frequency.setValueAtTime(440, ctx.currentTime + 0.1); // A4 tone
      gain.gain.setValueAtTime(0.2, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.25);
      osc.start();
      osc.stop(ctx.currentTime + 0.25);
    } else if (type === 'error') {
      osc.type = 'sawtooth';
      osc.frequency.setValueAtTime(220, ctx.currentTime);
      gain.gain.setValueAtTime(0.25, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.35);
      osc.start();
      osc.stop(ctx.currentTime + 0.35);
    }
  } catch (e) {
    console.warn('Audio feedback error:', e);
  }

  // Trigger vibration if enabled
  if (state.currentSession?.settings?.vibrateEnabled && navigator.vibrate) {
    if (type === 'success') navigator.vibrate(60);
    else if (type === 'warning') navigator.vibrate([100, 50, 100]);
    else if (type === 'error') navigator.vibrate([200, 100, 200]);
  }
}

// ==========================================
// 3. UI ROUTER & NAVIGATION MODULE
// ==========================================
function switchView(viewId) {
  document.querySelectorAll('.view-panel').forEach(panel => panel.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(btn => btn.classList.remove('active'));

  const targetPanel = document.getElementById(`view-${viewId}`);
  const targetNavBtn = document.querySelector(`.nav-btn[data-view="${viewId}"]`);

  if (targetPanel) targetPanel.classList.add('active');
  if (targetNavBtn) targetNavBtn.classList.add('active');

  // Focus scan input if switching to scan view
  if (viewId === 'scan') {
    setTimeout(() => {
      const input = document.getElementById('input-barcode-scan');
      if (input) input.focus();
    }, 150);
  }

  // Refresh view specific data
  if (viewId === 'overview') renderInventoryOverview();
  if (viewId === 'export') renderExportReview();
}

function updateNavPill() {
  const pill = document.getElementById('active-session-pill');
  const navBtns = ['nav-mapping-btn', 'nav-setup-btn', 'nav-scan-btn', 'nav-overview-btn', 'nav-export-btn'];

  if (state.currentSession) {
    pill.classList.remove('hidden');
    document.getElementById('nav-supplier-name').textContent = state.currentSession.supplierName;
    document.getElementById('nav-invoice-no').textContent = state.currentSession.invoiceNo;
    navBtns.forEach(id => document.getElementById(id)?.classList.remove('hidden'));
  } else {
    pill.classList.add('hidden');
    navBtns.forEach(id => document.getElementById(id)?.classList.add('hidden'));
  }
}

function showToast(message, type = 'info') {
  const container = document.getElementById('toast-container');
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.innerHTML = `<span>${message}</span>`;
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}

// ==========================================
// 4. EXCEL IMPORT & COLUMN MAPPING ENGINE
// ==========================================
document.getElementById('input-excel-file').addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (file) {
    document.getElementById('file-name-display').textContent = `${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
  }
});

document.getElementById('form-new-session').addEventListener('submit', async (e) => {
  e.preventDefault();
  const supplierName = document.getElementById('input-supplier').value.trim();
  const invoiceNo = document.getElementById('input-invoice').value.trim();
  const date = document.getElementById('input-date').value;
  const fileInput = document.getElementById('input-excel-file');

  if (!fileInput.files || fileInput.files.length === 0) {
    showToast('Please select a valid Excel file', 'error');
    return;
  }

  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = (evt) => {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

      if (jsonRows.length < 2) {
        showToast('Excel file contains no data rows', 'error');
        return;
      }

      state.excelRawData = jsonRows;
      state.excelHeaders = jsonRows[0].map(h => String(h).trim());
      state.excelRows = jsonRows.slice(1).filter(row => row.some(cell => String(cell).trim() !== ''));

      // Create new session object
      state.currentSession = {
        id: `session_${Date.now()}`,
        supplierName,
        invoiceNo,
        verificationDate: date,
        fileName: file.name,
        excelHeaders: state.excelHeaders,
        excelRawData: state.excelRawData,
        columnMapping: {},
        verificationFields: { mrp: true, qty: true, size: true, brand: false, design: false, nature: false },
        settings: { soundEnabled: true, vibrateEnabled: true },
        items: [],
        extraStock: [],
        scanHistory: []
      };

      renderColumnMappingUI();
      autoSuggestMapping();
      updateNavPill();
      switchView('mapping');
      showToast('Excel uploaded & parsed successfully', 'success');

    } catch (err) {
      console.error('Excel parse error:', err);
      showToast('Failed to parse Excel file. Ensure valid .xlsx, .xls, or .csv', 'error');
    }
  };

  reader.readAsArrayBuffer(file);
});

function renderColumnMappingUI() {
  document.getElementById('mapping-file-meta').textContent = 
    `File: ${state.currentSession.fileName} | Total Rows: ${state.excelRows.length} | Columns: ${state.excelHeaders.length}`;

  const renderGrid = (fields, containerId) => {
    const container = document.getElementById(containerId);
    container.innerHTML = fields.map(field => `
      <div class="mapping-item">
        <label class="mapping-label">
          <span>${field.label} ${field.required ? '<span class="required">*</span>' : ''}</span>
        </label>
        <select class="column-select" data-field="${field.key}" ${field.required ? 'required' : ''}>
          <option value="">-- Select Excel Header --</option>
          ${state.excelHeaders.map((hdr, idx) => `<option value="${idx}">${hdr}</option>`).join('')}
        </select>
      </div>
    `).join('');
  };

  renderGrid(COLUMN_FIELDS.core, 'core-mapping-grid');
  renderGrid(COLUMN_FIELDS.optional, 'optional-mapping-grid');

  // Attach change listener to update data preview & warnings
  document.querySelectorAll('.column-select').forEach(sel => {
    sel.addEventListener('change', () => {
      saveMappingFromUI();
      renderMappingPreviewAndWarnings();
    });
  });
}

function autoSuggestMapping() {
  const allFields = [...COLUMN_FIELDS.core, ...COLUMN_FIELDS.optional];
  const selects = document.querySelectorAll('.column-select');

  selects.forEach(select => {
    const fieldKey = select.getAttribute('data-field');
    const fieldDef = allFields.find(f => f.key === fieldKey);
    if (!fieldDef) return;

    let bestMatchIdx = -1;
    state.excelHeaders.forEach((hdr, idx) => {
      const cleanHdr = hdr.toLowerCase().replace(/[^a-z0-9_]/g, '');
      if (fieldDef.keywords.some(kw => cleanHdr.includes(kw))) {
        if (bestMatchIdx === -1) bestMatchIdx = idx;
      }
    });

    if (bestMatchIdx !== -1) {
      select.value = String(bestMatchIdx);
    }
  });

  saveMappingFromUI();
  renderMappingPreviewAndWarnings();
}

document.getElementById('btn-auto-map').addEventListener('click', autoSuggestMapping);

function saveMappingFromUI() {
  const mapping = {};
  document.querySelectorAll('.column-select').forEach(select => {
    const key = select.getAttribute('data-field');
    const val = select.value;
    if (val !== '') {
      mapping[key] = parseInt(val, 10);
    }
  });
  state.currentSession.columnMapping = mapping;
}

function renderMappingPreviewAndWarnings() {
  const mapping = state.currentSession.columnMapping;
  const eanIdx = mapping.ean;
  const qtyIdx = mapping.qty;

  const warnings = [];
  if (eanIdx === undefined) {
    warnings.push('EAN / Barcode column is not mapped.');
  }
  if (qtyIdx === undefined) {
    warnings.push('Expected Quantity column is not mapped.');
  }

  if (eanIdx !== undefined && qtyIdx !== undefined) {
    let emptyEanCount = 0;
    let invalidQtyCount = 0;
    const eanSet = new Set();
    const duplicateEans = new Set();

    state.excelRows.forEach((row) => {
      const eanVal = String(row[eanIdx] || '').trim();
      const qtyVal = parseFloat(row[qtyIdx]);

      if (!eanVal) emptyEanCount++;
      if (isNaN(qtyVal) || qtyVal < 0) invalidQtyCount++;

      if (eanVal) {
        if (eanSet.has(eanVal)) duplicateEans.add(eanVal);
        else eanSet.add(eanVal);
      }
    });

    if (emptyEanCount > 0) warnings.push(`${emptyEanCount} rows are missing EAN / Barcode values.`);
    if (invalidQtyCount > 0) warnings.push(`${invalidQtyCount} rows have invalid or empty quantity values.`);
    if (duplicateEans.size > 0) warnings.push(`${duplicateEans.size} duplicate EAN codes detected across rows/sizes. Safe duplicate handler will prompt during scan.`);
  }

  const warningsBox = document.getElementById('mapping-warnings-box');
  const warningsList = document.getElementById('mapping-warnings-list');

  if (warnings.length > 0) {
    warningsBox.classList.remove('hidden');
    warningsList.innerHTML = warnings.map(w => `<li>${w}</li>`).join('');
  } else {
    warningsBox.classList.add('hidden');
  }

  // Render Data Preview
  const theadRow = document.getElementById('preview-thead-row');
  const tbody = document.getElementById('preview-tbody');

  theadRow.innerHTML = state.excelHeaders.map((hdr, i) => {
    const mappedKey = Object.keys(mapping).find(k => mapping[k] === i);
    return `<th>${hdr} ${mappedKey ? `<br><small class="badge badge-info">${mappedKey.toUpperCase()}</small>` : ''}</th>`;
  }).join('');

  const sampleRows = state.excelRows.slice(0, 5);
  tbody.innerHTML = sampleRows.map(row => `
    <tr>
      ${state.excelHeaders.map((_, i) => `<td>${row[i] !== undefined ? row[i] : ''}</td>`).join('')}
    </tr>
  `).join('');
}

document.getElementById('btn-back-to-dashboard').addEventListener('click', () => switchView('dashboard'));
document.getElementById('btn-proceed-to-setup').addEventListener('click', () => {
  saveMappingFromUI();
  const mapping = state.currentSession.columnMapping;
  if (mapping.ean === undefined || mapping.qty === undefined) {
    showToast('Please map EAN / Barcode and Expected Quantity before proceeding', 'error');
    return;
  }
  buildSessionItemsFromMapping();
  switchView('setup');
});

function buildSessionItemsFromMapping() {
  const m = state.currentSession.columnMapping;
  
  state.currentSession.items = state.excelRows.map((row, idx) => {
    const rawEan = String(row[m.ean] || '').trim();
    const rawQty = parseFloat(row[m.qty]) || 0;
    const rawMrp = m.mrp !== undefined ? (parseFloat(row[m.mrp]) || 0) : 0;
    const rawSize = m.size !== undefined ? String(row[m.size] || '').trim() : '';

    return {
      rowIndex: idx + 1, // Excel 1-based data row index
      ean: rawEan,
      expectedQty: rawQty,
      scannedQty: 0,
      actualPhysicalQty: 0,
      expectedMrp: rawMrp,
      actualMrp: rawMrp,
      mrpStatus: 'Not Verified',
      expectedSize: rawSize,
      actualSize: rawSize,
      sizeStatus: 'Not Verified',
      rate: m.rate !== undefined ? (parseFloat(row[m.rate]) || 0) : 0,
      brand: m.brand !== undefined ? String(row[m.brand] || '').trim() : '',
      pattern: m.pattern !== undefined ? String(row[m.pattern] || '').trim() : '',
      section: m.section !== undefined ? String(row[m.section] || '').trim() : '',
      season: m.season !== undefined ? String(row[m.season] || '').trim() : '',
      design: m.design !== undefined ? String(row[m.design] || '').trim() : '',
      nature: m.nature !== undefined ? String(row[m.nature] || '').trim() : '',
      remarks: '',
      rawRowData: row
    };
  });

  DB.saveSession(state.currentSession);
}

// ==========================================
// 5. VERIFICATION SETUP ENGINE
// ==========================================
document.getElementById('btn-back-to-mapping').addEventListener('click', () => switchView('mapping'));
document.getElementById('btn-start-scanning').addEventListener('click', () => {
  // Save checkbox config
  state.currentSession.verificationFields = {
    mrp: document.getElementById('chk-verify-mrp').checked,
    qty: true,
    size: document.getElementById('chk-verify-size').checked,
    brand: document.getElementById('chk-verify-brand').checked,
    design: document.getElementById('chk-verify-design').checked,
    nature: document.getElementById('chk-verify-nature').checked
  };

  state.currentSession.settings = {
    soundEnabled: document.getElementById('chk-sound-enabled').checked,
    vibrateEnabled: document.getElementById('chk-vibrate-enabled').checked
  };

  DB.saveSession(state.currentSession);
  updateScanViewSummary();
  switchView('scan');
});

// ==========================================
// 6. SCAN & VERIFY OPERATIONAL MODULE
// ==========================================

// Barcode Scan Input (USB Scanner & Manual)
document.getElementById('form-scan-barcode').addEventListener('submit', (e) => {
  e.preventDefault();
  const input = document.getElementById('input-barcode-scan');
  const code = input.value.trim();
  if (code) {
    handleBarcodeScan(code);
    input.value = '';
    input.focus();
  }
});

function handleBarcodeScan(scannedCode) {
  if (!state.currentSession || !state.currentSession.items) return;

  const codeStr = String(scannedCode).trim();
  const matchingItems = state.currentSession.items.filter(item => 
    item.ean.toLowerCase() === codeStr.toLowerCase()
  );

  if (matchingItems.length === 0) {
    // Check extra stock items
    const extraMatch = state.currentSession.extraStock.find(item => item.ean.toLowerCase() === codeStr.toLowerCase());
    if (extraMatch) {
      displayProductCard(extraMatch, true);
      playBeep('success');
      showToast(`Extra Item Scanned: ${codeStr}`, 'info');
      return;
    }

    // Unlisted Product
    playBeep('error');
    if (confirm(`Item EAN '${codeStr}' not found in supplier Excel sheet!\n\nWould you like to add it as Extra / Unlisted Stock?`)) {
      openExtraStockModal(codeStr);
    }
    return;
  }

  if (matchingItems.length === 1) {
    // Single Match
    executeScanForItem(matchingItems[0]);
  } else {
    // Multiple Matches for same EAN (Duplicate Barcode across rows/sizes)
    playBeep('warning');
    openDuplicateEanModal(codeStr, matchingItems);
  }
}

function executeScanForItem(item) {
  item.scannedQty += 1;
  // Keep physical qty synced with scanned count unless user manually overrides
  if (item.actualPhysicalQty === (item.scannedQty - 1)) {
    item.actualPhysicalQty = item.scannedQty;
  }

  state.currentSession.scanHistory.push({
    timestamp: new Date().toISOString(),
    ean: item.ean,
    rowIndex: item.rowIndex,
    type: 'excel'
  });

  const isExcess = item.scannedQty > item.expectedQty;
  playBeep(isExcess ? 'warning' : 'success');

  state.activeScanItem = item;
  displayProductCard(item, false);
  updateScanViewSummary();
  DB.saveSession(state.currentSession);
}

function displayProductCard(item, isExtra = false) {
  document.getElementById('scan-empty-state').classList.add('hidden');
  document.getElementById('scan-product-content').classList.remove('hidden');

  // Title & Header
  document.getElementById('product-status-badge').className = `badge ${isExtra ? 'badge-danger' : 'badge-success'}`;
  document.getElementById('product-status-badge').textContent = isExtra ? 'UNLISTED EXTRA PRODUCT' : 'LISTED EXCEL ITEM';
  document.getElementById('product-title').textContent = `${item.brand || 'Item'} - ${item.nature || 'Garment'} (${item.design || 'Style N/A'})`;
  document.getElementById('product-ean').textContent = item.ean;
  document.getElementById('product-row-num').textContent = isExtra ? 'N/A' : `#${item.rowIndex}`;

  // Specs
  document.getElementById('val-brand').textContent = item.brand || '-';
  document.getElementById('val-nature').textContent = item.nature || '-';
  document.getElementById('val-design').textContent = item.design || '-';
  document.getElementById('val-pattern').textContent = item.pattern || '-';
  document.getElementById('val-section').textContent = item.section || '-';
  document.getElementById('val-season').textContent = item.season || '-';
  document.getElementById('val-rate').textContent = item.rate ? `₹${item.rate.toFixed(2)}` : '-';

  // Quantities
  const remaining = Math.max(0, item.expectedQty - item.scannedQty);
  const excess = Math.max(0, item.actualPhysicalQty - item.expectedQty);

  document.getElementById('qty-expected').textContent = item.expectedQty || 0;
  document.getElementById('qty-scanned').textContent = item.scannedQty;
  document.getElementById('qty-remaining').textContent = remaining;
  document.getElementById('qty-excess').textContent = excess;

  document.getElementById('input-physical-qty').value = item.actualPhysicalQty;

  // Qty Status Badge & Excess Banner
  const qtyStatusBadge = document.getElementById('qty-status-badge');
  const excessBanner = document.getElementById('excess-alert-banner');

  if (item.scannedQty > item.expectedQty) {
    qtyStatusBadge.className = 'qty-status-badge badge badge-danger';
    qtyStatusBadge.textContent = 'EXCESS STOCK';
    excessBanner.classList.remove('hidden');
  } else if (item.scannedQty === item.expectedQty) {
    qtyStatusBadge.className = 'qty-status-badge badge badge-success';
    qtyStatusBadge.textContent = 'MATCHED';
    excessBanner.classList.add('hidden');
  } else {
    qtyStatusBadge.className = 'qty-status-badge badge badge-warning';
    qtyStatusBadge.textContent = 'PENDING';
    excessBanner.classList.add('hidden');
  }

  // MRP Verification Box
  const boxMrp = document.getElementById('box-mrp-verify');
  const fieldsConfig = state.currentSession.verificationFields || {};
  if (fieldsConfig.mrp !== false) {
    boxMrp.classList.remove('hidden');
    document.getElementById('mrp-expected-val').textContent = `₹${(item.expectedMrp || 0).toFixed(2)}`;
    renderMrpVerifyButtons(item);
  } else {
    boxMrp.classList.add('hidden');
  }

  // Size Verification Box
  const boxSize = document.getElementById('box-size-verify');
  if (fieldsConfig.size !== false) {
    boxSize.classList.remove('hidden');
    document.getElementById('size-expected-val').textContent = item.expectedSize || 'Unspecified';
    renderSizeVerifyButtons(item);
  } else {
    boxSize.classList.add('hidden');
  }

  // Remarks
  const remarksInput = document.getElementById('input-item-remarks');
  remarksInput.value = item.remarks || '';
  remarksInput.onchange = () => {
    item.remarks = remarksInput.value;
    DB.saveSession(state.currentSession);
  };
}

function renderMrpVerifyButtons(item) {
  const btnMatch = document.getElementById('btn-mrp-match');
  const btnDiff = document.getElementById('btn-mrp-diff');
  const groupDiff = document.getElementById('group-mrp-diff');
  const inputActualMrp = document.getElementById('input-actual-mrp');

  btnMatch.className = `btn-toggle ${item.mrpStatus === 'Match' ? 'active-match' : ''}`;
  btnDiff.className = `btn-toggle ${item.mrpStatus === 'Mismatch' ? 'active-diff' : ''}`;

  if (item.mrpStatus === 'Mismatch') {
    groupDiff.classList.remove('hidden');
    inputActualMrp.value = item.actualMrp || '';
  } else {
    groupDiff.classList.add('hidden');
  }

  btnMatch.onclick = () => {
    item.mrpStatus = 'Match';
    item.actualMrp = item.expectedMrp;
    renderMrpVerifyButtons(item);
    updateScanViewSummary();
    DB.saveSession(state.currentSession);
  };

  btnDiff.onclick = () => {
    item.mrpStatus = 'Mismatch';
    renderMrpVerifyButtons(item);
    inputActualMrp.focus();
    updateScanViewSummary();
    DB.saveSession(state.currentSession);
  };

  inputActualMrp.onchange = () => {
    item.actualMrp = parseFloat(inputActualMrp.value) || 0;
    DB.saveSession(state.currentSession);
  };
}

function renderSizeVerifyButtons(item) {
  const btnMatch = document.getElementById('btn-size-match');
  const btnDiff = document.getElementById('btn-size-diff');
  const groupDiff = document.getElementById('group-size-diff');
  const inputActualSize = document.getElementById('input-actual-size');

  btnMatch.className = `btn-toggle ${item.sizeStatus === 'Match' ? 'active-match' : ''}`;
  btnDiff.className = `btn-toggle ${item.sizeStatus === 'Mismatch' ? 'active-diff' : ''}`;

  if (item.sizeStatus === 'Mismatch') {
    groupDiff.classList.remove('hidden');
    inputActualSize.value = item.actualSize || '';
  } else {
    groupDiff.classList.add('hidden');
  }

  btnMatch.onclick = () => {
    item.sizeStatus = 'Match';
    item.actualSize = item.expectedSize;
    renderSizeVerifyButtons(item);
    updateScanViewSummary();
    DB.saveSession(state.currentSession);
  };

  btnDiff.onclick = () => {
    item.sizeStatus = 'Mismatch';
    renderSizeVerifyButtons(item);
    inputActualSize.focus();
    updateScanViewSummary();
    DB.saveSession(state.currentSession);
  };

  inputActualSize.onchange = () => {
    item.actualSize = inputActualSize.value.trim();
    DB.saveSession(state.currentSession);
  };
}

// Quantity Actions (+1, Undo, Physical Qty input)
document.getElementById('btn-qty-add').addEventListener('click', () => {
  if (state.activeScanItem) {
    executeScanForItem(state.activeScanItem);
  }
});

document.getElementById('btn-qty-undo').addEventListener('click', () => {
  if (!state.currentSession || !state.currentSession.scanHistory.length) {
    showToast('No scan history to undo', 'info');
    return;
  }
  const lastScan = state.currentSession.scanHistory.pop();
  if (lastScan && lastScan.type === 'excel') {
    const item = state.currentSession.items.find(i => i.rowIndex === lastScan.rowIndex);
    if (item && item.scannedQty > 0) {
      item.scannedQty -= 1;
      if (item.actualPhysicalQty > item.scannedQty) {
        item.actualPhysicalQty = item.scannedQty;
      }
      displayProductCard(item, false);
      updateScanViewSummary();
      DB.saveSession(state.currentSession);
      showToast(`Scan undone for EAN ${item.ean}`, 'info');
    }
  }
});

document.getElementById('input-physical-qty').addEventListener('change', (e) => {
  if (state.activeScanItem) {
    const val = parseInt(e.target.value, 10);
    if (!isNaN(val) && val >= 0) {
      state.activeScanItem.actualPhysicalQty = val;
      displayProductCard(state.activeScanItem, false);
      updateScanViewSummary();
      DB.saveSession(state.currentSession);
    }
  }
});

function updateScanViewSummary() {
  if (!state.currentSession || !state.currentSession.items) return;

  const items = state.currentSession.items;
  const extras = state.currentSession.extraStock || [];

  let totalExpected = 0;
  let totalScanned = 0;
  let totalPending = 0;
  let mrpDiff = 0;
  let sizeDiff = 0;

  items.forEach(item => {
    totalExpected += (item.expectedQty || 0);
    totalScanned += item.scannedQty;
    totalPending += Math.max(0, item.expectedQty - item.scannedQty);

    if (item.mrpStatus === 'Mismatch') mrpDiff++;
    if (item.sizeStatus === 'Mismatch') sizeDiff++;
  });

  document.getElementById('stat-total-expected').textContent = totalExpected;
  document.getElementById('stat-total-scanned').textContent = totalScanned;
  document.getElementById('stat-total-pending').textContent = totalPending;
  document.getElementById('stat-mrp-diff').textContent = mrpDiff;
  document.getElementById('stat-size-diff').textContent = sizeDiff;
  document.getElementById('stat-extra-items').textContent = extras.length;
}

// ==========================================
// 7. DUPLICATE EAN & EXTRA STOCK MODALS
// ==========================================
function openDuplicateEanModal(ean, items) {
  const modal = document.getElementById('modal-duplicate-ean');
  const container = document.getElementById('duplicate-items-list');

  container.innerHTML = items.map(item => `
    <button type="button" class="duplicate-item-btn" data-row="${item.rowIndex}">
      <div><strong>Row #${item.rowIndex}</strong> &bull; Size: <strong>${item.expectedSize || 'N/A'}</strong> &bull; MRP: <strong>₹${item.expectedMrp}</strong></div>
      <div style="font-size:0.85rem; color: var(--text-muted); margin-top:0.2rem;">
        Pattern: ${item.pattern || '-'} | Design: ${item.design || '-'} | Expected: ${item.expectedQty} | Scanned: ${item.scannedQty}
      </div>
    </button>
  `).join('');

  modal.classList.remove('hidden');

  container.querySelectorAll('.duplicate-item-btn').forEach(btn => {
    btn.onclick = () => {
      const rowIdx = parseInt(btn.getAttribute('data-row'), 10);
      const selectedItem = items.find(i => i.rowIndex === rowIdx);
      modal.classList.add('hidden');
      if (selectedItem) executeScanForItem(selectedItem);
    };
  });
}

document.getElementById('btn-close-duplicate-modal').addEventListener('click', () => {
  document.getElementById('modal-duplicate-ean').classList.add('hidden');
});

function openExtraStockModal(defaultEan = '') {
  const modal = document.getElementById('modal-extra-stock');
  document.getElementById('form-extra-stock').reset();
  document.getElementById('extra-ean').value = defaultEan;
  modal.classList.remove('hidden');
}

document.getElementById('btn-manual-add-extra').addEventListener('click', () => openExtraStockModal(''));
document.getElementById('btn-cancel-extra-modal').addEventListener('click', () => {
  document.getElementById('modal-extra-stock').classList.add('hidden');
});

document.getElementById('form-extra-stock').addEventListener('submit', (e) => {
  e.preventDefault();
  const ean = document.getElementById('extra-ean').value.trim();
  const qty = parseInt(document.getElementById('extra-qty').value, 10) || 1;
  const mrp = parseFloat(document.getElementById('extra-mrp').value) || 0;
  const size = document.getElementById('extra-size').value.trim();
  const brand = document.getElementById('extra-brand').value.trim();
  const design = document.getElementById('extra-design').value.trim();
  const nature = document.getElementById('extra-nature').value.trim();
  const remarks = document.getElementById('extra-remarks').value.trim();

  const extraItem = {
    id: `extra_${Date.now()}`,
    ean,
    expectedQty: 0,
    scannedQty: qty,
    actualPhysicalQty: qty,
    expectedMrp: mrp,
    actualMrp: mrp,
    mrpStatus: 'Unlisted Extra Item',
    expectedSize: size,
    actualSize: size,
    sizeStatus: 'Unlisted Extra Item',
    brand,
    design,
    nature,
    remarks
  };

  state.currentSession.extraStock.push(extraItem);
  document.getElementById('modal-extra-stock').classList.add('hidden');
  displayProductCard(extraItem, true);
  updateScanViewSummary();
  DB.saveSession(state.currentSession);
  showToast('Unlisted Extra Stock Item added', 'success');
});

// ==========================================
// 8. CAMERA BARCODE SCANNER MODULE (ZXing)
// ==========================================
document.getElementById('btn-open-camera').addEventListener('click', async () => {
  const modal = document.getElementById('modal-camera-scanner');
  modal.classList.remove('hidden');
  startCameraScanner();
});

document.getElementById('btn-close-camera').addEventListener('click', () => {
  stopCameraScanner();
  document.getElementById('modal-camera-scanner').classList.add('hidden');
});

async function startCameraScanner() {
  if (state.isScanningCamera) return;

  try {
    if (!window.ZXing) {
      showToast('ZXing camera library loading...', 'info');
      return;
    }

    const codeReader = new ZXing.BrowserMultiFormatReader();
    state.zxingReader = codeReader;

    const videoInputDevices = await codeReader.listVideoInputDevices();
    const select = document.getElementById('select-camera-device');
    select.innerHTML = videoInputDevices.map((device, idx) => 
      `<option value="${device.deviceId}">${device.label || `Camera ${idx + 1}`}</option>`
    ).join('');

    const selectedDeviceId = videoInputDevices.length > 0 ? videoInputDevices[0].deviceId : undefined;

    let lastScanTime = 0;
    let lastScanCode = '';

    codeReader.decodeFromVideoDevice(selectedDeviceId, 'camera-video', (result, err) => {
      if (result) {
        const code = result.getText();
        const now = Date.now();

        // Debounce continuous camera scanning (500ms debounce window)
        if (code !== lastScanCode || (now - lastScanTime) > 800) {
          lastScanCode = code;
          lastScanTime = now;
          handleBarcodeScan(code);
        }
      }
    });

    state.isScanningCamera = true;

  } catch (err) {
    console.error('Camera scanner error:', err);
    showToast('Could not access camera scanner', 'error');
  }
}

function stopCameraScanner() {
  if (state.zxingReader) {
    state.zxingReader.reset();
    state.zxingReader = null;
  }
  state.isScanningCamera = false;
}

// ==========================================
// 9. INVENTORY OVERVIEW & EXCEPTION SCREEN
// ==========================================
function renderInventoryOverview() {
  if (!state.currentSession) return;

  const items = state.currentSession.items || [];
  const extras = state.currentSession.extraStock || [];
  const filterVal = document.getElementById('select-overview-filter').value;
  const searchVal = document.getElementById('input-overview-search').value.toLowerCase().trim();

  // Aggregate stats
  let totalExpected = 0, totalScanned = 0, totalPending = 0, mrpDiff = 0, sizeDiff = 0;
  items.forEach(item => {
    totalExpected += (item.expectedQty || 0);
    totalScanned += item.scannedQty;
    totalPending += Math.max(0, item.expectedQty - item.scannedQty);
    if (item.mrpStatus === 'Mismatch') mrpDiff++;
    if (item.sizeStatus === 'Mismatch') sizeDiff++;
  });

  document.getElementById('ov-stat-total-items').textContent = items.length;
  document.getElementById('ov-stat-expected-units').textContent = totalExpected;
  document.getElementById('ov-stat-scanned-units').textContent = totalScanned;
  document.getElementById('ov-stat-pending-units').textContent = totalPending;
  document.getElementById('ov-stat-mrp-diff').textContent = mrpDiff;
  document.getElementById('ov-stat-size-diff').textContent = sizeDiff;
  document.getElementById('ov-stat-extra-items').textContent = extras.length;

  // Filter List
  let combined = [
    ...items.map(i => ({ ...i, isExtra: false })),
    ...extras.map(e => ({ ...e, isExtra: true }))
  ];

  if (searchVal) {
    combined = combined.filter(i => 
      (i.ean && i.ean.toLowerCase().includes(searchVal)) ||
      (i.brand && i.brand.toLowerCase().includes(searchVal)) ||
      (i.design && i.design.toLowerCase().includes(searchVal)) ||
      (i.nature && i.nature.toLowerCase().includes(searchVal))
    );
  }

  if (filterVal === 'pending') combined = combined.filter(i => !i.isExtra && i.scannedQty < i.expectedQty);
  if (filterVal === 'verified') combined = combined.filter(i => i.scannedQty > 0);
  if (filterVal === 'mrp_diff') combined = combined.filter(i => i.mrpStatus === 'Mismatch');
  if (filterVal === 'size_diff') combined = combined.filter(i => i.sizeStatus === 'Mismatch');
  if (filterVal === 'excess') combined = combined.filter(i => i.actualPhysicalQty > i.expectedQty);
  if (filterVal === 'extra') combined = combined.filter(i => i.isExtra);

  const tbody = document.getElementById('tbody-inventory-overview');
  tbody.innerHTML = combined.map(item => {
    const qtyDiff = item.actualPhysicalQty - (item.expectedQty || 0);
    let statusBadge = '<span class="badge badge-warning">Pending</span>';

    if (item.isExtra) {
      statusBadge = '<span class="badge badge-danger">Unlisted Extra</span>';
    } else if (item.actualPhysicalQty > item.expectedQty) {
      statusBadge = '<span class="badge badge-danger">Excess Stock</span>';
    } else if (item.mrpStatus === 'Mismatch' || item.sizeStatus === 'Mismatch') {
      statusBadge = '<span class="badge badge-warning">Mismatch Alert</span>';
    } else if (item.actualPhysicalQty === item.expectedQty && item.expectedQty > 0) {
      statusBadge = '<span class="badge badge-success">Verified OK</span>';
    }

    return `
      <tr>
        <td>${item.isExtra ? 'EXTRA' : `#${item.rowIndex}`}</td>
        <td><strong>${item.ean}</strong></td>
        <td>${item.brand || 'Item'} - ${item.nature || ''} (${item.design || ''})</td>
        <td>Exp: ${item.expectedSize || '-'} | Act: <strong>${item.actualSize || '-'}</strong></td>
        <td>Exp: ₹${item.expectedMrp || 0} | Act: <strong>₹${item.actualMrp || 0}</strong></td>
        <td>${item.expectedQty || 0}</td>
        <td><strong>${item.scannedQty}</strong></td>
        <td><strong>${item.actualPhysicalQty}</strong></td>
        <td>${statusBadge} ${qtyDiff !== 0 ? `<br><small>Diff: ${qtyDiff > 0 ? '+' : ''}${qtyDiff}</small>` : ''}</td>
        <td>
          <button class="btn btn-sm btn-outline btn-verify-row" data-row="${item.rowIndex}" data-extra="${item.isExtra ? item.id : ''}">
            Verify
          </button>
        </td>
      </tr>
    `;
  }).join('');

  tbody.querySelectorAll('.btn-verify-row').forEach(btn => {
    btn.onclick = () => {
      const extraId = btn.getAttribute('data-extra');
      const rowNum = parseInt(btn.getAttribute('data-row'), 10);
      if (extraId) {
        const extraItem = state.currentSession.extraStock.find(e => e.id === extraId);
        if (extraItem) displayProductCard(extraItem, true);
      } else {
        const item = state.currentSession.items.find(i => i.rowIndex === rowNum);
        if (item) displayProductCard(item, false);
      }
      switchView('scan');
    };
  });
}

document.getElementById('input-overview-search').addEventListener('input', renderInventoryOverview);
document.getElementById('select-overview-filter').addEventListener('change', renderInventoryOverview);

// ==========================================
// 10. REVIEW & EXPORT ENGINE
// ==========================================
function renderExportReview() {
  if (!state.currentSession) return;

  const s = state.currentSession;
  const items = s.items || [];
  const extras = s.extraStock || [];

  let totalExpected = 0, totalPhysical = 0, totalPending = 0, mrpDiff = 0, sizeDiff = 0;
  items.forEach(item => {
    totalExpected += (item.expectedQty || 0);
    totalPhysical += item.actualPhysicalQty;
    totalPending += Math.max(0, item.expectedQty - item.scannedQty);
    if (item.mrpStatus === 'Mismatch') mrpDiff++;
    if (item.sizeStatus === 'Mismatch') sizeDiff++;
  });

  document.getElementById('exp-val-supplier').textContent = s.supplierName;
  document.getElementById('exp-val-invoice').textContent = s.invoiceNo;
  document.getElementById('exp-val-date').textContent = s.verificationDate;
  document.getElementById('exp-val-lines').textContent = items.length;
  document.getElementById('exp-val-expected').textContent = totalExpected;
  document.getElementById('exp-val-physical').textContent = totalPhysical;
  document.getElementById('exp-val-pending').textContent = totalPending;
  document.getElementById('exp-val-mrp-mismatch').textContent = mrpDiff;
  document.getElementById('exp-val-size-mismatch').textContent = sizeDiff;
  document.getElementById('exp-val-extra').textContent = extras.length;
}

// Download Verified Excel (.xlsx)
document.getElementById('btn-download-excel').addEventListener('click', () => {
  if (!state.currentSession || !state.currentSession.excelRawData) {
    showToast('No active session data to export', 'error');
    return;
  }

  try {
    const rawData = state.currentSession.excelRawData; // 2D Array
    const originalHeaders = rawData[0].map(h => String(h).trim());

    // 10 Additional Verification Columns
    const verificationHeaders = [
      'Scanned Qty',
      'Actual Physical Qty',
      'Qty Difference',
      'Actual MRP',
      'MRP Verification Status',
      'Actual Size',
      'Size Verification Status',
      'Actual Brand / Item',
      'Verification Status',
      'Remarks'
    ];

    const newHeaderRow = [...originalHeaders, ...verificationHeaders];
    const newRows = [newHeaderRow];

    const items = state.currentSession.items || [];
    const mapping = state.currentSession.columnMapping;

    // Process Original Listed Excel Rows
    state.excelRows.forEach((rawRow, idx) => {
      const item = items.find(i => i.rowIndex === (idx + 1)) || {};

      const qtyDiff = (item.actualPhysicalQty || 0) - (item.expectedQty || 0);
      let statusStr = 'Pending';
      if (item.actualPhysicalQty > item.expectedQty) {
        statusStr = 'Excess Stock';
      } else if (item.mrpStatus === 'Mismatch' || item.sizeStatus === 'Mismatch') {
        statusStr = 'Verified - Mismatch Alert';
      } else if (item.actualPhysicalQty === item.expectedQty && item.expectedQty > 0) {
        statusStr = 'Verified - OK';
      }

      const verCols = [
        item.scannedQty || 0,
        item.actualPhysicalQty || 0,
        qtyDiff,
        item.actualMrp !== undefined ? item.actualMrp : (item.expectedMrp || ''),
        item.mrpStatus || 'Not Verified',
        item.actualSize || (item.expectedSize || ''),
        item.sizeStatus || 'Not Verified',
        item.brand || '',
        statusStr,
        item.remarks || ''
      ];

      newRows.push([...rawRow, ...verCols]);
    });

    // Append Extra / Unlisted Stock Rows at bottom
    const extras = state.currentSession.extraStock || [];
    extras.forEach(extra => {
      const extraRow = new Array(originalHeaders.length).fill('');
      
      // Place mapped values into correct column positions
      if (mapping.ean !== undefined) extraRow[mapping.ean] = extra.ean;
      if (mapping.qty !== undefined) extraRow[mapping.qty] = 0;
      if (mapping.mrp !== undefined) extraRow[mapping.mrp] = extra.actualMrp;
      if (mapping.size !== undefined) extraRow[mapping.size] = extra.actualSize;
      if (mapping.brand !== undefined) extraRow[mapping.brand] = extra.brand;
      if (mapping.design !== undefined) extraRow[mapping.design] = extra.design;
      if (mapping.nature !== undefined) extraRow[mapping.nature] = extra.nature;

      const verCols = [
        extra.scannedQty || extra.actualPhysicalQty,
        extra.actualPhysicalQty,
        extra.actualPhysicalQty, // Qty Diff
        extra.actualMrp,
        'Unlisted Extra Item',
        extra.actualSize,
        'Unlisted Extra Item',
        extra.brand || '',
        'Unlisted Extra Item',
        extra.remarks || ''
      ];

      newRows.push([...extraRow, ...verCols]);
    });

    // Create Worksheet & Download
    const ws = XLSX.utils.aoa_to_sheet(newRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Verified Inventory');

    const cleanSupplier = state.currentSession.supplierName.replace(/[^a-zA-Z0-9]/g, '_');
    const cleanInvoice = state.currentSession.invoiceNo.replace(/[^a-zA-Z0-9]/g, '_');
    const fileName = `Verified_Inventory_${cleanSupplier}_${cleanInvoice}.xlsx`;

    XLSX.writeFile(wb, fileName);
    showToast('Verified Excel report downloaded successfully!', 'success');

  } catch (err) {
    console.error('Excel Export Error:', err);
    showToast('Failed to generate Excel export file', 'error');
  }
});

// JSON Session Backup Export & Restore
document.getElementById('btn-download-json-backup').addEventListener('click', () => {
  if (!state.currentSession) return;
  const jsonStr = JSON.stringify(state.currentSession, null, 2);
  const blob = new Blob([jsonStr], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Backup_${state.currentSession.supplierName}_${state.currentSession.invoiceNo}.json`;
  a.click();
  URL.revokeObjectURL(url);
  showToast('JSON Session backup downloaded', 'success');
});

document.getElementById('input-json-backup').addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = async (evt) => {
    try {
      const sessionData = JSON.parse(evt.target.result);
      if (!sessionData.id || !sessionData.supplierName) {
        showToast('Invalid JSON session backup file', 'error');
        return;
      }

      state.currentSession = sessionData;
      await DB.saveSession(sessionData);
      updateNavPill();
      updateScanViewSummary();
      switchView('scan');
      showToast('Session backup restored successfully!', 'success');
    } catch (err) {
      showToast('Failed to parse backup JSON file', 'error');
    }
  };
  reader.readAsText(file);
});

// Saved Sessions Manager
async function renderSavedSessionsList() {
  const sessions = await DB.getAllSessions();
  const container = document.getElementById('saved-sessions-list');

  if (sessions.length === 0) {
    container.innerHTML = `<div class="empty-state"><p>No saved sessions found.</p></div>`;
    return;
  }

  container.innerHTML = sessions.map(s => `
    <div class="session-card">
      <div>
        <div class="session-title-text">${s.supplierName}</div>
        <div class="session-meta-text">Inv: ${s.invoiceNo} &bull; ${s.verificationDate}</div>
      </div>
      <div class="session-actions">
        <button class="btn btn-sm btn-primary btn-resume-session" data-id="${s.id}">Resume</button>
        <button class="btn btn-sm btn-outline btn-delete-session" data-id="${s.id}">Delete</button>
      </div>
    </div>
  `).join('');

  container.querySelectorAll('.btn-resume-session').forEach(btn => {
    btn.onclick = async () => {
      const id = btn.getAttribute('data-id');
      const session = await DB.getSession(id);
      if (session) {
        state.currentSession = session;
        state.excelRawData = session.excelRawData;
        state.excelHeaders = session.excelHeaders || [];
        state.excelRows = session.excelRawData ? session.excelRawData.slice(1) : [];
        updateNavPill();
        updateScanViewSummary();
        switchView('scan');
        showToast('Session resumed', 'success');
      }
    };
  });

  container.querySelectorAll('.btn-delete-session').forEach(btn => {
    btn.onclick = async () => {
      const id = btn.getAttribute('data-id');
      if (confirm('Are you sure you want to delete this session?')) {
        await DB.deleteSession(id);
        renderSavedSessionsList();
        showToast('Session deleted', 'info');
      }
    };
  });
}

document.getElementById('btn-refresh-sessions').addEventListener('click', renderSavedSessionsList);

// App Navigation Clicks
document.querySelectorAll('.nav-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    const viewId = btn.getAttribute('data-view');
    switchView(viewId);
  });
});

// Init Application
window.addEventListener('DOMContentLoaded', () => {
  // Set default date input to today
  const dateInput = document.getElementById('input-date');
  if (dateInput) dateInput.value = new Date().toISOString().split('T')[0];

  renderSavedSessionsList();
});
