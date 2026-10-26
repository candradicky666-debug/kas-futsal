// Buku Kas Futsal - Interactive Pro Spreadsheet Web App (Table & Card View)
// Integrated with Google Apps Script & IndexedDB

const GSCRIPT_URL = 'https://script.google.com/macros/s/AKfycbz6mx3yARb6jFM1omx9sLjIC4rI8EyIZzPzbMAjNbaGHKPnIauOjJ_okx8BwcFEdd1LJA/exec';
const STORAGE_KEY = 'KAS_FUTSAL_TRANSACTIONS_V1';
const AUTO_BALANCE_KEY = 'KAS_FUTSAL_AUTO_BALANCE_V1';
const VIEW_PREF_KEY = 'KAS_FUTSAL_VIEW_PREF_V1';
const SNAPSHOT_KEY = 'KAS_FUTSAL_BACKUP_SNAPSHOT_V1';

const CATEGORIES = [
  'Iuran Kas',
  'Sewa Lapangan',
  'Konsumsi',
  'Perlengkapan',
  'Turnamen',
  'Uang Sparing',
  'Kas Masuk Lain',
  'Lain-lain'
];

// Initial starter data template
const DEFAULT_TRANSACTIONS = [
  { id: 1, date: '2026-10-01', desc: 'Saldo Awal Kas Tim', category: 'Kas Masuk Lain', income: 500000, expense: 0, balance: 500000 },
  { id: 2, date: '2026-10-04', desc: 'Iuran Kas Mingguan Pemain (Pekan 1)', category: 'Iuran Kas', income: 300000, expense: 0, balance: 800000 },
  { id: 3, date: '2026-10-04', desc: 'Sewa Lapangan Futsal (2 Jam)', category: 'Sewa Lapangan', income: 0, expense: 350000, balance: 450000 },
  { id: 4, date: '2026-10-04', desc: 'Air Mineral 1 Karton + Es Batu', category: 'Konsumsi', income: 0, expense: 45000, balance: 405000 },
  { id: 5, date: '2026-10-11', desc: 'Iuran Kas Mingguan Pemain (Pekan 2)', category: 'Iuran Kas', income: 280000, expense: 0, balance: 685000 },
  { id: 6, date: '2026-10-11', desc: 'Sewa Lapangan Futsal (2 Jam)', category: 'Sewa Lapangan', income: 0, expense: 350000, balance: 335000 },
  { id: 7, date: '2026-10-15', desc: 'Beli Bola Futsal Specs (1 pcs)', category: 'Perlengkapan', income: 0, expense: 240000, balance: 95000 },
  { id: 8, date: '2026-10-18', desc: 'Iuran Kas Mingguan Pemain (Pekan 3)', category: 'Iuran Kas', income: 300000, expense: 0, balance: 395000 },
  { id: 9, date: '2026-10-18', desc: 'Sewa Lapangan Futsal (2 Jam)', category: 'Sewa Lapangan', income: 0, expense: 350000, balance: 45000 },
  { id: 10, date: '2026-10-18', desc: 'Uang Pendaftaran Turnamen Antar Tim', category: 'Turnamen', income: 0, expense: 200000, balance: -155000 },
  { id: 11, date: '2026-10-25', desc: 'Hadiah Juara 3 Turnamen Futsal', category: 'Kas Masuk Lain', income: 500000, expense: 0, balance: 345000 },
  { id: 12, date: '2026-10-25', desc: 'Iuran Kas Mingguan Pemain (Pekan 4)', category: 'Iuran Kas', income: 320000, expense: 0, balance: 665000 }
];

let transactions = [];
let isAutoBalance = true;
let currentView = 'cards';
let searchQuery = '';
let isSyncing = false;

// DOM Elements
const mobileCardsSection = document.getElementById('mobileCardsSection');
const spreadsheetTableSection = document.getElementById('spreadsheetTableSection');
const tableBody = document.getElementById('tableBody');

const tabViewCards = document.getElementById('tabViewCards');
const tabViewTable = document.getElementById('tabViewTable');
const dockBtnSwitchView = document.getElementById('dockBtnSwitchView');
const dockSwitchIcon = document.getElementById('dockSwitchIcon');
const dockSwitchLabel = document.getElementById('dockSwitchLabel');

const kpiIncome = document.getElementById('kpiIncome');
const kpiExpense = document.getElementById('kpiExpense');
const kpiBalance = document.getElementById('kpiBalance');
const kpiStatusText = document.getElementById('kpiStatusText');

const totIncomeFoot = document.getElementById('totIncomeFoot');
const totExpenseFoot = document.getElementById('totExpenseFoot');
const totBalanceFoot = document.getElementById('totBalanceFoot');
const rowCountLabel = document.getElementById('rowCountLabel');

const searchInput = document.getElementById('searchInput');
const btnAddRow = document.getElementById('btnAddRow');
const dockBtnAdd = document.getElementById('dockBtnAdd');
const btnToggleAutoBalance = document.getElementById('btnToggleAutoBalance');
const autoBalanceIcon = document.getElementById('autoBalanceIcon');
const autoBalanceLabel = document.getElementById('autoBalanceLabel');

// Sync elements
const syncStatusDot = document.getElementById('syncStatusDot');
const syncStatusText = document.getElementById('syncStatusText');
const btnTriggerSync = document.getElementById('btnTriggerSync');
const btnSyncGSheet = document.getElementById('btnSyncGSheet');

// Modal Elements
const iosModalBackdrop = document.getElementById('iosModalBackdrop');
const btnOpenIosGuide = document.getElementById('btnOpenIosGuide');
const dockBtnIos = document.getElementById('dockBtnIos');
const btnCloseIosModal = document.getElementById('btnCloseIosModal');

const exportModalBackdrop = document.getElementById('exportModalBackdrop');
const btnExportMenu = document.getElementById('btnExportMenu');
const dockBtnExport = document.getElementById('dockBtnExport');
const btnCloseExportModal = document.getElementById('btnCloseExportModal');
const btnExportCsv = document.getElementById('btnExportCsv');
const btnExportJson = document.getElementById('btnExportJson');
const btnRestoreSnapshot = document.getElementById('btnRestoreSnapshot');
const btnResetToTemplate = document.getElementById('btnResetToTemplate');

const appToast = document.getElementById('appToast');
const toastMessage = document.getElementById('toastMessage');

// ==============================================================================
// 1. DATA PERSISTENCE & STORAGE (ANTI-HAPUS)
// ==============================================================================

async function requestPersistentStorage() {
  if (navigator.storage && navigator.storage.persist) {
    const isPersisted = await navigator.storage.persist();
    console.log('Persistent Storage active:', isPersisted);
  }
}

function openDB() {
  return new Promise((resolve) => {
    if (!window.indexedDB) {
      resolve(null);
      return;
    }
    const request = indexedDB.open('KasFutsalAppDB', 1);
    request.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains('transactions')) {
        db.createObjectStore('transactions', { keyPath: 'store_id' });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => resolve(null);
  });
}

async function saveToIndexedDB(data) {
  try {
    const db = await openDB();
    if (!db) return;
    const tx = db.transaction('transactions', 'readwrite');
    const store = tx.objectStore('transactions');
    store.put({ store_id: 'current_tx', items: data, updatedAt: new Date().toISOString() });
  } catch (err) {
    console.warn('IndexedDB write error:', err);
  }
}

async function loadFromIndexedDB() {
  try {
    const db = await openDB();
    if (!db) return null;
    return new Promise((resolve) => {
      const tx = db.transaction('transactions', 'readonly');
      const store = tx.objectStore('transactions');
      const req = store.get('current_tx');
      req.onsuccess = () => {
        if (req.result && Array.isArray(req.result.items)) {
          resolve(req.result.items);
        } else {
          resolve(null);
        }
      };
      req.onerror = () => resolve(null);
    });
  } catch (err) {
    return null;
  }
}

// Format number with commas (e.g. 350000 -> "350,000")
function formatCommaNumber(val) {
  if (val === null || val === undefined || isNaN(val) || val === '') return '0';
  const num = Math.round(Number(val));
  const isNeg = num < 0;
  const abs = Math.abs(num).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return isNeg ? `-${abs}` : abs;
}

// Format full currency string with Rp (e.g. 350000 -> "Rp 350,000")
function formatRupiah(num) {
  if (isNaN(num) || num === null || num === '') return 'Rp 0';
  const prefix = num < 0 ? '-Rp ' : 'Rp ';
  const absNum = Math.abs(Math.round(num));
  return prefix + absNum.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

// Clean number input string to float
function parseNumber(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  const cleaned = val.toString().replace(/[^0-9.-]/g, '');
  const parsed = parseFloat(cleaned);
  return isNaN(parsed) ? 0 : parsed;
}

function showToast(msg, icon = '✅') {
  toastMessage.textContent = msg;
  document.getElementById('toastIcon').textContent = icon;
  appToast.classList.add('show');
  setTimeout(() => {
    appToast.classList.remove('show');
  }, 2200);
}

// Load data
async function loadLocalData() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      transactions = JSON.parse(raw);
    } else {
      const idbData = await loadFromIndexedDB();
      if (idbData && idbData.length > 0) {
        transactions = idbData;
      } else {
        transactions = JSON.parse(JSON.stringify(DEFAULT_TRANSACTIONS));
      }
      saveLocalData(false);
    }

    const autoPref = localStorage.getItem(AUTO_BALANCE_KEY);
    if (autoPref !== null) {
      isAutoBalance = autoPref === 'true';
    }

    const viewPref = localStorage.getItem(VIEW_PREF_KEY);
    if (viewPref) {
      currentView = viewPref;
    } else {
      currentView = window.innerWidth <= 768 ? 'cards' : 'table';
    }
  } catch (e) {
    console.error('Error loading data:', e);
    transactions = JSON.parse(JSON.stringify(DEFAULT_TRANSACTIONS));
  }
  updateAutoBalanceUI();
  updateViewUI();
}

function saveLocalData(notify = true) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(transactions));
    localStorage.setItem(AUTO_BALANCE_KEY, isAutoBalance);

    if (transactions.length > 0) {
      localStorage.setItem(SNAPSHOT_KEY, JSON.stringify({
        timestamp: new Date().toISOString(),
        items: transactions
      }));
    }

    saveToIndexedDB(transactions);

    syncStatusDot.style.background = '#10B981';
    syncStatusText.textContent = 'Data tersimpan di HP (' + new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) + ')';

    if (notify) {
      showToast('Perubahan tersimpan otomatis');
    }
  } catch (e) {
    console.error('Error saving data:', e);
  }
}

// Google Sheets Sync
async function syncWithGoogleSheets(pushData = true) {
  if (isSyncing) return;
  isSyncing = true;
  syncStatusDot.style.background = '#F59E0B';
  syncStatusText.textContent = 'Menyinkronkan ke Google Spreadsheet...';

  try {
    if (pushData) {
      await fetch(GSCRIPT_URL, {
        method: 'POST',
        mode: 'no-cors',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8'
        },
        body: JSON.stringify({ transactions: transactions })
      });
      syncStatusDot.style.background = '#10B981';
      syncStatusText.textContent = 'Tersinkron ke Spreadsheet (' + new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) + ')';
      showToast('Data tersinkron ke Google Sheets', '☁️');
    } else {
      const res = await fetch(GSCRIPT_URL);
      if (res.ok) {
        const json = await res.json();
        if (json.status === 'success' && Array.isArray(json.data) && json.data.length > 0) {
          transactions = json.data;
          saveLocalData(false);
          renderAll();
          syncStatusDot.style.background = '#10B981';
          syncStatusText.textContent = 'Data terbaru dari Spreadsheet dimuat';
          showToast('Data terbaru dari Spreadsheet dimuat', '📥');
        }
      }
    }
  } catch (err) {
    console.warn('Sync issue:', err);
    syncStatusDot.style.background = '#10B981';
    syncStatusText.textContent = 'Tersimpan aman di HP (Offline Mode)';
  } finally {
    isSyncing = false;
  }
}

function recalculateBalances() {
  if (!isAutoBalance) return;
  let running = 0;
  transactions.forEach((tx) => {
    const inc = parseNumber(tx.income);
    const exp = parseNumber(tx.expense);
    running += (inc - exp);
    tx.balance = running;
  });
}

function updateTotals() {
  let totalIncome = 0;
  let totalExpense = 0;

  transactions.forEach((tx) => {
    totalIncome += parseNumber(tx.income);
    totalExpense += parseNumber(tx.expense);
  });

  const totalBalance = totalIncome - totalExpense;

  kpiIncome.textContent = formatRupiah(totalIncome);
  kpiExpense.textContent = formatRupiah(totalExpense);
  kpiBalance.textContent = formatRupiah(totalBalance);

  if (totalBalance >= 0) {
    kpiBalance.style.color = 'var(--accent-balance)';
    kpiStatusText.textContent = 'Saldo kas tim tersedia dan aman';
  } else {
    kpiBalance.style.color = 'var(--accent-expense)';
    kpiStatusText.textContent = 'Perhatian: Kas tim sedang defisit!';
  }

  totIncomeFoot.textContent = formatRupiah(totalIncome);
  totExpenseFoot.textContent = formatRupiah(totalExpense);
  totBalanceFoot.textContent = formatRupiah(totalBalance);

  rowCountLabel.textContent = `${transactions.length} baris transaksi`;
}

function getFilteredTransactions() {
  return transactions.filter((tx) => {
    if (!searchQuery) return true;
    const q = searchQuery.toLowerCase();
    return (
      (tx.desc && tx.desc.toLowerCase().includes(q)) ||
      (tx.category && tx.category.toLowerCase().includes(q)) ||
      (tx.date && tx.date.includes(q))
    );
  });
}

// ==============================================================================
// 2. RENDERING PRO: MOBILE CARDS & SPREADSHEET TABLE DENGAN BADGE RP
// ==============================================================================

function renderCards() {
  mobileCardsSection.innerHTML = '';
  const filtered = getFilteredTransactions();

  if (filtered.length === 0) {
    mobileCardsSection.innerHTML = `
      <div style="background: #FFFFFF; border-radius: 14px; padding: 32px 16px; text-align: center; color: var(--text-muted); border: 1px solid var(--border);">
        <div style="font-size: 32px; margin-bottom: 8px;">🔍</div>
        <p style="font-size: 0.92rem; font-weight: 700;">Tidak ada transaksi yang cocok.</p>
        <p style="font-size: 0.78rem; margin-top: 4px;">Ketuk "+ Tambah Baris" untuk membuat catatan kas baru.</p>
      </div>
    `;
    return;
  }

  filtered.forEach((tx, index) => {
    const card = document.createElement('div');
    card.className = 'mobile-tx-card';

    const optionsHtml = CATEGORIES.map((c) => {
      const sel = tx.category === c ? 'selected' : '';
      return `<option value="${c}" ${sel}>${c}</option>`;
    }).join('');

    card.innerHTML = `
      <div class="card-top-row">
        <div class="card-badge-row">
          <span class="card-no-badge">#${tx.id || index + 1}</span>
          <input type="date" class="card-date-input" value="${tx.date || ''}" data-field="date" data-idx="${index}">
          <select class="card-category-select" data-field="category" data-idx="${index}">
            ${optionsHtml}
          </select>
        </div>
        <button class="btn-card-del" data-delete-idx="${index}" title="Hapus Transaksi">🗑️</button>
      </div>

      <div class="card-desc-box">
        <input type="text" class="card-desc-input" placeholder="Keterangan transaksi..." value="${tx.desc || ''}" data-field="desc" data-idx="${index}">
      </div>

      <div class="card-finance-grid">
        <div class="finance-box">
          <span class="finance-label income-lbl">Pemasukan</span>
          <div class="currency-box-card">
            <span class="currency-tag-card">Rp</span>
            <input type="text" class="finance-input income-in currency-field" placeholder="0" value="${formatCommaNumber(tx.income)}" data-field="income" data-idx="${index}">
          </div>
        </div>
        <div class="finance-box">
          <span class="finance-label expense-lbl">Pengeluaran</span>
          <div class="currency-box-card">
            <span class="currency-tag-card">Rp</span>
            <input type="text" class="finance-input expense-in currency-field" placeholder="0" value="${formatCommaNumber(tx.expense)}" data-field="expense" data-idx="${index}">
          </div>
        </div>
        <div class="finance-box">
          <span class="finance-label balance-lbl">Sisa Kas</span>
          <div class="currency-box-card">
            <span class="currency-tag-card">Rp</span>
            <input type="text" class="finance-input balance-in currency-field" placeholder="0" value="${formatCommaNumber(tx.balance)}" data-field="balance" data-idx="${index}">
          </div>
        </div>
      </div>
    `;

    mobileCardsSection.appendChild(card);
  });
}

function renderTable() {
  tableBody.innerHTML = '';
  const filtered = getFilteredTransactions();

  if (filtered.length === 0) {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td colspan="8" style="text-align: center; padding: 24px; color: var(--text-muted); font-weight: 600;">
        Tidak ada data transaksi yang cocok.
      </td>
    `;
    tableBody.appendChild(tr);
    return;
  }

  filtered.forEach((tx, index) => {
    const tr = document.createElement('tr');

    const optionsHtml = CATEGORIES.map((c) => {
      const sel = tx.category === c ? 'selected' : '';
      return `<option value="${c}" ${sel}>${c}</option>`;
    }).join('');

    tr.innerHTML = `
      <td class="cell-row-header">${tx.id || index + 1}</td>
      <td><input type="date" class="cell-input text-center" value="${tx.date || ''}" data-field="date" data-idx="${index}"></td>
      <td><input class="cell-input" placeholder="Tulis rincian..." value="${tx.desc || ''}" data-field="desc" data-idx="${index}"></td>
      <td>
        <select class="cell-select" data-field="category" data-idx="${index}">
          ${optionsHtml}
        </select>
      </td>
      <td>
        <div class="currency-cell-wrapper">
          <span class="currency-prefix">Rp</span>
          <input type="text" class="cell-input text-right font-mono income-val currency-field" placeholder="0" value="${formatCommaNumber(tx.income)}" data-field="income" data-idx="${index}">
        </div>
      </td>
      <td>
        <div class="currency-cell-wrapper">
          <span class="currency-prefix">Rp</span>
          <input type="text" class="cell-input text-right font-mono expense-val currency-field" placeholder="0" value="${formatCommaNumber(tx.expense)}" data-field="expense" data-idx="${index}">
        </div>
      </td>
      <td>
        <div class="currency-cell-wrapper">
          <span class="currency-prefix">Rp</span>
          <input type="text" class="cell-input text-right font-mono balance-val currency-field" placeholder="0" value="${formatCommaNumber(tx.balance)}" data-field="balance" data-idx="${index}">
        </div>
      </td>
      <td style="text-align: center;"><button class="btn-del-row" data-delete-idx="${index}" title="Hapus Baris">🗑️</button></td>
    `;

    tableBody.appendChild(tr);
  });
}

function renderAll() {
  renderCards();
  renderTable();
  updateTotals();
}

function setView(viewMode) {
  currentView = viewMode;
  localStorage.setItem(VIEW_PREF_KEY, currentView);
  updateViewUI();
}

function updateViewUI() {
  if (currentView === 'cards') {
    mobileCardsSection.style.display = 'flex';
    spreadsheetTableSection.style.display = 'none';
    tabViewCards.classList.add('active');
    tabViewTable.classList.remove('active');

    dockSwitchIcon.textContent = '📊';
    dockSwitchLabel.textContent = 'Tabel';
  } else {
    mobileCardsSection.style.display = 'none';
    spreadsheetTableSection.style.display = 'block';
    tabViewCards.classList.remove('active');
    tabViewTable.classList.add('active');

    dockSwitchIcon.textContent = '📱';
    dockSwitchLabel.textContent = 'Kartu HP';
  }
}

// Unified input change handler
function handleFieldChange(e) {
  const target = e.target;
  const field = target.dataset.field;
  const idx = parseInt(target.dataset.idx, 10);

  if (field && !isNaN(idx) && transactions[idx]) {
    const rawVal = target.value;
    
    if (field === 'income' || field === 'expense') {
      transactions[idx][field] = parseNumber(rawVal);
      if (isAutoBalance) {
        recalculateBalances();
      }
    } else if (field === 'balance') {
      transactions[idx].balance = parseNumber(rawVal);
    } else if (field === 'id') {
      transactions[idx].id = rawVal;
    } else {
      transactions[idx][field] = rawVal;
    }

    saveLocalData(true);
    renderAll();
    syncWithGoogleSheets(true);
  }
}

mobileCardsSection.addEventListener('change', handleFieldChange);
tableBody.addEventListener('change', handleFieldChange);

function handleDelete(idx) {
  if (!isNaN(idx) && transactions[idx]) {
    const confirmDelete = confirm(`Hapus transaksi "${transactions[idx].desc || 'baris ini'}"?`);
    if (confirmDelete) {
      transactions.splice(idx, 1);
      if (isAutoBalance) recalculateBalances();
      saveLocalData(true);
      renderAll();
      syncWithGoogleSheets(true);
      showToast('Transaksi berhasil dihapus');
    }
  }
}

mobileCardsSection.addEventListener('click', (e) => {
  const btn = e.target.closest('[data-delete-idx]');
  if (btn) handleDelete(parseInt(btn.dataset.deleteIdx, 10));
});

tableBody.addEventListener('click', (e) => {
  const btn = e.target.closest('[data-delete-idx]');
  if (btn) handleDelete(parseInt(btn.dataset.deleteIdx, 10));
});

function addNewRow() {
  const today = new Date().toISOString().split('T')[0];
  const newId = transactions.length > 0 ? transactions.length + 1 : 1;
  const lastBalance = transactions.length > 0 ? transactions[transactions.length - 1].balance : 0;

  const newTx = {
    id: newId,
    date: today,
    desc: '',
    category: 'Iuran Kas',
    income: 0,
    expense: 0,
    balance: lastBalance
  };

  transactions.push(newTx);
  saveLocalData(true);
  renderAll();
  syncWithGoogleSheets(true);

  setTimeout(() => {
    if (currentView === 'cards') {
      const cards = mobileCardsSection.querySelectorAll('.mobile-tx-card');
      if (cards.length > 0) {
        const lastCard = cards[cards.length - 1];
        lastCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        const descInput = lastCard.querySelector('[data-field="desc"]');
        if (descInput) descInput.focus();
      }
    } else {
      const rows = tableBody.querySelectorAll('tr');
      if (rows.length > 0) {
        const lastRow = rows[rows.length - 1];
        lastRow.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        const descInput = lastRow.querySelector('[data-field="desc"]');
        if (descInput) descInput.focus();
      }
    }
  }, 100);

  showToast('Baris transaksi baru ditambahkan');
}

function toggleAutoBalance() {
  isAutoBalance = !isAutoBalance;
  updateAutoBalanceUI();
  if (isAutoBalance) {
    recalculateBalances();
    showToast('Mode: Sisa dihitung otomatis');
  } else {
    showToast('Mode: Sisa bisa diisi bebas');
  }
  saveLocalData(false);
  renderAll();
  syncWithGoogleSheets(true);
}

function updateAutoBalanceUI() {
  if (isAutoBalance) {
    autoBalanceIcon.textContent = '🔄';
    autoBalanceLabel.textContent = 'Sisa: Otomatis';
    btnToggleAutoBalance.classList.add('btn-primary');
    btnToggleAutoBalance.classList.remove('btn-outline');
    document.getElementById('sheetBadge').textContent = 'Formula Saldo Otomatis';
  } else {
    autoBalanceIcon.textContent = '✏️';
    autoBalanceLabel.textContent = 'Sisa: Manual';
    btnToggleAutoBalance.classList.remove('btn-primary');
    btnToggleAutoBalance.classList.add('btn-outline');
    document.getElementById('sheetBadge').textContent = 'Semua Kolom Bebas Diisi';
  }
}

function exportToCsv() {
  let csvContent = '\uFEFF';
  csvContent += 'No,Tanggal,Keterangan Transaksi,Kategori,Pemasukan (Rp),Pengeluaran (Rp),Sisa Kas (Rp)\n';

  let totalInc = 0;
  let totalExp = 0;

  transactions.forEach((tx) => {
    const safeDesc = `"${(tx.desc || '').replace(/"/g, '""')}"`;
    const safeCat = `"${(tx.category || '').replace(/"/g, '""')}"`;
    const inc = parseNumber(tx.income);
    const exp = parseNumber(tx.expense);
    const bal = parseNumber(tx.balance);

    totalInc += inc;
    totalExp += exp;

    csvContent += `${tx.id || ''},${tx.date || ''},${safeDesc},${safeCat},${inc},${exp},${bal}\n`;
  });

  csvContent += `TOTAL,,,,${totalInc},${totalExp},${totalInc - totalExp}\n`;

  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Kas_Futsal_${new Date().toISOString().split('T')[0]}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  showToast('File CSV berhasil diunduh');
  exportModalBackdrop.classList.remove('open');
}

function exportToJson() {
  const dataStr = 'data:text/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(transactions, null, 2));
  const a = document.createElement('a');
  a.href = dataStr;
  a.download = `Backup_Kas_Futsal_${new Date().toISOString().split('T')[0]}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  showToast('Backup JSON berhasil diunduh');
  exportModalBackdrop.classList.remove('open');
}

function restoreFromSnapshot() {
  try {
    const raw = localStorage.getItem(SNAPSHOT_KEY);
    if (!raw) {
      alert('Belum ada riwayat cadangan otomatis yang tersimpan.');
      return;
    }
    const snap = JSON.parse(raw);
    const dateStr = snap.timestamp ? new Date(snap.timestamp).toLocaleString() : 'Sebelumnya';
    const confirmRestore = confirm(`Pulihkan data kas dari cadangan otomatis (${dateStr})?`);
    if (confirmRestore && Array.isArray(snap.items)) {
      transactions = snap.items;
      saveLocalData(false);
      renderAll();
      syncWithGoogleSheets(true);
      showToast('Data berhasil dipulihkan dari cadangan', '⏮️');
      exportModalBackdrop.classList.remove('open');
    }
  } catch (e) {
    alert('Gagal memulihkan cadangan: ' + e);
  }
}

function resetToTemplate() {
  const confirmReset = confirm('PERINGATAN: Apakah Anda yakin ingin mereset data ke contoh awal?');
  if (confirmReset) {
    transactions = JSON.parse(JSON.stringify(DEFAULT_TRANSACTIONS));
    isAutoBalance = true;
    updateAutoBalanceUI();
    saveLocalData(false);
    renderAll();
    syncWithGoogleSheets(true);
    showToast('Data direset ke contoh awal');
    exportModalBackdrop.classList.remove('open');
  }
}

// Event Listeners
btnAddRow.addEventListener('click', addNewRow);
dockBtnAdd.addEventListener('click', addNewRow);
btnToggleAutoBalance.addEventListener('click', toggleAutoBalance);

tabViewCards.addEventListener('click', () => setView('cards'));
tabViewTable.addEventListener('click', () => setView('table'));
dockBtnSwitchView.addEventListener('click', () => {
  setView(currentView === 'cards' ? 'table' : 'cards');
});

searchInput.addEventListener('input', (e) => {
  searchQuery = e.target.value.trim();
  renderAll();
});

btnTriggerSync.addEventListener('click', () => syncWithGoogleSheets(true));
btnSyncGSheet.addEventListener('click', () => syncWithGoogleSheets(true));

btnOpenIosGuide.addEventListener('click', () => iosModalBackdrop.classList.add('open'));
dockBtnIos.addEventListener('click', () => iosModalBackdrop.classList.add('open'));
btnCloseIosModal.addEventListener('click', () => iosModalBackdrop.classList.remove('open'));
iosModalBackdrop.addEventListener('click', (e) => {
  if (e.target === iosModalBackdrop) iosModalBackdrop.classList.remove('open');
});

btnExportMenu.addEventListener('click', () => exportModalBackdrop.classList.add('open'));
dockBtnExport.addEventListener('click', () => exportModalBackdrop.classList.add('open'));
btnCloseExportModal.addEventListener('click', () => exportModalBackdrop.classList.remove('open'));
exportModalBackdrop.addEventListener('click', (e) => {
  if (e.target === exportModalBackdrop) exportModalBackdrop.classList.remove('open');
});

btnExportCsv.addEventListener('click', exportToCsv);
btnExportJson.addEventListener('click', exportToJson);
btnRestoreSnapshot.addEventListener('click', restoreFromSnapshot);
btnResetToTemplate.addEventListener('click', resetToTemplate);

document.getElementById('dockBtnSave').addEventListener('click', () => {
  saveLocalData(true);
  syncWithGoogleSheets(true);
});

window.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'hidden') {
    saveLocalData(false);
    syncWithGoogleSheets(true);
  }
});

window.addEventListener('pagehide', () => {
  saveLocalData(false);
});

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./sw.js')
      .then((reg) => console.log('ServiceWorker registered:', reg.scope))
      .catch((err) => console.warn('ServiceWorker registration failed:', err));
  });
}

(async function init() {
  requestPersistentStorage();
  await loadLocalData();
  if (isAutoBalance) recalculateBalances();
  renderAll();
  syncWithGoogleSheets(false);
})();
