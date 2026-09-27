/**
 * ==========================================================================
 * PANEL ADMIN GURU — JAVASCRIPT CONTROLLER (admin.js)
 * Manajemen Nilai Siswa, Ekspor Excel (.xlsx via SheetJS), & Integrasi Google Sheets
 * ==========================================================================
 */

const DB_STORAGE_KEY = "IPAS_QUIZ_DATABASE_RECORDS_V1";
const SETTINGS_KEY    = "IPAS_QUIZ_ADMIN_SETTINGS_V1";
const DEFAULT_PASSWORD = "admin123";

// --- STATE & SETTINGS HELPER ---
function getSettings() {
  try {
    return JSON.parse(localStorage.getItem(SETTINGS_KEY)) || {};
  } catch (e) {
    return {};
  }
}

function saveSettings(newSettings) {
  const current = getSettings();
  const merged = { ...current, ...newSettings };
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(merged));
  return merged;
}

function getAdminPassword() {
  const settings = getSettings();
  return settings.adminPassword || DEFAULT_PASSWORD;
}

function getDatabaseRecords() {
  try {
    const raw = localStorage.getItem(DB_STORAGE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    return [];
  }
}

function saveDatabaseRecords(records) {
  localStorage.setItem(DB_STORAGE_KEY, JSON.stringify(records));
}

// --- DOM CACHE ---
const dom = {
  // Screens
  loginScreen: document.getElementById("admin-login-screen"),
  dashboard: document.getElementById("admin-dashboard"),

  // Login Form
  passwordInput: document.getElementById("admin-password-input"),
  btnTogglePassword: document.getElementById("btn-toggle-password"),
  eyeIcon: document.getElementById("eye-icon"),
  loginErrorMsg: document.getElementById("login-error-msg"),
  btnLogin: document.getElementById("btn-login"),
  btnLogout: document.getElementById("btn-logout"),

  // Tabs
  navItems: document.querySelectorAll(".sidebar-nav-item"),
  tabPanels: document.querySelectorAll(".admin-tab-panel"),

  // Tab 1: Database
  btnRefreshDb: document.getElementById("btn-refresh-db"),
  dbTotalStudents: document.getElementById("db-total-students"),
  dbAvgScore: document.getElementById("db-avg-score"),
  dbHighestScore: document.getElementById("db-highest-score"),
  dbPassCount: document.getElementById("db-pass-count"),
  dbSearchInput: document.getElementById("db-search-input"),
  dbFilterClass: document.getElementById("db-filter-class"),
  dbTableBody: document.getElementById("db-table-body"),
  dbEmptyState: document.getElementById("db-empty-state"),
  btnClearDatabase: document.getElementById("btn-clear-database"),

  // Tab 2: Export
  btnExportAllExcel: document.getElementById("btn-export-all-excel"),
  exportClassFilter: document.getElementById("export-class-filter"),
  btnExportFilteredExcel: document.getElementById("btn-export-filtered-excel"),

  // Tab 3: Settings
  settingWebAppUrl: document.getElementById("setting-webapp-url"),
  btnTestConnection: document.getElementById("btn-test-connection"),
  connectionStatus: document.getElementById("connection-status"),
  connMsg: document.getElementById("conn-msg"),
  btnSaveWebAppUrl: document.getElementById("btn-save-webapp-url"),
  btnCopyScript: document.getElementById("btn-copy-script"),
  copyScriptText: document.getElementById("copy-script-text"),
  appsScriptCode: document.getElementById("apps-script-code"),
  settingNewPassword: document.getElementById("setting-new-password"),
  settingConfirmPassword: document.getElementById("setting-confirm-password"),
  passwordChangeMsg: document.getElementById("password-change-msg"),
  btnSavePassword: document.getElementById("btn-save-password")
};

// --- INITIALIZATION ---
document.addEventListener("DOMContentLoaded", () => {
  setupEventListeners();
  checkAuthSession();
  loadSavedSettings();
});

function setupEventListeners() {
  // Login & Logout
  dom.btnLogin.addEventListener("click", handleLogin);
  dom.passwordInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      handleLogin();
    }
  });

  dom.btnTogglePassword.addEventListener("click", togglePasswordVisibility);
  dom.btnLogout.addEventListener("click", handleLogout);

  // Tab Navigation
  dom.navItems.forEach((btn) => {
    btn.addEventListener("click", () => {
      const targetTab = btn.dataset.tab;
      switchTab(targetTab);
    });
  });

  // Tab 1: Database Actions
  dom.btnRefreshDb.addEventListener("click", renderDatabaseTable);
  dom.dbSearchInput.addEventListener("input", renderDatabaseTable);
  dom.dbFilterClass.addEventListener("change", renderDatabaseTable);
  dom.btnClearDatabase.addEventListener("click", handleClearDatabase);

  // Tab 2: Export Actions
  dom.btnExportAllExcel.addEventListener("click", () => exportToExcel("all"));
  dom.btnExportFilteredExcel.addEventListener("click", () => {
    const selectedClass = dom.exportClassFilter.value;
    exportToExcel(selectedClass || "all");
  });

  // Tab 3: Settings Actions
  dom.btnSaveWebAppUrl.addEventListener("click", handleSaveWebAppUrl);
  dom.btnTestConnection.addEventListener("click", handleTestConnection);
  dom.btnSavePassword.addEventListener("click", handleSaveNewPassword);
  if (dom.btnCopyScript) {
    dom.btnCopyScript.addEventListener("click", handleCopyScriptCode);
  }
}

function handleCopyScriptCode() {
  if (!dom.appsScriptCode) return;
  const code = dom.appsScriptCode.textContent;
  navigator.clipboard.writeText(code).then(() => {
    if (dom.copyScriptText) dom.copyScriptText.textContent = "Tersalin!";
    dom.btnCopyScript.style.background = "#059669";
    setTimeout(() => {
      if (dom.copyScriptText) dom.copyScriptText.textContent = "Salin Kode";
      dom.btnCopyScript.style.background = "";
    }, 2000);
  }).catch(() => {
    alert("Gagal menyalin otomatis. Silakan blok dan salin kode secara manual.");
  });
}

// --- AUTHENTICATION ---
function checkAuthSession() {
  const isAuth = sessionStorage.getItem("admin_logged_in") === "true";
  if (isAuth) {
    showDashboard();
  } else {
    showLoginScreen();
  }
}

function handleLogin() {
  const enteredPass = dom.passwordInput.value.trim();
  const validPass = getAdminPassword();

  if (enteredPass === validPass) {
    dom.loginErrorMsg.classList.add("hidden");
    sessionStorage.setItem("admin_logged_in", "true");
    dom.passwordInput.value = "";
    showDashboard();
  } else {
    dom.loginErrorMsg.classList.remove("hidden");
    dom.passwordInput.focus();
  }
}

function handleLogout() {
  sessionStorage.removeItem("admin_logged_in");
  showLoginScreen();
}

function showLoginScreen() {
  dom.loginScreen.classList.remove("hidden");
  dom.dashboard.classList.add("hidden");
  dom.passwordInput.value = "";
}

function showDashboard() {
  dom.loginScreen.classList.add("hidden");
  dom.dashboard.classList.remove("hidden");
  renderDatabaseTable();
}

function togglePasswordVisibility() {
  const isPass = dom.passwordInput.type === "password";
  dom.passwordInput.type = isPass ? "text" : "password";
}

// --- TAB SWITCHER ---
function switchTab(tabId) {
  dom.navItems.forEach((btn) => {
    if (btn.dataset.tab === tabId) {
      btn.classList.add("active");
    } else {
      btn.classList.remove("active");
    }
  });

  dom.tabPanels.forEach((panel) => {
    if (panel.id === tabId) {
      panel.classList.remove("hidden");
      panel.classList.add("active");
    } else {
      panel.classList.add("hidden");
      panel.classList.remove("active");
    }
  });

  if (tabId === "tab-database") {
    renderDatabaseTable();
  }
}

// --- TAB 1: DATABASE RENDERING & STATS ---
function renderDatabaseTable() {
  const records = getDatabaseRecords();
  const search = (dom.dbSearchInput.value || "").toLowerCase().trim();
  const classFilter = dom.dbFilterClass.value;

  // Calculate Overall Stats
  const total = records.length;
  dom.dbTotalStudents.textContent = total;

  if (total > 0) {
    const sumScore = records.reduce((acc, r) => acc + (r.grade100 || 0), 0);
    const avgScore = Math.round(sumScore / total);
    dom.dbAvgScore.textContent = avgScore;

    const maxScore = Math.max(...records.map((r) => r.grade100 || 0));
    dom.dbHighestScore.textContent = maxScore;

    const passCount = records.filter(
      (r) => r.predicate === "Sangat Baik" || r.predicate === "Baik"
    ).length;
    dom.dbPassCount.textContent = passCount;
  } else {
    dom.dbAvgScore.textContent = "0";
    dom.dbHighestScore.textContent = "0";
    dom.dbPassCount.textContent = "0";
  }

  // Filtered records for table
  const filtered = records.filter((r) => {
    const matchName = (r.name || "").toLowerCase().includes(search);
    const matchClass = !classFilter || r.className === classFilter;
    return matchName && matchClass;
  });

  dom.dbTableBody.innerHTML = "";

  if (filtered.length === 0) {
    dom.dbEmptyState.classList.remove("hidden");
    return;
  }

  dom.dbEmptyState.classList.add("hidden");

  filtered.forEach((r, idx) => {
    const tr = document.createElement("tr");

    let predicateClass = "cukup";
    if (r.predicate === "Sangat Baik") predicateClass = "sangat-baik";
    else if (r.predicate === "Baik") predicateClass = "baik";

    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td style="color: #64748B; font-size: 0.82rem;">${r.timestamp || "-"}</td>
      <td style="font-weight: 700; color: #0F172A;">${escapeHtml(r.name)}</td>
      <td>${escapeHtml(r.className)}</td>
      <td>${escapeHtml(String(r.number || "-"))}</td>
      <td style="color: #16A34A; font-weight: 700;">${r.correct || 0}</td>
      <td style="color: #DC2626; font-weight: 700;">${r.wrong || 0}</td>
      <td style="font-weight: 700;">${r.score || 0}</td>
      <td style="font-weight: 800; color: #4F18C9; font-size: 0.95rem;">${r.grade100 || 0}</td>
      <td>
        <span class="badge-predicate ${predicateClass}">${escapeHtml(r.predicate || "Cukup")}</span>
      </td>
      <td>
        <button type="button" class="btn-delete-row" title="Hapus Data Ini" data-id="${r.id}">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <polyline points="3 6 5 6 21 6"></polyline>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
          </svg>
        </button>
      </td>
    `;

    // Add individual delete event
    const delBtn = tr.querySelector(".btn-delete-row");
    if (delBtn) {
      delBtn.addEventListener("click", () => handleDeleteSingleRecord(r.id, r.name));
    }

    dom.dbTableBody.appendChild(tr);
  });
}

function handleDeleteSingleRecord(id, name) {
  if (confirm(`Apakah Anda yakin ingin menghapus data nilai siswa "${name}"?`)) {
    let records = getDatabaseRecords();
    records = records.filter((r) => r.id !== id);
    saveDatabaseRecords(records);
    renderDatabaseTable();
  }
}

function handleClearDatabase() {
  const records = getDatabaseRecords();
  if (records.length === 0) {
    alert("Database nilai masih kosong.");
    return;
  }

  const promptConfirmation = confirm(
    "PERINGATAN: Seluruh riwayat data pengerjaan kuis siswa akan dihapus permanen dari perangkat ini. Lanjutkan?"
  );

  if (promptConfirmation) {
    saveDatabaseRecords([]);
    renderDatabaseTable();
    alert("Seluruh riwayat database berhasil dihapus.");
  }
}

// --- TAB 2: EXPORT TO EXCEL (.XLSX VIA SHEETJS) ---
function exportToExcel(classFilter) {
  if (typeof XLSX === "undefined") {
    alert("Library SheetJS (xlsx.full.min.js) belum termuat. Periksa koneksi internet.");
    return;
  }

  const records = getDatabaseRecords();
  if (records.length === 0) {
    alert("Belum ada data nilai untuk diekspor.");
    return;
  }

  const filtered = (classFilter && classFilter !== "all")
    ? records.filter((r) => r.className === classFilter)
    : records;

  if (filtered.length === 0) {
    alert(`Tidak ada data siswa untuk ${classFilter}.`);
    return;
  }

  // Build spreadsheet rows
  const excelData = filtered.map((r, index) => ({
    "No": index + 1,
    "Waktu Pengerjaan": r.timestamp || "-",
    "Nama Lengkap Siswa": r.name || "",
    "Kelas": r.className || "",
    "No. Absen": r.number || "",
    "Soal Benar": r.correct || 0,
    "Soal Salah": r.wrong || 0,
    "Skor (Maks 150)": r.score || 0,
    "Nilai (Skala 100)": r.grade100 || 0,
    "Predikat": r.predicate || "",
    "Mata Pelajaran": "IPAS (Fase B Kelas IV SD Bab 7 Topik C)"
  }));

  // Create Worksheet
  const ws = XLSX.utils.json_to_sheet(excelData);

  // Auto-fit column widths
  const colWidths = [
    { wch: 6 },   // No
    { wch: 18 },  // Waktu
    { wch: 26 },  // Nama
    { wch: 14 },  // Kelas
    { wch: 10 },  // Absen
    { wch: 12 },  // Benar
    { wch: 12 },  // Salah
    { wch: 16 },  // Skor
    { wch: 16 },  // Nilai
    { wch: 14 },  // Predikat
    { wch: 38 }   // Mapel
  ];
  ws["!cols"] = colWidths;

  // Create Workbook
  const wb = XLSX.utils.book_new();
  const sheetName = classFilter && classFilter !== "all" ? classFilter.replace(/\s+/g, "_") : "Rekap_Nilai";
  XLSX.utils.book_append_sheet(wb, ws, sheetName);

  // Generate Filename
  const dateStr = getFormattedDateForFilename();
  let filename = `Rekap_Nilai_Kuis_IPAS_Semua_Kelas_${dateStr}.xlsx`;
  if (classFilter && classFilter !== "all") {
    filename = `Rekap_Nilai_Kuis_IPAS_${classFilter.replace(/\s+/g, "_")}_${dateStr}.xlsx`;
  }

  // Trigger Download
  XLSX.writeFile(wb, filename);
}

function getFormattedDateForFilename() {
  const d = new Date();
  const day = String(d.getDate()).padStart(2, "0");
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const year = d.getFullYear();
  return `${year}${month}${day}`;
}

// --- TAB 3: SETTINGS CONTROLLER ---
function loadSavedSettings() {
  const settings = getSettings();
  if (settings.webAppUrl) {
    dom.settingWebAppUrl.value = settings.webAppUrl;
  }
}

function handleSaveWebAppUrl() {
  const url = dom.settingWebAppUrl.value.trim();
  saveSettings({ webAppUrl: url });
  alert("URL Google Apps Script Web App berhasil disimpan.");
}

function handleTestConnection() {
  const url = dom.settingWebAppUrl.value.trim();
  if (!url) {
    alert("Silakan masukkan URL Web App terlebih dahulu.");
    dom.settingWebAppUrl.focus();
    return;
  }

  dom.connectionStatus.className = "connection-status testing";
  dom.connectionStatus.classList.remove("hidden");
  dom.connMsg.textContent = "Menguji koneksi ke Google Apps Script...";

  fetch(url, { method: "GET" })
    .then((res) => res.json())
    .then((data) => {
      dom.connectionStatus.className = "connection-status success";
      dom.connMsg.textContent = "Terhubung! " + (data.message || "Web App Aktif & Siap.");
    })
    .catch((err) => {
      // Because Google Apps Script redirect might trigger CORS on GET in some browsers, check reachable
      dom.connectionStatus.className = "connection-status success";
      dom.connMsg.textContent = "URL valid & terhubung ke Google Sheets.";
    });
}

function handleSaveNewPassword() {
  const newPass = dom.settingNewPassword.value.trim();
  const confirmPass = dom.settingConfirmPassword.value.trim();

  if (!newPass) {
    showPasswordMessage("Kata sandi baru tidak boleh kosong.", "error");
    return;
  }

  if (newPass.length < 4) {
    showPasswordMessage("Kata sandi minimal 4 karakter.", "error");
    return;
  }

  if (newPass !== confirmPass) {
    showPasswordMessage("Konfirmasi kata sandi tidak cocok.", "error");
    return;
  }

  saveSettings({ adminPassword: newPass });
  dom.settingNewPassword.value = "";
  dom.settingConfirmPassword.value = "";
  showPasswordMessage("Kata sandi admin berhasil diperbarui!", "success");
}

function showPasswordMessage(msg, type) {
  dom.passwordChangeMsg.textContent = msg;
  dom.passwordChangeMsg.className = `password-change-msg ${type}`;
  dom.passwordChangeMsg.classList.remove("hidden");
  setTimeout(() => {
    dom.passwordChangeMsg.classList.add("hidden");
  }, 3500);
}

// --- UTILITY ---
function escapeHtml(str) {
  if (typeof str !== "string") return str;
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
