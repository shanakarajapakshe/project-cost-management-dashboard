'use strict';
/* ============================================================
   Profit Dashboard – Renderer Process
   ============================================================
   Modules:
     DataManager      – all data & Excel IPC calls
     UIController     – DOM manipulation, charts, toasts
     UpdateController – electron-updater UI
     AppController    – wires everything together
   ============================================================ */

// ────────────────────────────────────────────────────────────────────────────
// DataManager
// ────────────────────────────────────────────────────────────────────────────
class DataManager {
  constructor() {
    this.currentMonth = null;
    this.currentYear  = null;
    this.projects     = [];
    /** monthlyData is used for the trend chart only (in-memory across sessions). */
    this.monthlyData  = {};
    this.isElectron   = typeof window.electronAPI !== 'undefined';
  }

  setCurrentPeriod(month, year) {
    this.currentMonth = String(month).padStart(2, '0');
    this.currentYear  = String(year);
  }

  getCurrentPeriod() {
    return {
      month: this.currentMonth,
      year:  this.currentYear,
      key:   `${this.currentYear}-${this.currentMonth}`
    };
  }

  // ── Calculations ─────────────────────────────────────────────────────────
  /**
   * BUG-FIX: original engineerCost = engineerSalary only (ignored numEngineers).
   * Correct formula: engineerCost = numEngineers * engineerSalary.
   */
  calculateProjectMetrics(data) {
    const numEngineers      = Number(data.numEngineers)      || 0;
    const engineerSalary    = Number(data.engineerSalary)    || 0;
    const ceVisitCharge     = Number(data.ceVisitCharge)     || 0;
    const visitsPerMonth    = Number(data.visitsPerMonth)    || 0;
    const transportCost     = Number(data.transportCost)     || 0;
    const clientPayment     = Number(data.clientPayment)     || 0;
    const overheadAllocation= Number(data.overheadAllocation)|| 0;

    const engineerCost  = numEngineers * engineerSalary;
    const ceVisitCost   = ceVisitCharge * visitsPerMonth;
    const directCost    = engineerCost + ceVisitCost + transportCost;
    const overheadCost  = directCost * (overheadAllocation / 100);
    const totalCost     = directCost + overheadCost;
    const profit        = clientPayment - totalCost;

    return {
      ...data,
      numEngineers, engineerSalary, ceVisitCharge, visitsPerMonth,
      transportCost, clientPayment, overheadAllocation,
      engineerCost, ceVisitCost, directCost, overheadCost, totalCost, profit,
      timestamp: new Date().toISOString()
    };
  }

  // ── CRUD ─────────────────────────────────────────────────────────────────
  async loadProjects() {
    if (this.isElectron) {
      const result = await window.electronAPI.loadOrCreateExcel(this.currentMonth, this.currentYear);
      if (!result.success) throw new Error(result.error || 'Failed to load projects');
      this.projects = (result.data || []).map(p => this._normaliseLoaded(p));
    } else {
      this.projects = this._loadFromLocalStorage();
    }
    this._updateMonthlyData();
    return this.projects;
  }

  async addProject(rawData) {
    const data = this.calculateProjectMetrics(rawData);
    if (this.isElectron) {
      const result = await window.electronAPI.saveProject(this.currentMonth, this.currentYear, data);
      if (!result.success) throw new Error(result.error || 'Failed to save project');
    }
    this.projects.push(data);
    if (!this.isElectron) this._saveToLocalStorage();
    this._updateMonthlyData();
    return data;
  }

  async updateProject(originalName, rawData) {
    const data = this.calculateProjectMetrics(rawData);
    if (this.isElectron) {
      const result = await window.electronAPI.updateProject(
        this.currentMonth, this.currentYear, originalName, data
      );
      if (!result.success) throw new Error(result.error || 'Failed to update project');
    }
    const idx = this.projects.findIndex(p => p.projectName === originalName);
    if (idx !== -1) this.projects[idx] = data;
    if (!this.isElectron) this._saveToLocalStorage();
    this._updateMonthlyData();
    return data;
  }

  async deleteProject(index) {
    const project = this.projects[index];
    if (!project) throw new Error('Project not found at index ' + index);
    if (this.isElectron) {
      const result = await window.electronAPI.deleteProject(
        this.currentMonth, this.currentYear, project.projectName
      );
      if (!result.success) throw new Error(result.error || 'Failed to delete project');
    }
    this.projects.splice(index, 1);
    if (!this.isElectron) this._saveToLocalStorage();
    this._updateMonthlyData();
  }

  async createBackup() {
    if (this.isElectron) {
      const result = await window.electronAPI.createBackup(this.currentMonth, this.currentYear);
      if (!result.success) throw new Error(result.error || 'Backup failed');
      return result.data; // path
    }
  }

  async exportToExcel() {
    if (!this.isElectron) throw new Error('Excel export requires the desktop app.');
    const result = await window.electronAPI.exportExcelFile(this.currentMonth, this.currentYear);
    if (result.canceled) return { canceled: true };
    if (!result.success) throw new Error(result.error || 'Export failed');
    return { path: result.data };
  }

  getProjects()     { return this.projects; }
  getSummaryStats() {
    const totalProjects = this.projects.length;
    const totalRevenue  = this.projects.reduce((s, p) => s + (p.clientPayment ?? 0), 0);
    const totalCosts    = this.projects.reduce((s, p) => s + (p.totalCost     ?? 0), 0);
    const netProfit     = totalRevenue - totalCosts;
    return { totalProjects, totalRevenue, totalCosts, netProfit };
  }

  getMonthlyTrend() {
    const sorted = Object.keys(this.monthlyData).sort();
    return sorted.slice(-6).map(key => ({
      period: key,
      profit: this.monthlyData[key]?.stats?.netProfit ?? 0
    }));
  }

  // ── Private ───────────────────────────────────────────────────────────────
  _normaliseLoaded(p) {
    // Ensure all numeric fields are numbers (ExcelJS may return objects)
    const nums = ['numEngineers','engineerSalary','ceVisitCharge','visitsPerMonth',
                  'transportCost','clientPayment','overheadAllocation',
                  'engineerCost','ceVisitCost','directCost','overheadCost','totalCost','profit'];
    const out = { ...p };
    nums.forEach(k => { out[k] = parseFloat(out[k]) || 0; });
    return out;
  }

  _updateMonthlyData() {
    const { key } = this.getCurrentPeriod();
    this.monthlyData[key] = { projects: [...this.projects], stats: this.getSummaryStats() };
  }

  _localStorageKey() {
    const { key } = this.getCurrentPeriod();
    return `profit_dashboard_${key}`;
  }

  _saveToLocalStorage() {
    try { localStorage.setItem(this._localStorageKey(), JSON.stringify(this.projects)); }
    catch (_) { /* storage full / not available */ }
  }

  _loadFromLocalStorage() {
    try {
      const raw = localStorage.getItem(this._localStorageKey());
      return raw ? JSON.parse(raw) : [];
    } catch (_) { return []; }
  }
}

// ────────────────────────────────────────────────────────────────────────────
// UIController
// ────────────────────────────────────────────────────────────────────────────
class UIController {
  constructor(dataManager) {
    this.dataManager   = dataManager;
    this.charts        = {};
    this.detailModal   = null;
    this._detailIndex  = null;
    this._editingName  = null;
    this._initElements();
    this._initCharts();
    this._initLivePreview();
  }

  _initElements() {
    const $ = id => document.getElementById(id);
    this.el = {
      monthSelect:       $('monthSelect'),
      yearSelect:        $('yearSelect'),
      loadMonthBtn:      $('loadMonthBtn'),
      currentPeriod:     $('currentPeriod'),
      mainContent:       $('mainContent'),
      projectForm:       $('projectForm'),
      formTitle:         $('formTitle'),
      submitFormBtn:     $('submitFormBtn'),
      resetFormBtn:      $('resetFormBtn'),
      cancelEditBtn:     $('cancelEditBtn'),
      editingProjectName:$('editingProjectName'),
      // form fields
      projectName:       $('projectName'),
      numEngineers:      $('numEngineers'),
      engineerSalary:    $('engineerSalary'),
      ceVisitCharge:     $('ceVisitCharge'),
      visitsPerMonth:    $('visitsPerMonth'),
      transportCost:     $('transportCost'),
      clientPayment:     $('clientPayment'),
      overheadAllocation:$('overheadAllocation'),
      calcPreview:       $('calcPreview'),
      // dashboard
      projectsTableBody: $('projectsTableBody'),
      totalProjects:     $('totalProjects'),
      totalRevenue:      $('totalRevenue'),
      totalCosts:        $('totalCosts'),
      netProfit:         $('netProfit'),
      backupBtn:         $('backupBtn'),
      printDashboardBtn: $('printDashboardBtn'),
      exportExcelBtn:    $('exportExcelBtn'),
      // toast
      toast:             $('toast'),
      toastHeader:       $('toastHeader'),
      toastMessage:      $('toastMessage'),
      // loading
      loadingOverlay:    $('loadingOverlay'),
      // update
      updateBanner:      $('updateBanner'),
      updateBannerText:  $('updateBannerText'),
      updateDownloadBtn: $('updateDownloadBtn'),
      updateInstallBtn:  $('updateInstallBtn'),
      updateProgress:    $('updateProgress'),
      updateProgressBar: $('updateProgressBar'),
      checkUpdateBtn:    $('checkUpdateBtn'),
      appVersionBadge:   $('appVersionBadge'),
      // modal
      detailModalBody:   $('detailModalBody'),
      editFromDetailBtn: $('editFromDetailBtn'),
    };
    this.detailModal = new bootstrap.Modal(document.getElementById('detailModal'));
  }

  // ── Year / month population ───────────────────────────────────────────────
  populateYearSelect() {
    const cur = new Date().getFullYear();
    for (let y = cur - 5; y <= cur + 2; y++) {
      const opt = document.createElement('option');
      opt.value = y; opt.textContent = y;
      if (y === cur) opt.selected = true;
      this.el.yearSelect.appendChild(opt);
    }
  }

  setCurrentMonth() {
    const m = new Date().getMonth() + 1;
    this.el.monthSelect.value = String(m).padStart(2, '0');
  }

  // ── Visibility ────────────────────────────────────────────────────────────
  showMainContent() {
    this.el.mainContent.classList.remove('d-none');
    this.el.mainContent.classList.add('fade-in');
  }

  showLoading(on) {
    this.el.loadingOverlay.classList.toggle('d-none', !on);
  }

  // ── Period display ────────────────────────────────────────────────────────
  updatePeriodDisplay() {
    const MONTHS = ['January','February','March','April','May','June',
                    'July','August','September','October','November','December'];
    const { month, year } = this.dataManager.getCurrentPeriod();
    const label = MONTHS[parseInt(month, 10) - 1];
    this.el.currentPeriod.innerHTML = `
      <strong>Current Period:</strong> ${label} ${year}
      <span class="ms-2 badge bg-primary">${this.dataManager.projects.length} project(s)</span>`;
    this.el.currentPeriod.classList.remove('d-none');
  }

  // ── Summary stats ─────────────────────────────────────────────────────────
  updateStats() {
    const { totalProjects, totalRevenue, totalCosts, netProfit } = this.dataManager.getSummaryStats();
    this.el.totalProjects.textContent = totalProjects;
    this.el.totalRevenue.textContent  = this.formatCurrency(totalRevenue);
    this.el.totalCosts.textContent    = this.formatCurrency(totalCosts);
    this.el.netProfit.textContent     = this.formatCurrency(netProfit);
    this.el.netProfit.className       = 'stat-value ' + (netProfit >= 0 ? 'profit-positive' : 'profit-negative');
  }

  // ── Projects table ────────────────────────────────────────────────────────
  updateTable() {
    const projects = this.dataManager.getProjects();
    const tbody    = this.el.projectsTableBody;

    if (!projects.length) {
      tbody.innerHTML = '<tr><td colspan="7" class="text-center text-muted py-4">No projects added yet</td></tr>';
      return;
    }

    tbody.innerHTML = projects.map((p, i) => `
      <tr>
        <td class="text-muted">${i + 1}</td>
        <td><strong>${this._esc(p.projectName)}</strong></td>
        <td>${p.numEngineers ?? 0}</td>
        <td>${this.formatCurrency(p.totalCost)}</td>
        <td>${this.formatCurrency(p.clientPayment)}</td>
        <td class="${(p.profit ?? 0) >= 0 ? 'profit-positive' : 'profit-negative'} fw-bold">
          ${this.formatCurrency(p.profit)}
        </td>
        <td class="text-nowrap">
          <button class="btn btn-sm btn-outline-primary me-1"
                  onclick="APP.viewProject(${i})" title="View details">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/>
            </svg>
          </button>
          <button class="btn btn-sm btn-outline-warning me-1"
                  onclick="APP.editProject(${i})" title="Edit">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/>
              <path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/>
            </svg>
          </button>
          <button class="btn btn-sm btn-outline-danger"
                  onclick="APP.deleteProject(${i})" title="Delete">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6m3 0V4a1 1 0 011-1h4a1 1 0 011 1v2"/>
            </svg>
          </button>
        </td>
      </tr>`).join('');
  }

  // ── Form helpers ──────────────────────────────────────────────────────────
  getFormData() {
    return {
      projectName:        this.el.projectName.value.trim(),
      numEngineers:       parseFloat(this.el.numEngineers.value)       || 0,
      engineerSalary:     parseFloat(this.el.engineerSalary.value)     || 0,
      ceVisitCharge:      parseFloat(this.el.ceVisitCharge.value)      || 0,
      visitsPerMonth:     parseFloat(this.el.visitsPerMonth.value)     || 0,
      transportCost:      parseFloat(this.el.transportCost.value)      || 0,
      clientPayment:      parseFloat(this.el.clientPayment.value)      || 0,
      overheadAllocation: parseFloat(this.el.overheadAllocation.value) || 0,
    };
  }

  fillForm(project) {
    this.el.projectName.value        = project.projectName        ?? '';
    this.el.numEngineers.value       = project.numEngineers       ?? '';
    this.el.engineerSalary.value     = project.engineerSalary     ?? '';
    this.el.ceVisitCharge.value      = project.ceVisitCharge      ?? '';
    this.el.visitsPerMonth.value     = project.visitsPerMonth     ?? '';
    this.el.transportCost.value      = project.transportCost      ?? '';
    this.el.clientPayment.value      = project.clientPayment      ?? '';
    this.el.overheadAllocation.value = project.overheadAllocation ?? '';
    this._updatePreview();
  }

  enterEditMode(projectName) {
    this._editingName = projectName;
    this.el.editingProjectName.value = projectName;
    this.el.formTitle.textContent    = 'Edit Project';
    this.el.submitFormBtn.textContent = 'Update Project';
    this.el.cancelEditBtn.classList.remove('d-none');
    this.el.projectForm.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  exitEditMode() {
    this._editingName = null;
    this.el.editingProjectName.value  = '';
    this.el.formTitle.textContent     = 'Add New Project';
    this.el.submitFormBtn.textContent = 'Save Project';
    this.el.cancelEditBtn.classList.add('d-none');
    this.el.projectForm.classList.remove('was-validated');
    this.el.projectForm.reset();
    this._updatePreview();
  }

  getEditingName() { return this._editingName; }

  resetForm() {
    this.exitEditMode();
  }

  validateForm() {
    this.el.projectForm.classList.add('was-validated');
    const data = this.getFormData();
    if (!data.projectName)                        return false;
    if (data.overheadAllocation < 0 || data.overheadAllocation > 100) return false;
    const numericFields = ['numEngineers','engineerSalary','ceVisitCharge','visitsPerMonth',
                           'transportCost','clientPayment'];
    for (const f of numericFields) {
      if (isNaN(data[f]) || data[f] < 0) return false;
    }
    return true;
  }

  // ── Live preview ──────────────────────────────────────────────────────────
  _initLivePreview() {
    const inputs = ['numEngineers','engineerSalary','ceVisitCharge','visitsPerMonth',
                    'transportCost','clientPayment','overheadAllocation'];
    inputs.forEach(id => {
      document.getElementById(id)?.addEventListener('input', () => this._updatePreview());
    });
  }

  _updatePreview() {
    const data = this.getFormData();
    const m    = this.dataManager.calculateProjectMetrics(data);
    const show = Object.values(data).some(v => v !== '' && v !== 0 && v !== '0');
    this.el.calcPreview.classList.toggle('d-none', !show);
    if (!show) return;
    document.getElementById('prev_engineerCost').textContent = this.formatCurrency(m.engineerCost);
    document.getElementById('prev_ceVisitCost').textContent  = this.formatCurrency(m.ceVisitCost);
    document.getElementById('prev_directCost').textContent   = this.formatCurrency(m.directCost);
    document.getElementById('prev_overheadCost').textContent = this.formatCurrency(m.overheadCost);
    document.getElementById('prev_totalCost').textContent    = this.formatCurrency(m.totalCost);
    const pEl = document.getElementById('prev_profit');
    pEl.textContent = this.formatCurrency(m.profit);
    pEl.className   = 'profit-value ' + (m.profit >= 0 ? 'profit-positive' : 'profit-negative');
  }

  // ── Project detail modal ──────────────────────────────────────────────────
  showProjectDetail(index) {
    const p = this.dataManager.projects[index];
    if (!p) return;
    this._detailIndex = index;
    this.el.detailModalBody.innerHTML = `
      <div class="row g-3">
        <div class="col-12"><h5 class="text-primary">${this._esc(p.projectName)}</h5></div>
        <div class="col-md-6">
          <table class="table table-sm table-bordered">
            <tbody>
              <tr><th>Engineers</th><td>${p.numEngineers} × ${this.formatCurrency(p.engineerSalary)}</td></tr>
              <tr><th>Engineer Cost</th><td>${this.formatCurrency(p.engineerCost)}</td></tr>
              <tr><th>CE Visits</th><td>${p.visitsPerMonth} × ${this.formatCurrency(p.ceVisitCharge)}</td></tr>
              <tr><th>CE Visit Cost</th><td>${this.formatCurrency(p.ceVisitCost)}</td></tr>
              <tr><th>Transport Cost</th><td>${this.formatCurrency(p.transportCost)}</td></tr>
              <tr><th>Direct Cost</th><td>${this.formatCurrency(p.directCost)}</td></tr>
            </tbody>
          </table>
        </div>
        <div class="col-md-6">
          <table class="table table-sm table-bordered">
            <tbody>
              <tr><th>Overhead %</th><td>${p.overheadAllocation}%</td></tr>
              <tr><th>Overhead Cost</th><td>${this.formatCurrency(p.overheadCost)}</td></tr>
              <tr class="table-danger"><th>Total Cost</th><td><strong>${this.formatCurrency(p.totalCost)}</strong></td></tr>
              <tr class="table-success"><th>Client Payment</th><td><strong>${this.formatCurrency(p.clientPayment)}</strong></td></tr>
              <tr class="${(p.profit??0)>=0?'table-success':'table-danger'}">
                <th>Profit</th>
                <td><strong class="${(p.profit??0)>=0?'profit-positive':'profit-negative'}">${this.formatCurrency(p.profit)}</strong></td>
              </tr>
            </tbody>
          </table>
        </div>
        ${p.timestamp ? `<div class="col-12 text-muted small">Added: ${new Date(p.timestamp).toLocaleString()}</div>` : ''}
      </div>`;
    this.detailModal.show();
  }

  // ── Charts ────────────────────────────────────────────────────────────────
  _initCharts() {
    const base = {
      responsive: true, maintainAspectRatio: true,
      plugins: { legend: { position: 'top' } }
    };

    this.charts.costPayment = new Chart(document.getElementById('costPaymentChart'), {
      type: 'bar',
      data: {
        labels: [],
        datasets: [
          { label:'Total Cost',      data:[], backgroundColor:'rgba(255,99,132,0.7)',   borderColor:'rgba(255,99,132,1)',   borderWidth:1 },
          { label:'Client Payment',  data:[], backgroundColor:'rgba(75,192,192,0.7)',   borderColor:'rgba(75,192,192,1)',   borderWidth:1 }
        ]
      },
      options: { ...base }
    });

    this.charts.profit = new Chart(document.getElementById('profitChart'), {
      type: 'bar',
      data: {
        labels: [],
        datasets: [{ label:'Profit', data:[], backgroundColor:[], borderColor:[], borderWidth:1 }]
      },
      options: { ...base }
    });

    this.charts.trend = new Chart(document.getElementById('trendChart'), {
      type: 'line',
      data: {
        labels: [],
        datasets: [{
          label: 'Monthly Profit', data:[],
          borderColor: 'rgba(54,162,235,1)', backgroundColor: 'rgba(54,162,235,0.2)',
          tension: 0.4, fill: true
        }]
      },
      options: { ...base }
    });

    this.charts.overhead = new Chart(document.getElementById('overheadChart'), {
      type: 'scatter',
      data: {
        datasets: [{
          label: 'Overhead % vs Profit', data:[],
          backgroundColor:'rgba(153,102,255,0.7)', borderColor:'rgba(153,102,255,1)', pointRadius:6
        }]
      },
      options: {
        ...base,
        scales: {
          x: { title:{ display:true, text:'Overhead %' } },
          y: { title:{ display:true, text:'Profit (LKR)' } }
        }
      }
    });
  }

  updateCharts() {
    const projects  = this.dataManager.getProjects();
    const labels    = projects.map(p => p.projectName);
    const trendData = this.dataManager.getMonthlyTrend();

    // Cost vs Payment
    const cp = this.charts.costPayment;
    cp.data.labels                  = labels;
    cp.data.datasets[0].data        = projects.map(p => p.totalCost    ?? 0);
    cp.data.datasets[1].data        = projects.map(p => p.clientPayment ?? 0);
    cp.update('none');

    // Profit bars (green / red)
    const pr = this.charts.profit;
    pr.data.labels                          = labels;
    pr.data.datasets[0].data               = projects.map(p => p.profit ?? 0);
    pr.data.datasets[0].backgroundColor    = projects.map(p => (p.profit??0)>=0?'rgba(75,192,192,0.7)':'rgba(255,99,132,0.7)');
    pr.data.datasets[0].borderColor        = projects.map(p => (p.profit??0)>=0?'rgba(75,192,192,1)':'rgba(255,99,132,1)');
    pr.update('none');

    // Trend
    const tr = this.charts.trend;
    tr.data.labels              = trendData.map(d => d.period);
    tr.data.datasets[0].data   = trendData.map(d => d.profit);
    tr.update('none');

    // Scatter
    const oh = this.charts.overhead;
    oh.data.datasets[0].data = projects.map(p => ({ x: p.overheadAllocation??0, y: p.profit??0 }));
    oh.update('none');
  }

  // ── Toast ─────────────────────────────────────────────────────────────────
  showToast(message, type = 'info') {
    const colours = { success:'bg-success', danger:'bg-danger', warning:'bg-warning', info:'bg-info' };
    this.el.toastHeader.className = `toast-header ${colours[type] ?? 'bg-info'} text-white`;
    this.el.toastMessage.innerHTML = message; // allow HTML for rich details
    bootstrap.Toast.getOrCreateInstance(this.el.toast, { delay: 4000 }).show();
  }

  // ── Misc ──────────────────────────────────────────────────────────────────
  formatCurrency(amount) {
    const n = parseFloat(amount) || 0;
    return 'LKR ' + n.toLocaleString('en-US', { minimumFractionDigits:2, maximumFractionDigits:2 });
  }

  printDashboard() {
    const { month, year } = this.dataManager.getCurrentPeriod();
    const orig = document.title;
    document.title = `Profit Dashboard – ${month}/${year}`;
    window.print();
    document.title = orig;
  }

  refreshDashboard() {
    this.updateStats();
    this.updateTable();
    this.updatePeriodDisplay();
    this.updateCharts();
  }

  /** Escape HTML to prevent XSS in dynamic content */
  _esc(str) {
    return String(str ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  }

  getDetailIndex() { return this._detailIndex; }
}

// ────────────────────────────────────────────────────────────────────────────
// UpdateController  (only active inside Electron)
// ────────────────────────────────────────────────────────────────────────────
class UpdateController {
  constructor(ui) {
    this.ui = ui;
    if (typeof window.electronAPI === 'undefined' || !window.electronAPI.onUpdaterStatus) return;

    window.electronAPI.onUpdaterStatus(payload => this._handleStatus(payload));

    // Show app version
    window.electronAPI.getAppVersion().then(v => {
      if (v) ui.el.appVersionBadge.textContent = 'v' + v;
    }).catch(() => {});

    ui.el.checkUpdateBtn.addEventListener('click', () => {
      window.electronAPI.checkForUpdates();
      ui.showToast('Checking for updates…', 'info');
    });

    ui.el.updateDownloadBtn.addEventListener('click', () => {
      window.electronAPI.downloadUpdate();
      ui.el.updateDownloadBtn.disabled = true;
      ui.el.updateProgress.classList.remove('d-none');
    });

    ui.el.updateInstallBtn.addEventListener('click', () => {
      window.electronAPI.installUpdate();
    });
  }

  _handleStatus({ event, version, percent, message }) {
    const { updateBanner, updateBannerText, updateDownloadBtn, updateInstallBtn,
            updateProgress, updateProgressBar } = this.ui.el;

    switch (event) {
      case 'available':
        updateBannerText.textContent = `Version ${version} is available.`;
        updateBanner.classList.remove('d-none');
        break;
      case 'progress':
        updateProgressBar.style.width = percent + '%';
        updateProgressBar.textContent = percent + '%';
        break;
      case 'downloaded':
        updateDownloadBtn.classList.add('d-none');
        updateInstallBtn.classList.remove('d-none');
        updateProgress.classList.add('d-none');
        updateBannerText.textContent = `Version ${version} downloaded. Ready to install.`;
        break;
      case 'error':
        this.ui.showToast('Update check failed: ' + (message || 'Unknown error'), 'warning');
        break;
    }
  }
}

// ────────────────────────────────────────────────────────────────────────────
// AppController  – wires everything together
// ────────────────────────────────────────────────────────────────────────────
class AppController {
  constructor() {
    this.data    = new DataManager();
    this.ui      = new UIController(this.data);
    this.updater = new UpdateController(this.ui);
    this._init();
  }

  _init() {
    this.ui.populateYearSelect();
    this.ui.setCurrentMonth();
    this._bindEvents();
  }

  _bindEvents() {
    // Load period
    this.ui.el.loadMonthBtn.addEventListener('click', () => this._loadPeriod());

    // Form submit (add or update)
    this.ui.el.projectForm.addEventListener('submit', e => { e.preventDefault(); this._handleSubmit(); });

    // Reset / cancel edit
    this.ui.el.resetFormBtn.addEventListener('click',   () => this.ui.resetForm());
    this.ui.el.cancelEditBtn.addEventListener('click',  () => this.ui.exitEditMode());

    // Dashboard actions
    this.ui.el.backupBtn.addEventListener('click',         () => this._createBackup());
    this.ui.el.printDashboardBtn.addEventListener('click', () => this.ui.printDashboard());
    this.ui.el.exportExcelBtn.addEventListener('click',    () => this._exportExcel());

    // Edit from detail modal
    this.ui.el.editFromDetailBtn.addEventListener('click', () => {
      const idx = this.ui.getDetailIndex();
      if (idx !== null) {
        this.ui.detailModal.hide();
        this.editProject(idx);
      }
    });
  }

  // ── Period load ───────────────────────────────────────────────────────────
  async _loadPeriod() {
    const month = this.ui.el.monthSelect.value;
    const year  = this.ui.el.yearSelect.value;
    this.data.setCurrentPeriod(month, year);
    this.ui.showLoading(true);
    try {
      await this.data.loadProjects();
      this.ui.showMainContent();
      this.ui.refreshDashboard();
      this.ui.showToast(`Loaded data for ${month}/${year}`, 'success');
    } catch (err) {
      console.error(err);
      this.ui.showToast('Error loading data: ' + err.message, 'danger');
    } finally {
      this.ui.showLoading(false);
    }
  }

  // ── Form submit ───────────────────────────────────────────────────────────
  async _handleSubmit() {
    if (!this.ui.validateForm()) {
      this.ui.showToast('Please fill all required fields correctly.', 'warning');
      return;
    }
    const formData    = this.ui.getFormData();
    const editingName = this.ui.getEditingName();
    this.ui.showLoading(true);

    try {
      if (editingName) {
        await this.data.updateProject(editingName, formData);
        this.ui.exitEditMode();
        this.ui.showToast(`Project "${formData.projectName}" updated.`, 'success');
      } else {
        await this.data.addProject(formData);
        this.ui.resetForm();
        this.ui.showToast(`Project "${formData.projectName}" added.`, 'success');
      }
      this.ui.refreshDashboard();
    } catch (err) {
      console.error(err);
      this.ui.showToast('Save error: ' + err.message, 'danger');
    } finally {
      this.ui.showLoading(false);
    }
  }

  // ── Public: called from inline onclick in table rows ─────────────────────
  viewProject(index) {
    this.ui.showProjectDetail(index);
  }

  editProject(index) {
    const p = this.data.projects[index];
    if (!p) return;
    this.ui.fillForm(p);
    this.ui.enterEditMode(p.projectName);
  }

  async deleteProject(index) {
    const p = this.data.projects[index];
    if (!p) return;

    const confirmed = await this._confirm(
      `Delete project "${p.projectName}"? This cannot be undone.`
    );
    if (!confirmed) return;

    this.ui.showLoading(true);
    try {
      await this.data.deleteProject(index);
      this.ui.refreshDashboard();
      this.ui.showToast('Project deleted.', 'success');
    } catch (err) {
      console.error(err);
      this.ui.showToast('Delete error: ' + err.message, 'danger');
    } finally {
      this.ui.showLoading(false);
    }
  }

  // ── Backup ────────────────────────────────────────────────────────────────
  async _createBackup() {
    this.ui.showLoading(true);
    try {
      const backupPath = await this.data.createBackup();
      this.ui.showToast('Backup created' + (backupPath ? ': ' + backupPath : '.'), 'success');
    } catch (err) {
      console.error(err);
      this.ui.showToast('Backup error: ' + err.message, 'danger');
    } finally {
      this.ui.showLoading(false);
    }
  }

  // ── Export ────────────────────────────────────────────────────────────────
  async _exportExcel() {
    this.ui.showLoading(true);
    try {
      const res = await this.data.exportToExcel();
      if (res?.canceled) {
        this.ui.showToast('Export cancelled.', 'info');
      } else {
        this.ui.showToast('Exported to: ' + res.path, 'success');
      }
    } catch (err) {
      console.error(err);
      this.ui.showToast('Export error: ' + err.message, 'danger');
    } finally {
      this.ui.showLoading(false);
    }
  }

  // ── Native/fallback confirm ────────────────────────────────────────────────
  async _confirm(message) {
    if (window.electronAPI?.showMessage) {
      const res = await window.electronAPI.showMessage({
        type: 'question',
        buttons: ['Cancel', 'Delete'],
        defaultId: 0,
        cancelId: 0,
        message
      });
      return res.response === 1;
    }
    return window.confirm(message);
  }
}

// ────────────────────────────────────────────────────────────────────────────
// Bootstrap
// ────────────────────────────────────────────────────────────────────────────
// APP is on window so inline onclick handlers can reach it
document.addEventListener('DOMContentLoaded', () => {
  window.APP = new AppController();
});
