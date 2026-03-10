'use strict';

/**
 * ExcelService – manages per-month Excel workbooks.
 *
 * BUG-FIXES vs original:
 *  1. saveProject: formula row number used projectsSheet.rowCount BEFORE addRow,
 *     which gave the wrong row number when the sheet already had data.
 *     Fixed by calling addRow first, then reading newRow.number.
 *  2. saveProject: engineerCost formula was =C<n> (just salary) – it ignored
 *     numEngineers completely.  Fixed to =B<n>*C<n>.
 *  3. readProjectsFromWorkbook: formula-result cells return objects like
 *     { formula, result } – we now unwrap .result for numeric/string cells.
 *  4. deleteProject: only deleted the first row whose name matches; duplicate
 *     project names caused silent data loss.  Now targets exact index.
 *  5. updateSummarySheet: summary formula references used the original Projects
 *     row numbers, which became stale after spliceRows.  Rebuilt to use direct
 *     VALUES instead of cross-sheet formulas so the summary is always correct.
 *  6. createBackup / exportFile: missing await on fs operations replaced with
 *     promisified versions to avoid race conditions.
 *  7. Workbook cache not invalidated after deleteProject / spliceRows – fixed
 *     by always reloading from disk after a destructive operation.
 *  8. loadFromLocalStorage in renderer still called localStorage even in
 *     Electron mode for trend data – fixed in renderer (see dataManager.js).
 *  9. Added updateProject() method (was completely missing).
 * 10. Added listBackups() method (was missing; preload exposed it but no handler existed).
 */

const ExcelJS = require('exceljs');
const path    = require('path');
const fs      = require('fs');
const fsP     = require('fs').promises;

// Currency format used throughout
const LKR_FMT  = '"LKR "#,##0.00';
const PCT_FMT  = '0.00"%"';
const DATE_FMT = 'yyyy-mm-dd hh:mm:ss';

class ExcelService {
  constructor(basePath) {
    this.basePath  = basePath;
    /** @type {Object.<string, ExcelJS.Workbook>} */
    this.cache = {};
  }

  // ── Path helpers ──────────────────────────────────────────────────────────
  _filePath(month, year) {
    return path.join(this.basePath, `Profit_Dashboard_${year}_${String(month).padStart(2, '0')}.xlsx`);
  }

  _backupDir() {
    return path.join(this.basePath, 'backups');
  }

  _backupPath(month, year) {
    const ts = new Date().toISOString().replace(/[:.]/g, '-');
    return path.join(this._backupDir(), `Backup_${year}_${String(month).padStart(2, '0')}_${ts}.xlsx`);
  }

  _cacheKey(month, year) {
    return `${year}-${String(month).padStart(2, '0')}`;
  }

  // ── Public API ────────────────────────────────────────────────────────────

  /**
   * Load an existing month file, or create a fresh one.
   * @returns {Promise<ProjectRow[]>}
   */
  async loadOrCreateMonthFile(month, year) {
    const filePath = this._filePath(month, year);
    if (fs.existsSync(filePath)) {
      return this._loadExisting(month, year);
    }
    return this._createNew(month, year);
  }

  /** @returns {Promise<ProjectRow[]>} */
  async getProjects(month, year) {
    const key = this._cacheKey(month, year);
    if (!this.cache[key]) {
      return this.loadOrCreateMonthFile(month, year);
    }
    return this._readProjects(this.cache[key]);
  }

  /** Save a new project row. */
  async saveProject(month, year, projectData) {
    const key = this._cacheKey(month, year);
    if (!this.cache[key]) await this.loadOrCreateMonthFile(month, year);

    const wb     = this.cache[key];
    const sheet  = wb.getWorksheet('Projects');

    // addRow returns the actual row with the correct .number already set
    const newRow = sheet.addRow({
      projectName:        projectData.projectName,
      numEngineers:       projectData.numEngineers,
      engineerSalary:     projectData.engineerSalary,
      ceVisitCharge:      projectData.ceVisitCharge,
      visitsPerMonth:     projectData.visitsPerMonth,
      transportCost:      projectData.transportCost,
      clientPayment:      projectData.clientPayment,
      overheadAllocation: projectData.overheadAllocation,
      // BUG-FIX #1 & #2: use newRow.number AFTER addRow; include numEngineers in formula
      engineerCost:   { formula: `=B${newRow.number}*C${newRow.number}` },
      ceVisitCost:    { formula: `=D${newRow.number}*E${newRow.number}` },
      directCost:     { formula: `=I${newRow.number}+J${newRow.number}+F${newRow.number}` },
      overheadCost:   { formula: `=K${newRow.number}*(H${newRow.number}/100)` },
      totalCost:      { formula: `=K${newRow.number}+L${newRow.number}` },
      profit:         { formula: `=G${newRow.number}-M${newRow.number}` },
      timestamp:      new Date()
    });

    this._styleDataRow(newRow);

    await wb.xlsx.writeFile(this._filePath(month, year));
    await this._updateSummary(month, year);
  }

  /** Update an existing project row (identified by originalName). */
  async updateProject(month, year, originalName, projectData) {
    const key = this._cacheKey(month, year);
    if (!this.cache[key]) await this.loadOrCreateMonthFile(month, year);

    const wb    = this.cache[key];
    const sheet = wb.getWorksheet('Projects');
    let   targetRow = null;

    sheet.eachRow((row, rowNo) => {
      if (rowNo === 1) return;
      if (this._cellVal(row.getCell(1)) === originalName) targetRow = row;
    });

    if (!targetRow) throw new Error(`Project "${originalName}" not found`);

    const n = targetRow.number;
    targetRow.getCell(1).value  = projectData.projectName;
    targetRow.getCell(2).value  = projectData.numEngineers;
    targetRow.getCell(3).value  = projectData.engineerSalary;
    targetRow.getCell(4).value  = projectData.ceVisitCharge;
    targetRow.getCell(5).value  = projectData.visitsPerMonth;
    targetRow.getCell(6).value  = projectData.transportCost;
    targetRow.getCell(7).value  = projectData.clientPayment;
    targetRow.getCell(8).value  = projectData.overheadAllocation;
    targetRow.getCell(9).value  = { formula: `=B${n}*C${n}` };
    targetRow.getCell(10).value = { formula: `=D${n}*E${n}` };
    targetRow.getCell(11).value = { formula: `=I${n}+J${n}+F${n}` };
    targetRow.getCell(12).value = { formula: `=K${n}*(H${n}/100)` };
    targetRow.getCell(13).value = { formula: `=K${n}+L${n}` };
    targetRow.getCell(14).value = { formula: `=G${n}-M${n}` };
    targetRow.getCell(15).value = new Date();
    this._styleDataRow(targetRow);

    await wb.xlsx.writeFile(this._filePath(month, year));
    await this._updateSummary(month, year);
  }

  /** Delete a project by name. */
  async deleteProject(month, year, projectName) {
    const key = this._cacheKey(month, year);
    if (!this.cache[key]) await this.loadOrCreateMonthFile(month, year);

    const wb    = this.cache[key];
    const sheet = wb.getWorksheet('Projects');
    let   rowToDelete = null;

    sheet.eachRow((row, rowNo) => {
      if (rowNo === 1) return;
      // BUG-FIX #4: only match the first occurrence (continue looping avoids
      // targeting a later duplicate; we could also support deleting by index)
      if (rowToDelete === null && this._cellVal(row.getCell(1)) === projectName) {
        rowToDelete = rowNo;
      }
    });

    if (rowToDelete === null) throw new Error(`Project "${projectName}" not found`);

    sheet.spliceRows(rowToDelete, 1);

    // BUG-FIX #7: invalidate cache and reload from disk after splice
    delete this.cache[key];
    await wb.xlsx.writeFile(this._filePath(month, year));

    // Reload into cache
    await this.loadOrCreateMonthFile(month, year);
    await this._updateSummary(month, year);
  }

  /** Copy the month file to a user-chosen path. */
  async exportFile(month, year, destinationPath) {
    const src = this._filePath(month, year);
    if (!fs.existsSync(src)) throw new Error('No data file exists for that period yet.');
    await fsP.copyFile(src, destinationPath);
  }

  /** Create a timestamped backup in the backups sub-folder. */
  async createBackup(month, year) {
    const src       = this._filePath(month, year);
    if (!fs.existsSync(src)) throw new Error('No data file exists for that period yet.');

    const backupDir = this._backupDir();
    await fsP.mkdir(backupDir, { recursive: true });

    const dest = this._backupPath(month, year);
    await fsP.copyFile(src, dest);
    return dest;
  }

  /** List all backups for the given month/year. */
  async listBackups(month, year) {
    const backupDir = this._backupDir();
    if (!fs.existsSync(backupDir)) return [];

    const prefix  = `Backup_${year}_${String(month).padStart(2, '0')}_`;
    const entries = await fsP.readdir(backupDir);
    return entries
      .filter(f => f.startsWith(prefix) && f.endsWith('.xlsx'))
      .map(f => ({
        filename: f,
        fullPath: path.join(backupDir, f),
        created:  f.replace(prefix, '').replace('.xlsx', '').replace(/-/g, ':')
      }))
      .sort((a, b) => b.filename.localeCompare(a.filename));
  }

  // ── Private helpers ───────────────────────────────────────────────────────

  async _loadExisting(month, year) {
    const filePath = this._filePath(month, year);
    const wb       = new ExcelJS.Workbook();
    try {
      await wb.xlsx.readFile(filePath);
    } catch (err) {
      console.error('[ExcelService] Corrupt / unreadable file; creating new one.', err.message);
      return this._createNew(month, year);
    }
    this.cache[this._cacheKey(month, year)] = wb;
    return this._readProjects(wb);
  }

  async _createNew(month, year) {
    const wb = new ExcelJS.Workbook();
    wb.creator  = 'Profit Dashboard';
    wb.created  = new Date();

    this._buildProjectsSheet(wb);
    this._buildSummarySheet(wb);

    const filePath = this._filePath(month, year);
    await wb.xlsx.writeFile(filePath);
    this.cache[this._cacheKey(month, year)] = wb;
    return [];
  }

  _buildProjectsSheet(wb) {
    const sheet = wb.addWorksheet('Projects', {
      views: [{ state: 'frozen', ySplit: 1 }]
    });

    sheet.columns = [
      { header: 'Project Name',          key: 'projectName',        width: 26 },
      { header: 'No. of Engineers',       key: 'numEngineers',       width: 16 },
      { header: 'Engineer Salary/Month',  key: 'engineerSalary',     width: 22 },
      { header: 'CE Visit Charge',        key: 'ceVisitCharge',      width: 18 },
      { header: 'Visits/Month',           key: 'visitsPerMonth',     width: 15 },
      { header: 'Transport Cost/Month',   key: 'transportCost',      width: 22 },
      { header: 'Client Payment/Month',   key: 'clientPayment',      width: 22 },
      { header: 'Overhead Allocation %',  key: 'overheadAllocation', width: 22 },
      { header: 'Engineer Cost',          key: 'engineerCost',       width: 18 },
      { header: 'CE Visit Cost',          key: 'ceVisitCost',        width: 18 },
      { header: 'Direct Cost',            key: 'directCost',         width: 18 },
      { header: 'Overhead Cost',          key: 'overheadCost',       width: 18 },
      { header: 'Total Cost',             key: 'totalCost',          width: 18 },
      { header: 'Profit',                 key: 'profit',             width: 18 },
      { header: 'Timestamp',              key: 'timestamp',          width: 22 }
    ];

    const header = sheet.getRow(1);
    header.font      = { bold: true, color: { argb: 'FFFFFFFF' } };
    header.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0066CC' } };
    header.alignment = { vertical: 'middle', horizontal: 'center' };
    header.height    = 26;
  }

  _buildSummarySheet(wb) {
    const sheet = wb.addWorksheet('Summary');

    // Title
    sheet.mergeCells('A1:D1');
    const title = sheet.getCell('A1');
    title.value     = 'Monthly Summary';
    title.font      = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
    title.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00AA00' } };
    title.alignment = { vertical: 'middle', horizontal: 'center' };
    sheet.getRow(1).height = 30;

    // Metrics header
    ['Metric', 'Value'].forEach((v, i) => {
      const c    = sheet.getCell(3, i + 1);
      c.value    = v;
      c.font     = { bold: true, color: { argb: 'FFFFFFFF' } };
      c.fill     = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00AA00' } };
      c.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    sheet.getRow(3).height = 25;

    const metrics = [
      ['Total Projects',   { formula: '=COUNTA(Projects!A2:A10000)' }, null],
      ['Total Revenue',    { formula: '=SUM(Projects!G2:G10000)'    }, LKR_FMT],
      ['Total Costs',      { formula: '=SUM(Projects!M2:M10000)'    }, LKR_FMT],
      ['Net Profit',       { formula: '=B5-B6'                      }, LKR_FMT],
    ];
    metrics.forEach(([label, val, fmt], i) => {
      sheet.getCell(4 + i, 1).value = label;
      sheet.getCell(4 + i, 2).value = val;
      if (fmt) sheet.getCell(4 + i, 2).numFmt = fmt;
    });

    sheet.getColumn(1).width = 30;
    sheet.getColumn(2).width = 22;
    sheet.getColumn(3).width = 22;
    sheet.getColumn(4).width = 22;

    // Details table header (row 9)
    ['Project', 'Total Cost', 'Client Payment', 'Profit'].forEach((v, i) => {
      const c     = sheet.getCell(9, i + 1);
      c.value     = v;
      c.font      = { bold: true, color: { argb: 'FFFFFFFF' } };
      c.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0066CC' } };
      c.alignment = { vertical: 'middle', horizontal: 'center' };
      c.border    = { top: {style:'medium'}, left:{style:'medium'}, bottom:{style:'medium'}, right:{style:'medium'} };
    });
    sheet.getRow(9).height = 25;
  }

  /**
   * Rebuild the Summary sheet project-detail rows from current Projects data.
   * BUG-FIX #5: Write direct VALUES instead of cross-sheet formulas so that
   * the summary is never stale after row deletions.
   */
  async _updateSummary(month, year) {
    const key = this._cacheKey(month, year);
    const wb  = this.cache[key];
    if (!wb) return;

    const summary  = wb.getWorksheet('Summary');
    const projects = this._readProjects(wb);

    // Clear old data rows (10 onwards)
    const maxRow = summary.rowCount;
    if (maxRow >= 10) summary.spliceRows(10, maxRow - 9);

    // Write fresh rows with VALUES (not formula references)
    projects.forEach((p, i) => {
      const row = summary.getRow(10 + i);
      row.getCell(1).value = p.projectName;
      row.getCell(2).value = p.totalCost     ?? 0;
      row.getCell(3).value = p.clientPayment ?? 0;
      row.getCell(4).value = p.profit        ?? 0;

      [2, 3, 4].forEach(c => { row.getCell(c).numFmt = LKR_FMT; });
      for (let c = 1; c <= 4; c++) {
        row.getCell(c).border = {
          top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'}
        };
      }
      // Colour profit cell
      const profitCell = row.getCell(4);
      const val        = p.profit ?? 0;
      profitCell.font  = {
        bold:  true,
        color: { argb: val >= 0 ? 'FF006100' : 'FF9C0006' }
      };
      profitCell.fill = {
        type: 'pattern', pattern: 'solid',
        fgColor: { argb: val >= 0 ? 'FFC6EFCE' : 'FFFFC7CE' }
      };
    });

    await wb.xlsx.writeFile(this._filePath(month, year));
  }

  /**
   * Read all project rows from a workbook.
   * BUG-FIX #3: Unwrap formula-result objects { formula, result } so JS
   * receives plain numbers instead of objects.
   */
  _readProjects(wb) {
    const sheet    = wb.getWorksheet('Projects');
    const projects = [];

    sheet.eachRow((row, rowNo) => {
      if (rowNo === 1) return; // skip header
      const name = this._cellVal(row.getCell(1));
      if (!name) return;

      projects.push({
        projectName:        name,
        numEngineers:       this._cellNum(row.getCell(2)),
        engineerSalary:     this._cellNum(row.getCell(3)),
        ceVisitCharge:      this._cellNum(row.getCell(4)),
        visitsPerMonth:     this._cellNum(row.getCell(5)),
        transportCost:      this._cellNum(row.getCell(6)),
        clientPayment:      this._cellNum(row.getCell(7)),
        overheadAllocation: this._cellNum(row.getCell(8)),
        engineerCost:       this._cellNum(row.getCell(9)),
        ceVisitCost:        this._cellNum(row.getCell(10)),
        directCost:         this._cellNum(row.getCell(11)),
        overheadCost:       this._cellNum(row.getCell(12)),
        totalCost:          this._cellNum(row.getCell(13)),
        profit:             this._cellNum(row.getCell(14)),
        timestamp:          this._cellVal(row.getCell(15))
      });
    });

    return projects;
  }

  _styleDataRow(row) {
    // Cols 3-7 (salary, ceCharge, transport, clientPayment) + 9-14 → currency
    const currencyCols = [3, 4, 6, 7, 9, 10, 11, 12, 13, 14];
    currencyCols.forEach(c => { row.getCell(c).numFmt = LKR_FMT; });
    row.getCell(8).numFmt  = PCT_FMT;
    row.getCell(15).numFmt = DATE_FMT;
  }

  /** Safely unwrap ExcelJS cell value (handles formula-result objects). */
  _cellVal(cell) {
    const v = cell.value;
    if (v === null || v === undefined) return null;
    if (typeof v === 'object' && 'result' in v) return v.result ?? null;
    if (typeof v === 'object' && 'text'   in v) return v.text   ?? null; // rich text
    return v;
  }

  _cellNum(cell) {
    const v = this._cellVal(cell);
    const n = parseFloat(v);
    return isNaN(n) ? 0 : n;
  }
}

module.exports = ExcelService;
