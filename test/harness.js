// GAS モック + 結合テストハーネス
// 実行: node test/harness.js(本体)/ node test/harness.js test/tests_setup.js(セットアップ)
const fs = require('fs');

function pad(n){return ('0'+n).slice(-2);}
global.Utilities = { formatDate: (d, tz, fmt) => fmt
  .replace('yyyy', d.getFullYear()).replace('MM', pad(d.getMonth()+1)).replace('dd', pad(d.getDate()))
  .replace('HH', pad(d.getHours())).replace('mm', pad(d.getMinutes())).replace('ss', pad(d.getSeconds())) };
global.Session = { getActiveUser: () => ({ getEmail: () => 'member01@example.com' }) };

class MockRange {
  constructor(sheet, row, col, numRows = 1, numCols = 1) {
    this.s = sheet; this.r = row; this.c = col; this.nr = numRows; this.nc = numCols;
  }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = [];
      for (let j = 0; j < this.nc; j++) {
        const v = (this.s.data[this.r - 1 + i] || [])[this.c - 1 + j];
        row.push(v === undefined ? '' : v);
      }
      out.push(row);
    }
    return out;
  }
  getValue() { return this.getValues()[0][0]; }
  setValues(vals) {
    for (let i = 0; i < this.nr; i++) {
      while (this.s.data.length < this.r + i) this.s.data.push([]);
      for (let j = 0; j < this.nc; j++) this.s.data[this.r - 1 + i][this.c - 1 + j] = vals[i][j];
    }
    return this;
  }
  setValue(v) { return this.setValues([[v]]); }
  clearContent() {
    for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) {
      if (this.s.data[this.r - 1 + i]) this.s.data[this.r - 1 + i][this.c - 1 + j] = '';
    }
    return this;
  }
  clear() { return this.clearContent(); }
  setNumberFormat() { return this; } setNumberFormats() { return this; }
  setFontWeight() { return this; } setBackground() { return this; } setFontColor() { return this; }
  setDataValidation() { return this; } insertCheckboxes() { return this; }
  getA1Notation() { return 'B1'; }
  getSheet() { return this.s; }
}

class MockSheet {
  constructor(name) { this.name = name; this.data = []; this.hidden = false; }
  getName() { return this.name; }
  getRange(a, b, c, d) {
    if (typeof a === 'string') { // 'B1' など最低限
      const m = a.match(/^([A-Z]+)(\d+)$/);
      const col = m[1].split('').reduce((s, ch) => s * 26 + ch.charCodeAt(0) - 64, 0);
      return new MockRange(this, Number(m[2]), col, 1, 1);
    }
    return new MockRange(this, a, b, c || 1, d || 1);
  }
  getDataRange() {
    const w = Math.max(1, ...this.data.map(r => r.length));
    return new MockRange(this, 1, 1, Math.max(this.data.length, 1), w);
  }
  appendRow(row) { this.data.push(row.slice()); return this; }
  getLastRow() { return this.data.length; }
  getLastColumn() { return Math.max(1, ...this.data.map(r => r.length)); }
  getMaxRows() { return Math.max(this.data.length, 1000); }
  setFrozenRows() {} hideSheet() { this.hidden = true; }
  setColumnWidth() {}
  getProtections() { return []; }
  protect() {
    const p = {};
    ['setDescription','setWarningOnly','removeEditors','addEditors','setUnprotectedRanges'].forEach(m => p[m] = () => p);
    p.getEditors = () => [];
    return p;
  }
}

const SHEETS = {};
global.SpreadsheetApp = {
  getActiveSpreadsheet: () => ({
    getSheetByName: n => SHEETS[n] || null,
    insertSheet: n => (SHEETS[n] = new MockSheet(n)),
    getSheets: () => Object.values(SHEETS),
    toast: (msg) => console.log('[toast]', msg),
    deleteSheet: () => {},
  }),
  getUi: () => { throw new Error('no ui in test'); },
  newDataValidation: () => ({ requireValueInList: () => ({ setAllowInvalid: () => ({ build: () => ({}) }) }) }),
  ProtectionType: { SHEET: 'SHEET' },
};
global.HtmlService = { createHtmlOutputFromFile: () => ({ setTitle: () => ({}) }) };

// ---- ソース読込 + テストを単一スコープで実行 ----
const path = require('path');
const ROOT = path.join(__dirname, '..');
const SRC = ['00_const.js','01_util.js','02_setup.js','03_menu.js','10_kintone_import.js','20_recalc.js','30_form_server.js','31_admin.js','40_views.js']
  .map(f => fs.readFileSync(path.join(ROOT, 'src', f), 'utf8')).join('\n');
const TESTS = fs.readFileSync(process.argv[2] || path.join(__dirname, 'tests.js'), 'utf8');
eval(SRC + '\n' + TESTS);
