/**
 * Comprehensive Exception & Anti-Hallucination Contract Tests for InvoiceFlow E2E
 * Verifies all rules from "Mo ta Cac buoc xuat hoa don TMV (1).xlsx" and fail-closed guarantees.
 */
const fs = require('fs');
const assert = require('assert');
const crypto = require('crypto');
const path = require('path');

const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
const code02 = fs.readFileSync(path.join(__dirname, '../code/02_config.js'), 'utf8');
const code09 = fs.readFileSync(path.join(__dirname, '../code/09_engine.js'), 'utf8');

function colNum(x) {
  if (typeof x === 'number') return x;
  return String(x).toUpperCase().split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
}
class Cell {
  constructor(row, col) { this.row = row; this.col = col; this.style = {}; }
  get value() { return this.row.data[this.col] ?? null; }
  set value(v) { this.row.data[this.col] = v; }
}
class Row {
  constructor(n, vals = {}) { this.number = n; this.data = { ...vals }; this.cells = {}; this.height = 20; }
  getCell(col) { col = colNum(col); return this.cells[col] || (this.cells[col] = new Cell(this, col)); }
  get hasValues() { return Object.values(this.data).some(v => v != null && v !== ''); }
}
class Sheet {
  constructor(data) { this.rows = {}; for (const [r, values] of Object.entries(data)) this.rows[r] = new Row(+r, values); }
  getRow(r) { return this.rows[r] || (this.rows[r] = new Row(r)); }
  addRow(vals) { const next = (this.lastRow ? this.lastRow.number : 0) + 1; return this.rows[next] = new Row(next, vals); }
  getCell(addr) { const m = String(addr).match(/^([A-Z]+)(\d+)$/); return this.getRow(+m[2]).getCell(m[1]); }
  eachRow(opts, fn) { if (typeof opts === 'function') fn = opts; for (const r of Object.values(this.rows).sort((a, b) => a.number - b.number)) if (r.hasValues) fn(r, r.number); }
  get lastRow() { return Object.values(this.rows).sort((a, b) => b.number - a.number)[0]; }
}
class Workbook {
  constructor() { this.calcProperties = {}; this.sheets = []; }
  get worksheets() { return this.sheets || []; }
  addWorksheet(name) { const s = Object.assign(new Sheet({}), { name }); this.sheets.push(s); return s; }
  getWorksheet(name) { return this.worksheets.find(s => s.name === name); }
  xlsx = {
    load: async (b) => {
      const raw = JSON.parse(Buffer.from(b).toString());
      this.sheets = Object.entries(raw).map(([name, rows]) => Object.assign(new Sheet(rows), { name }));
    },
    writeBuffer: async () => Buffer.from('XLSX_TEST_BYTES')
  };
}
const fakeRequire = (id) => id === 'exceljs' ? { Workbook } : id === 'crypto' ? crypto : require(id);

function makeCustoms(rows) {
  const sheet = { '1': { '1': 'DIFF', '2': 'DECL', '3': 'INV', '4': 'DATE' } };
  rows.forEach((r, idx) => { sheet[String(idx + 2)] = r; });
  return { Sheet1: sheet };
}
function makeTemplate() {
  return { Sheet1: { '11': { '4': 'D', '5': 'E', '6': 'F', '7': 'G', '8': 'H', '22': 'V' } } };
}
const baseCfg = {
  companyCode: 'HK63', date: '2026-08-14', requestedBy: 'test', profile: 'FINANCE_VN',
  sapSourceMode: 'LOCAL_FOLDER', customsPath: '/files/finance/source', templatePath: '/files/finance/template',
  outputFolder: '/files/finance/results', sapInvoiceUrl: 'https://sap.local/invoices', sapRateUrl: 'https://sap.local/rate',
  customsMappingVerified: true, sapContractVerified: true, fxContractVerified: true, templateMappingVerified: true,
  customs: { sheetName: '', headerRow: 1, colInvoice: 'C', colDate: 'D', colDeclaration: 'B', colCheckDifference: 'A', colAmount: '' },
  sap: { rowsPath: 'value', invoice: 'invoice', date: 'billingDate', F: 'contentF', G: 'contentG', K: 'fieldK', L: 'quantity', M: 'fieldM', N: 'fieldN', P: 'declarationText', reconcileDifference: 'reconcileDifference' },
  fx: { rate: 'rateVndPerUsd', validOn: 'validOn' }, templateSheet: '',
  toleranceCustoms: 0.03, toleranceSAP: 0.000001,
  requireCustomsDifference: true, requireSAPDifference: true, requireCustomsDate: true, requireFXValidOn: true, maxRows: 20000
};

async function runEngine(customsRows, localSapOverrides = {}) {
  const params = {
    cfg: JSON.parse(JSON.stringify(baseCfg)),
    customs: { fileName: 'Baocaotokhai.xlsx', base64: Buffer.from(JSON.stringify(makeCustoms(customsRows))).toString('base64'), candidates: 1, selection: 'ONLY_MATCH' },
    template: { fileName: 'template.xlsx', base64: Buffer.from(JSON.stringify(makeTemplate())).toString('base64') }
  };
  const localSap = {
    sapSourceMode: 'LOCAL_FOLDER',
    workFilePath: '',
    skippedCounts: { canceled: 0, wrongSalesOrg: 0, wrongBillingType: 0, wrongPlant: 0, wrongCustomer: 0 },
    missingPirMaterials: [],
    droppedZeroRows: [],
    sapPayload: { value: [] },
    fxPayload: { rateVndPerUsd: 25890, validOn: '2026-08-13' },
    ...localSapOverrides
  };
  const all = { '06 - Buffer template': [{ json: params }] };
  const dollar = (name) => ({ first: () => all[name][0], all: () => all[name] });
  const self = { helpers: { prepareBinaryData: async (buffer, name, mime) => ({ data: Buffer.from(buffer).toString('base64'), fileName: name, mimeType: mime }) } };
  const fn = new AsyncFunction('$', '$input', 'require', 'Buffer', code09);
  return (await fn.call(self, dollar, { first: () => ({ json: localSap }) }, fakeRequire, Buffer))[0];
}

function makeRawRow(qty, price, amount) {
  const r = new Array(70).fill('');
  r[34] = qty; r[38] = price; r[40] = amount;
  return r;
}

(async () => {
  // 1. Dynamic sapLocalFile in 02_config.js (no hardcoded HK63 08.26.XLSX)
  const r02 = await new AsyncFunction('$input', code02)({
    first: () => ({ json: { Ngay_xu_ly: '2026-09-14', Ma_cong_ty: 'HK13', Nguon_SAP: 'LOCAL_FOLDER', Ty_gia_USD_VND: '25900' } })
  });
  assert.strictEqual(r02[0].json.cfg.sapLocalPath, '/files/finance/Desktop/HK13 09.26.XLSX');
  console.log('PASS 1: Dynamic sapLocalPath per company & month/year');

  // 1b. Windows path inside finance_share is normalized seamlessly
  const r02Win = await new AsyncFunction('$input', code02)({
    first: () => ({ json: { Ngay_xu_ly: '2026-08-14', Ma_cong_ty: 'HK63', Nguon_SAP: 'LOCAL_FOLDER', Ty_gia_USD_VND: '25890', Ghi_de_path_sap: 'D:\\Kiến_tập\\code_space\\finance_share\\Desktop\\HK63 08.26.XLSX' } })
  });
  assert.strictEqual(r02Win[0].json.cfg.sapLocalPath, '/files/finance/Desktop/HK63 08.26.XLSX');
  console.log('PASS 1b: Windows path under finance_share normalized to Docker path');

  // 2. Missing FX rate in LOCAL_FOLDER mode fails closed (no mock/hardcoded fallback)
  await assert.rejects(
    () => new AsyncFunction('$input', code02)({ first: () => ({ json: { Ngay_xu_ly: '2026-08-14', Ma_cong_ty: 'HK63', Nguon_SAP: 'LOCAL_FOLDER', Ty_gia_USD_VND: '' } }) }),
    /FX_RATE_REQUIRED/
  );
  console.log('PASS 2: Missing FX rate throws FX_RATE_REQUIRED (no mock fallback)');

  // 3. Exception Row 81: Missing PIR price (#N/A in col M) -> ERROR + BLOCKED
  const resPir = await runEngine(
    [{ '1': 0, '2': '308853057240', '3': '2300009900', '4': '2026-08-14' }],
    {
      missingPirMaterials: [{ invoice: '2300009900', material: 'MAT_UNKNOWN', billingDate: '2026-08-14' }],
      sapPayload: { value: [{ invoice: '2300009900', billingDate: '2026-08-14', contentF: 'MAT_UNKNOWN', contentG: 'Item', fieldK: 'Pair', quantity: 100, fieldM: 0.1, fieldN: 10, reconcileDifference: 0 }] }
    }
  );
  assert.strictEqual(resPir.json.status, 'BLOCKED');
  assert.ok(!resPir.binary.excel);
  assert.ok(resPir.json.report.steps.some(s => s.step === 'EXCEPTION_PIR_NA' && s.level === 'ERROR'));
  console.log('PASS 3: Missing PIR price (#N/A) logs EXCEPTION_PIR_NA and blocks Excel');

  // 4. Exception Row 84-87: AC != 0 fixable by removing /10 in AA -> WARN EXCEPTION_AA_DIV10 + READY_FOR_REVIEW
  const resAA = await runEngine(
    [{ '1': 0, '2': '308853057240', '3': '2300009900', '4': '2026-08-14' }],
    {
      sapPayload: { value: [{ invoice: '2300009900', billingDate: '2026-08-14', contentF: 'MAT_01', contentG: 'Item', fieldK: 'Pair', quantity: 100, fieldM: 0.02, fieldN: 2, reconcileDifference: 0, rawRow: makeRawRow(100, 0.25, 25), colIndices: { iQty: 34, iPrice: 38, iAmount: 40 } }] }
    }
  );
  assert.strictEqual(resAA.json.status, 'READY_FOR_REVIEW');
  assert.ok(resAA.binary.excel);
  assert.ok(resAA.json.report.steps.some(s => s.step === 'EXCEPTION_AA_DIV10' && s.level === 'WARN'));
  console.log('PASS 4: AC != 0 fixable by removing /10 logs EXCEPTION_AA_DIV10');

  // 5. Exception Row 84-87: AC != 0 unresolvable even without /10 -> ERROR EXCEPTION_AC_UNRESOLVED + BLOCKED
  const resACFail = await runEngine(
    [{ '1': 0, '2': '308853057240', '3': '2300009900', '4': '2026-08-14' }],
    {
      sapPayload: { value: [{ invoice: '2300009900', billingDate: '2026-08-14', contentF: 'MAT_01', contentG: 'Item', fieldK: 'Pair', quantity: 100, fieldM: 0.02, fieldN: 2, reconcileDifference: 0, rawRow: makeRawRow(100, 0.25, 999), colIndices: { iQty: 34, iPrice: 38, iAmount: 40 } }] }
    }
  );
  assert.strictEqual(resACFail.json.status, 'BLOCKED');
  assert.ok(!resACFail.binary.excel);
  assert.ok(resACFail.json.report.steps.some(s => s.step === 'EXCEPTION_AC_UNRESOLVED' && s.level === 'ERROR'));
  console.log('PASS 5: Unresolvable AC != 0 logs EXCEPTION_AC_UNRESOLVED and blocks Excel');

  // 6. Exception Row 88: Qty = 0 dropped & logged as EXCEPTION_QTY_ZERO
  const resZero = await runEngine(
    [{ '1': 0, '2': '308853057240', '3': '2300009900', '4': '2026-08-14' }],
    {
      droppedZeroRows: [{ invoice: '2300009900', material: 'MAT_ZERO', billingDate: '2026-08-14' }],
      sapPayload: { value: [
        { invoice: '2300009900', billingDate: '2026-08-14', contentF: 'MAT_ZERO', contentG: 'Zero', fieldK: 'Pair', quantity: 0, fieldM: 0.02, fieldN: 0, reconcileDifference: 0 },
        { invoice: '2300009900', billingDate: '2026-08-14', contentF: 'MAT_OK', contentG: 'OK', fieldK: 'Pair', quantity: 50, fieldM: 0.02, fieldN: 1, reconcileDifference: 0 }
      ] }
    }
  );
  assert.strictEqual(resZero.json.status, 'READY_FOR_REVIEW');
  assert.strictEqual(resZero.json.summary.droppedZero, 1);
  assert.ok(resZero.json.report.steps.some(s => s.step === 'EXCEPTION_QTY_ZERO' && s.level === 'INFO'));
  console.log('PASS 6: Qty = 0 dropped and logged as EXCEPTION_QTY_ZERO');

  // 7. GAP-P0-04: Bidirectional check (customs invoice missing in SAP -> BLOCKED)
  const resBidirectional = await runEngine(
    [
      { '1': 0, '2': '308853057240', '3': '2300009900', '4': '2026-08-14' },
      { '1': 0, '2': '308853057241', '3': '2300009999', '4': '2026-08-14' }
    ],
    {
      sapPayload: { value: [{ invoice: '2300009900', billingDate: '2026-08-14', contentF: 'MAT_OK', contentG: 'OK', fieldK: 'Pair', quantity: 50, fieldM: 0.02, fieldN: 1, reconcileDifference: 0 }] }
    }
  );
  assert.strictEqual(resBidirectional.json.status, 'BLOCKED');
  assert.ok(resBidirectional.json.errors.some(e => e.includes('2300009999')));
  console.log('PASS 7: Bidirectional check blocks customs invoice missing in SAP (GAP-P0-04)');

  // 8. GAP-P0-05: Non-numeric Qty ("INVALID:abc") -> ERROR + BLOCKED (not silently dropped as 0)
  const resBadQty = await runEngine(
    [{ '1': 0, '2': '308853057240', '3': '2300009900', '4': '2026-08-14' }],
    {
      sapPayload: { value: [{ invoice: '2300009900', billingDate: '2026-08-14', contentF: 'MAT_OK', contentG: 'OK', fieldK: 'Pair', quantity: 'INVALID:abc', fieldM: 0.02, fieldN: 1, reconcileDifference: 0 }] }
    }
  );
  assert.strictEqual(resBadQty.json.status, 'BLOCKED');
  assert.ok(resBadQty.json.errors.some(e => e.includes('số lượng không hợp lệ')));
  console.log('PASS 8: Non-numeric Qty blocked instead of coerced to 0 (GAP-P0-05)');

  // 9. GAP-P0-06: Canceling customs deltas (+0.10 and -0.10) caught by absCheck -> BLOCKED
  const resCancelDelta = await runEngine(
    [
      { '1': 0.10, '2': '308853057240', '3': '2300009900', '4': '2026-08-14' },
      { '1': -0.10, '2': '308853057241', '3': '2300009900', '4': '2026-08-14' }
    ],
    {
      sapPayload: { value: [{ invoice: '2300009900', billingDate: '2026-08-14', contentF: 'MAT_OK', contentG: 'OK', fieldK: 'Pair', quantity: 50, fieldM: 0.02, fieldN: 1, reconcileDifference: 0 }] }
    }
  );
  assert.strictEqual(resCancelDelta.json.status, 'BLOCKED');
  assert.ok(resCancelDelta.json.errors.some(e => e.includes('tổng trị tuyệt đối')));
  console.log('PASS 9: Canceling customs deltas (+0.10/-0.10) caught by absCheck (GAP-P0-06)');

  // 10. SAP_API with manual Form FX override (07c skips 08 - SAP TCURR API)
  const apiParams = {
    cfg: { ...JSON.parse(JSON.stringify(baseCfg)), sapSourceMode: 'SAP_API', sapLocalRate: 25950, fxContractVerified: false },
    customs: { fileName: 'Baocaotokhai.xlsx', base64: Buffer.from(JSON.stringify(makeCustoms([{ '1': 0, '2': '308853057240', '3': '2300009900', '4': '2026-08-14' }]))).toString('base64'), candidates: 1, selection: 'ONLY_MATCH' },
    template: { fileName: 'template.xlsx', base64: Buffer.from(JSON.stringify(makeTemplate())).toString('base64') }
  };
  const apiAll = {
    '06 - Buffer template': [{ json: apiParams }],
    '07 - SAP YINVLST API': [{ json: { value: [{ invoice: '2300009900', billingDate: '2026-08-14', contentF: 'MAT_OK', contentG: 'OK', fieldK: 'Pair', quantity: 50, fieldM: 0.02, fieldN: 1, reconcileDifference: 0 }] } }]
  };
  const apiDollar = (name) => {
    if (name === '08 - SAP TCURR API') throw new Error('08 - SAP TCURR API must not be queried when manual FX is provided');
    return { first: () => apiAll[name][0], all: () => apiAll[name] };
  };
  const apiSelf = { helpers: { prepareBinaryData: async (buffer, name, mime) => ({ data: Buffer.from(buffer).toString('base64'), fileName: name, mimeType: mime }) } };
  const resApiManualFx = (await new AsyncFunction('$', '$input', 'require', 'Buffer', code09).call(apiSelf, apiDollar, { first: () => ({ json: {} }) }, fakeRequire, Buffer))[0];
  assert.strictEqual(resApiManualFx.json.status, 'READY_FOR_REVIEW');
  assert.strictEqual(resApiManualFx.json.report.rate.USD_VND, 25950);
  console.log('PASS 10: SAP_API with manual FX override skips 08 - SAP TCURR API (node 07c)');
})().catch(e => { console.error(e.stack); process.exit(1); });
