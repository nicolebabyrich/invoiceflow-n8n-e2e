/**
 * InvoiceFlow E2E | Gate 1 (LOCAL_FOLDER): Parse SAP Excel file from mounted finance share
 * Supports both raw SAP export (e.g. HK63 08.26.XLSX, header row 1) and processed Ms Tiên invoices (header row 3).
 * Normalizes rows into the exact SAP YINVLST + TCURR contract expected by 09_engine.js.
 */
const item = $input.first();
const cfg = $('02 - Config va bao ve path').first().json.cfg;
if (!item.binary || !item.binary.data) {
  throw new Error('LOCAL_SAP_NOT_FOUND: Không đọc được file SAP tại ' + cfg.sapLocalPath);
}
const XLSX = require('xlsx');
const buf = await this.helpers.getBinaryDataBuffer(0, 'data');
const wb = XLSX.read(buf, { type: 'buffer' });
const sheetName = cfg.sapLocalSheet && wb.SheetNames.includes(cfg.sapLocalSheet) ? cfg.sapLocalSheet : wb.SheetNames[0];
const ws = wb.Sheets[sheetName];
if (!ws) throw new Error('LOCAL_SAP_SHEET_NOT_FOUND: Không tìm thấy sheet ' + sheetName);

const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
const toIsoDate = (v) => {
  if (typeof v === 'number' && v > 20000 && v < 100000) {
    return new Date(Date.UTC(1899, 11, 30) + Math.trunc(v) * 86400000).toISOString().slice(0, 10);
  }
  const s = String(v ?? '').trim();
  let m = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})/);
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  m = s.match(/^(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  return '';
};
const toNum = (v) => {
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  const n = Number(String(v ?? '').replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : 0;
};
const toQtyStrict = (v) => {
  if (typeof v === 'number') return Number.isFinite(v) ? v : 'INVALID_QTY';
  const s = String(v ?? '').replace(/,/g, '').trim();
  if (!s) return 0;
  const n = Number(s);
  return Number.isFinite(n) ? n : `INVALID:${s}`;
};

const isRawSap = rows.length > 0 && String(rows[0][0] || '').trim() === 'Billing type' && String(rows[0][1] || '').trim() === 'Billing Doc. No';
const value = [];

// Dynamic discovery of working file "1. Sale invoices to upload*.xlsx" for the target month/year
const fs = require('fs');
const pirMap = new Map();
let foundWorkFile = '';
const yyyy = cfg.date.slice(0, 4), mm = cfg.date.slice(5, 7);
const dirCandidates = [
  `/files/finance/5. Sale invoices/${yyyy}/${mm}.${yyyy}`,
  '/files/finance/sap'
];
for (const dir of dirCandidates) {
  if (!fs.existsSync(dir)) continue;
  const files = fs.readdirSync(dir).filter(fn => /^1\.\s*Sale invoices to upload.*\.xlsx?$/i.test(fn) && !fn.startsWith('~$'));
  if (files.length) {
    foundWorkFile = `${dir}/${files[0]}`;
    const pWb = XLSX.read(fs.readFileSync(foundWorkFile), { type: 'buffer' });
    const pWs = pWb.Sheets['PIR price'];
    if (pWs) {
      const pRows = XLSX.utils.sheet_to_json(pWs, { header: 1, defval: '' });
      for (let i = 2; i < pRows.length; i++) {
        const mat = String(pRows[i][0] || pRows[i][1] || '').trim();
        const price = Number(pRows[i][13]);
        if (mat && Number.isFinite(price) && price > 0) pirMap.set(mat, price);
      }
    }
    break;
  }
}

const skippedCounts = { canceled: 0, wrongSalesOrg: 0, wrongBillingType: 0, wrongPlant: 0, wrongCustomer: 0 };
const missingPirMaterials = [];
const droppedZeroRows = [];

if (isRawSap) {
  // Resolve SAP columns dynamically by header name (fallback to standard HK63 indices)
  const hdr = rows[0].map(h => String(h || '').trim().toLowerCase());
  const colIdx = (names, fallback) => {
    for (const n of names) {
      const idx = hdr.indexOf(n.toLowerCase());
      if (idx >= 0) return idx;
    }
    return fallback;
  };
  const iBillType = colIdx(['Billing type'], 0);
  const iInv = colIdx(['Billing Doc. No', 'Billing Document'], 1);
  const iDate = colIdx(['Billing Date'], 3);
  const iMat = colIdx(['Material Code', 'Material'], 4);
  const iPlant = colIdx(['Plant'], 15);
  const iCust = colIdx(['Coats Customer Name', 'Customer Name'], 17);
  const iItem = colIdx(['ITEM', 'Material Description', 'Description'], 30);
  const iQty = colIdx(['Qty', 'Bill.qty', 'Billed Quantity'], 34);
  const iUom = colIdx(['Unit of Measure', 'Sales Unit'], 35);
  const iPrice = colIdx(['Price', 'Net Price'], 38);
  const iCanceled = colIdx(['Canceled', 'Cancelled'], 39);
  const iAmount = colIdx(['Amount', 'Net Value'], 40);
  const iSalesOrg = colIdx(['Sales Organization'], 47);

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const billingType = String(r[iBillType] || '').trim().toUpperCase();
    const inv = String(r[iInv] || '').trim();
    const billingDate = toIsoDate(r[iDate]);
    const mat = String(r[iMat] || '').trim();
    const plant = String(r[iPlant] || '').trim().toUpperCase();
    const custName = String(r[iCust] || '').trim().toUpperCase();
    const canceled = String(r[iCanceled] || '').trim();
    const salesOrg = String(r[iSalesOrg] || '').trim().toUpperCase();
    if (!inv || !billingDate) continue;
    if (canceled) { skippedCounts.canceled++; continue; }
    if (salesOrg && cfg.companyCode && salesOrg !== cfg.companyCode) { skippedCounts.wrongSalesOrg++; continue; }
    if (billingType && billingType !== 'ZEX') { skippedCounts.wrongBillingType++; continue; }
    if (plant && plant !== 'VN20') { skippedCounts.wrongPlant++; continue; }
    if (custName && !custName.includes('UNITED WELL')) { skippedCounts.wrongCustomer++; continue; }
    const qty = toQtyStrict(r[iQty]);
    const qtyValid = typeof qty === 'number';
    if (qty === 0) droppedZeroRows.push({ invoice: inv, material: mat, billingDate });
    const hasPir = pirMap.has(mat);
    if (qtyValid && qty !== 0 && pirMap.size > 0 && !hasPir) {
      missingPirMaterials.push({ invoice: inv, material: mat, billingDate });
    }
    const pirPrice = hasPir ? pirMap.get(mat) : toNum(r[iPrice]);
    const amt = (qtyValid && hasPir) ? Number((qty * pirPrice).toFixed(6)) : toNum(r[iAmount]);
    value.push({
      invoice: inv,
      billingDate,
      contentF: mat,
      contentG: String(r[iItem] || '').trim(),
      fieldK: String(r[iUom] || '').trim(),
      quantity: qty,
      fieldM: pirPrice,
      fieldN: amt,
      declarationText: '',
      reconcileDifference: 0,
      rawRow: r.slice(0, 70),
      colIndices: { iQty, iPrice, iAmount }
    });
  }
} else {
  // Processed Ms Tiên invoices: Row 3 (index 2) is header
  for (let i = 3; i < rows.length; i++) {
    const r = rows[i];
    const inv = String(r[19] || r[31] || '').trim();
    const billingDate = toIsoDate(r[1] || r[33]);
    if (!inv || !billingDate) continue;
    const salesOrg = String(r[77] || '').trim().toUpperCase();
    if (salesOrg && cfg.companyCode && salesOrg !== cfg.companyCode) continue;
    value.push({
      invoice: inv,
      billingDate,
      contentF: String(r[5] || '').trim(),
      contentG: String(r[6] || '').trim(),
      fieldK: String(r[10] || '').trim(),
      quantity: toNum(r[11]),
      fieldM: toNum(r[12]),
      fieldN: toNum(r[13]),
      declarationText: String(r[15] || '').trim(),
      reconcileDifference: toNum(r[28])
    });
  }
}

const yesterday = new Date(Date.parse(cfg.date + 'T00:00:00Z') - 86400000).toISOString().slice(0, 10);
return [{
  json: {
    sapSourceMode: 'LOCAL_FOLDER',
    workFilePath: foundWorkFile,
    pirCount: pirMap.size,
    skippedCounts,
    missingPirMaterials,
    droppedZeroRows,
    sapPayload: { value },
    fxPayload: {
      rateVndPerUsd: Number(cfg.sapLocalRate || 0),
      validOn: yesterday
    }
  }
}];
