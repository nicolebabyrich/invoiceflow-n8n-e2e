/**
 * InvoiceFlow E2E | Config node: staff only sets date/profile; IT edits profiles ONCE.
 * Paths are inside an n8n-server mounted root; user cannot access arbitrary files.
 * Configure the API URLs/column mapping to match REAL SAP and customs files before UAT.
 */
const f = $input.first().json;
const profiles = {
  FINANCE_VN: {
    mountRoot: '/files/finance',
    customsGlob: 'Bảng kê tờ khai HQ/TK*/BaoCaoToKhai*.xls*',
    templateFile: '5. Sale invoices/00. UPLOAD INVOICES VIETTEL.xlsx',
    sapLocalFile: 'Desktop/HK63 08.26.XLSX',
    sapLocalSheet: 'Sheet1',
    sapLocalRate: null,
    outputDirectory: 'results',
    sapInvoiceUrl: 'https://SAP_GATEWAY_HOST/finance/invoices',
    sapRateUrl: 'https://SAP_GATEWAY_HOST/finance/usd-vnd-rate',
    // Values below refer to physical source columns of actual BaoCaoToKhai*.xls (header row 9).
    customs: { sheetName: 'Sheet1', headerRow: 9, colInvoice: 'AD', colDate: 'E', colDeclaration: 'B', colCheckDifference: '', colAmount: 'AN' },
    // Contract: normalized JSON from SAP integration must contain these field names.
    sap: { rowsPath: 'value', invoice: 'invoice', date: 'billingDate', F: 'contentF', G: 'contentG', K: 'fieldK', L: 'quantity', M: 'fieldM', N: 'fieldN', P: 'declarationText', reconcileDifference: 'reconcileDifference' },
    fx: { rate: 'rateVndPerUsd', validOn: 'validOn' },
    templateSheet: 'Thông tin hóa đơn',
    toleranceCustoms: 0.03,
    toleranceSAP: 0.000001,
    requireCustomsDifference: false,
    requireSAPDifference: true,
    requireCustomsDate: true,
    requireFXValidOn: true,
    customsMappingVerified: true, // Verified against BaoCaoToKhai 14.8.26.xls
    sapContractVerified: false, // IT sets true after verifying real YINVLST API normalized data
    fxContractVerified: false, // IT sets true after verifying actual TCURR rate factors / validOn
    templateMappingVerified: true, // Verified against 00. UPLOAD INVOICES VIETTEL.xls -> .xlsx
    maxRows: 25000
  }
};
const id = String(f.Profile || 'FINANCE_VN');
if (!Object.hasOwn(profiles, id)) throw new Error('PROFILE_NOT_FOUND: Không có cấu hình profile này.');
const p = JSON.parse(JSON.stringify(profiles[id]));
const d = String(f.Ngay_xu_ly || '').trim();
if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || !Number.isFinite(Date.parse(`${d}T00:00:00Z`)) || new Date(`${d}T00:00:00Z`).toISOString().slice(0, 10) !== d) throw new Error('DATE_INVALID: Dùng ngày YYYY-MM-DD.');
const company = String(f.Ma_cong_ty || '').trim().toUpperCase();
if (!/^[A-Z0-9]{2,12}$/.test(company)) throw new Error('COMPANY_INVALID: Mã công ty không hợp lệ.');
const root = p.mountRoot.replace(/\\/g, '/').replace(/\/$/, '');
function safeUnderRoot(rel, withGlob = false) {
  let raw = String(rel || '').trim().replace(/\\/g, '/');
  if (raw.toLowerCase().startsWith('d:/kiến_tập/code_space/finance_share/')) raw = raw.slice('d:/kiến_tập/code_space/finance_share/'.length);
  else if (raw.startsWith(root + '/')) raw = raw.slice((root + '/').length);
  if (!raw || raw.startsWith('/') || /^[A-Za-z]:/.test(raw) || raw.split('/').some(s => s === '..' || s === '.' || !s) || /[\x00-\x1f]/.test(raw)) throw new Error('PATH_REJECTED: Chỉ cho phép đường dẫn tương đối hoặc đường dẫn nằm trong thư mục finance_share.');
  if (!withGlob && /[*?\[\]{}]/.test(raw)) throw new Error('PATH_REJECTED: Không dùng wildcard ở đường dẫn file cụ thể.');
  if (withGlob && /[{}]/.test(raw)) throw new Error('PATH_REJECTED: Không hỗ trợ pattern dạng {...}.');
  return `${root}/${raw}`;
}
const rawCustoms = String(f.Ghi_de_path_to_khai || '').trim() || p.customsGlob;
const rawTemplate = String(f.Ghi_de_path_template || '').trim() || p.templateFile;
const defaultSapLocal = `Desktop/${company} ${d.slice(5, 7)}.${d.slice(2, 4)}.XLSX`;
const rawSapLocal = String(f.Ghi_de_path_sap || '').trim() || defaultSapLocal;
const rawOutput = String(f.Ghi_de_thu_muc_ket_qua || '').trim() || p.outputDirectory;
const sapSourceMode = String(f.Nguon_SAP || 'LOCAL_FOLDER').trim().toUpperCase() === 'SAP_API' ? 'SAP_API' : 'LOCAL_FOLDER';
const rawFx = String(f.Ty_gia_USD_VND ?? '').trim();
let sapLocalRate = null;
if (rawFx !== '') {
  const parsedFx = Number(rawFx.replace(/,/g, ''));
  if (!Number.isFinite(parsedFx) || parsedFx <= 0) throw new Error('FX_RATE_INVALID: Tỷ giá USD/VND phải là số dương hợp lệ.');
  sapLocalRate = parsedFx;
} else if (f.Nguon_SAP === 'LOCAL_FOLDER') {
  throw new Error('FX_RATE_REQUIRED: Vui lòng nhập tỷ giá USD/VND khi chạy chế độ LOCAL_FOLDER.');
}
const yyyy = d.slice(0, 4), mm = d.slice(5, 7);
const resolveTokens = s => s.replace(/\{YYYY\}/g, yyyy).replace(/\{MM\}/g, mm).replace(/\{YYYYMM\}/g, yyyy + mm);
const from = `${yyyy}-${mm}-01`;
const next = new Date(Date.UTC(Number(yyyy), Number(mm), 1));
const to = new Date(next.getTime() - 86400000).toISOString().slice(0, 10);
const cfg = {
  ...p, profile: id, sapSourceMode, sapLocalRate, requestedBy: String(f.Nguoi_thuc_hien || f['User Email'] || 'n8n Form'),
  date: d, companyCode: company, monthStart: from, monthEnd: to,
  customsPath: safeUnderRoot(resolveTokens(rawCustoms), true),
  templatePath: safeUnderRoot(resolveTokens(rawTemplate)),
  sapLocalPath: safeUnderRoot(resolveTokens(rawSapLocal)),
  outputFolder: safeUnderRoot(resolveTokens(rawOutput)),
};
// This mapping is an explicit IT-owned integration contract, not a guessed SAP endpoint.
if (!/^https:\/\//i.test(cfg.sapInvoiceUrl) || !/^https:\/\//i.test(cfg.sapRateUrl)) throw new Error('SAP_API_CONFIG: URL phải dùng HTTPS.');
return [{ json: { cfg, form: { profile: id, date: d, company } } }];
