/**
 * InvoiceFlow E2E | n8n Code / Run Once for All Items
 * Sources: actual Baocaotokhai XLSX + normalized SAP YINVLST JSON + normalized TCURR JSON + Viettel template XLSX.
 * Creates a review-only Viettel XLSX (NO UPLOAD) and auditable HTML/JSON/CSV logs.
 * IMPORTANT: SAP/Customs field contracts must be verified by IT against real files/APIs.
 * External module required in n8n JS Task Runner: exceljs.
 */
const startedAt = new Date().toISOString();
const params = $('06 - Buffer template').first().json;
const cfg = params.cfg;
const steps = [], defects = [], warnings = [];
const stats={ customsRows:0, sapRows:0, sapRowsForDate:0, matched:0, droppedZero:0, details:0, noteRows:0, invoiceGroups:0, customsDifference:null, sapDifference:null };
const log = (level,step,message,details='')=>{steps.push({time:new Date().toISOString(),level,step,message:String(message),details:String(details)});if(level==='ERROR')defects.push(String(message));if(level==='WARN')warnings.push(String(message));};
const raw=v=>v&&typeof v==='object'&&Object.prototype.hasOwnProperty.call(v,'result')?v.result:v&&typeof v==='object'&&Object.prototype.hasOwnProperty.call(v,'text')?v.text:v;
const txt=v=>String(raw(v)??'').trim();
const num=v=>{v=raw(v);if(typeof v==='number')return Number.isFinite(v)?v:NaN;let s=txt(v).replace(/[\s\u00a0]/g,'');if(!s)return NaN;if(s.includes(',')&&s.includes('.'))s=s.lastIndexOf(',')>s.lastIndexOf('.')?s.replace(/\./g,'').replace(',','.'):s.replace(/,/g,'');else if(s.includes(','))s=s.replace(',','.');return Number(s)};
const date=v=>{v=raw(v);if(v instanceof Date)return `${v.getUTCFullYear()}-${String(v.getUTCMonth()+1).padStart(2,'0')}-${String(v.getUTCDate()).padStart(2,'0')}`;
if(typeof v==='number'&&v>20000&&v<100000){const d=new Date(Date.UTC(1899,11,30)+Math.trunc(v)*86400000);return d.toISOString().slice(0,10)};
let s=txt(v);let m=s.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);if(m)return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
m=s.match(/^(\d{1,2})[-./](\d{1,2})[-./](\d{4})/);if(m)return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;return s.slice(0,10)};
const invoice=v=>txt(v).replace(/\s/g,'').replace(/\.0$/,'').toUpperCase();
const expand=v=>{const x=invoice(v).split(/[\/;,]/).filter(Boolean);if(!x.length)return[];const first=x[0];return [...new Set([first,...x.slice(1).map(p=>/^\d+$/.test(first)&&/^\d+$/.test(p)&&p.length<first.length?first.slice(0,first.length-p.length)+p:p)])]};
const escape=v=>txt(v).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const csv=v=>'"'+txt(v).replace(/"/g,'""')+'"';
const pick=(o,k)=>String(k||'').split('.').filter(Boolean).reduce((v,p)=>v==null?undefined:v[p],o);
const colIndex=c=>typeof c==='number'?c:String(c||'').toUpperCase().split('').reduce((n,ch)=>n*26+ch.charCodeAt(0)-64,0);
const sha256=v=>{try{return require('crypto').createHash('sha256').update(v).digest('hex')}catch{return 'NO_CRYPTO_AVAILABLE'}};
const safe=`${cfg.companyCode}_${cfg.date.replace(/-/g,'')}_${Date.now()}`.replace(/[^A-Za-z0-9_-]/g,'');
const filenames={excel:'00. UPLOAD INVOICES VIETTEL.xlsx',audit_xlsx:`${safe}_AUDIT.xlsx`,html:`${safe}_AUDIT.html`,csv:`${safe}_AUDIT.csv`,json:`${safe}_AUDIT.json`};
let ExcelJS=null,xlsx=null,fxRate=null,fxDate=null,customsHash='',processedDecls=[],processedCustomsDirs=[];
try {
  const isApiMode = cfg.sapSourceMode !== 'LOCAL_FOLDER';
  const hasManualFx = Number(cfg.sapLocalRate) > 0;
  if(!cfg.customsMappingVerified||!cfg.templateMappingVerified) throw new Error('CONFIG_NOT_APPROVED: IT phải nghiệm thu mapping tờ khai và mẫu Viettel, sau đó bật *Verified trong node 02.');
  if(isApiMode && (!cfg.sapContractVerified||(!cfg.fxContractVerified&&!hasManualFx))) throw new Error('CONFIG_NOT_APPROVED: Chế độ SAP API yêu cầu IT nghiệm thu hợp đồng SAP YINVLST/TCURR và bật sapContractVerified/fxContractVerified trong node 02.');
  try{ExcelJS=require('exceljs')}catch(e){throw new Error('MISSING_EXCELJS: IT chưa cài hoặc chưa cho phép exceljs trên JS task runner. '+e.message)}
  const customsList=(Array.isArray(params.customsFiles)&&params.customsFiles.length)?params.customsFiles:[params.customs];
  const tbuf=Buffer.from(params.template.base64,'base64');
  customsHash=sha256(Buffer.from(customsList[0].base64,'base64'));
  for(const cf of customsList){
    const cHash=sha256(Buffer.from(cf.base64,'base64'));
    log('INFO','SOURCE',`Tờ khai: ${cf.fileName} (${cf.candidates} file khớp; cách chọn ${cf.selection})`,cHash);
    if(cf.selection==='FILENAME_LEXICAL')log('WARN','SOURCE','Nhiều file trùng glob, nhưng không có timestamp trong filename: đã chọn theo tên file, KHÔNG theo Date modified. IT nên chuẩn hóa quy tắc tên file.');
  }
  log('INFO','SOURCE',`Template: ${params.template.fileName}`,sha256(tbuf));
  const tw=new ExcelJS.Workbook();await tw.xlsx.load(tbuf);
  const vs=cfg.templateSheet?tw.getWorksheet(cfg.templateSheet):tw.worksheets[0];
  if(!vs)throw new Error('TEMPLATE_SHEET_NOT_FOUND: Sheet template Viettel không tồn tại.');
  const cols=cfg.customs;
  const docs=new Map(),rawCustomsRows=[],seenCustomsDecls=new Set();let check=0,absCheck=0,checkN=0;
  const processCustomsRow=(getVal,ri,rawArr)=>{
    if(ri<=Number(cols.headerRow||1))return;
    const inv=invoice(getVal(cols.colInvoice));if(!inv)return;
    if(cfg.requireCustomsDate&&date(getVal(cols.colDate))!==cfg.date)return;
    const decl=txt(getVal(cols.colDeclaration));
    const dedupKey=`${decl}|${inv}`;
    if(decl&&seenCustomsDecls.has(dedupKey))return;
    if(decl){seenCustomsDecls.add(dedupKey);if(!processedDecls.includes(decl))processedDecls.push(decl);}
    stats.customsRows++;
    if(rawArr)rawCustomsRows.push({decl,inv,arr:rawArr});
    if(!decl)log('ERROR','CUSTOMS',`Dòng ${ri} thiếu số tờ khai`,inv);
    if(cfg.requireCustomsDifference){const delta=num(getVal(cols.colCheckDifference));if(!Number.isFinite(delta))log('ERROR','CUSTOMS',`Dòng ${ri} thiếu giá trị kiểm tra chênh lệch (cột ${cols.colCheckDifference}).`,inv);else {check+=delta;absCheck+=Math.abs(delta);checkN++;}}
    const amount=cols.colAmount?num(getVal(cols.colAmount)):NaN;
    for(const key of expand(inv)){
      if(!docs.has(key))docs.set(key,{groupKey:inv,declarations:new Set(),amounts:[]});
      const bucket=docs.get(key);if(decl)bucket.declarations.add(decl);if(Number.isFinite(amount)&&!bucket._seenDeclAmt?.has(decl)){
        if(!bucket._seenDeclAmt)bucket._seenDeclAmt=new Set();
        if(decl)bucket._seenDeclAmt.add(decl);
        bucket.amounts.push(amount);
      }
    }
  };
  for(const cf of customsList){
    if(cf.directory&&!processedCustomsDirs.includes(cf.directory))processedCustomsDirs.push(cf.directory);
    const cbuf=Buffer.from(cf.base64,'base64');
    if(cbuf.length>=4&&cbuf[0]===0xD0&&cbuf[1]===0xCF&&cbuf[2]===0x11&&cbuf[3]===0xE0){
      let XLSX;try{XLSX=require('xlsx')}catch(e){throw new Error('MISSING_XLSX: Cần cài xlsx (SheetJS) để đọc file .xls. '+e.message)}
      const wb=XLSX.read(cbuf,{type:'buffer'});
      const sname=cfg.customs.sheetName&&wb.SheetNames.includes(cfg.customs.sheetName)?cfg.customs.sheetName:wb.SheetNames[0];
      const ws=wb.Sheets[sname];if(!ws)throw new Error('CUSTOMS_SHEET_NOT_FOUND: Sheet tờ khai không tồn tại.');
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
      rows.forEach((r,idx)=>processCustomsRow(c=>r[colIndex(c)-1],idx+1,r));
    } else {
      const cw=new ExcelJS.Workbook();await cw.xlsx.load(cbuf);
      const cs=cfg.customs.sheetName?cw.getWorksheet(cfg.customs.sheetName):cw.worksheets[0];
      if(!cs)throw new Error('CUSTOMS_SHEET_NOT_FOUND: Sheet tờ khai không tồn tại.');
      cs.eachRow({includeEmpty:false},(row,ri)=>processCustomsRow(c=>raw(row.getCell(colIndex(c)).value),ri,Array.from({length:106},(_,k)=>raw(row.getCell(k+1).value))));
    }
  }
  stats.customsDifference=Number(check.toFixed(8));
  if(stats.customsRows===0)log('ERROR','CUSTOMS',`Không có tờ khai hợp lệ ngày ${cfg.date}. Kiểm tra path, sheet và cột mapping.`);
  if(cfg.requireCustomsDifference&&checkN!==stats.customsRows)log('ERROR','CUSTOMS','Có tờ khai chưa được kiểm tra chênh lệch đầy đủ.');
  if(Math.abs(check)>cfg.toleranceCustoms||absCheck>cfg.toleranceCustoms)log('ERROR','CUSTOMS',`Chênh lệch tờ khai vượt ±${cfg.toleranceCustoms} (tổng=${check.toFixed(5)}, tổng trị tuyệt đối=${absCheck.toFixed(5)}).`);
  log('INFO','CUSTOMS',`Đọc ${stats.customsRows} dòng, tổng delta=${stats.customsDifference}.`,`${docs.size} mã hóa đơn sau khi tách hóa đơn kép`);
  const localSapData = !isApiMode ? $input.first().json : null;
  const sapItems = isApiMode ? $('07 - SAP YINVLST API').all().map(i=>i.json) : [localSapData.sapPayload];
  for(const v of sapItems)if(v?.error||v?.statusCode>=400)log('ERROR','SAP_YINVLST',`API SAP trả lỗi: ${txt(v.error?.message||v.message||v.statusCode)}`);
  const sapRoot=sapItems.length===1?sapItems[0]:sapItems;
  const candidate=Array.isArray(sapRoot)?sapRoot:pick(sapRoot,cfg.sap.rowsPath);
  const sapRows=Array.isArray(candidate)?candidate:[];
  if(!sapRows.length)log('ERROR','SAP_YINVLST',isApiMode?'API YINVLST không trả danh sách invoice theo contract đã cấu hình.':'File SAP local không có dữ liệu invoice.');
  const yesterday=new Date(Date.parse(cfg.date+'T00:00:00Z')-86400000).toISOString().slice(0,10);
  const rateItems = !isApiMode
    ? [localSapData.fxPayload]
    : hasManualFx
      ? [{ rateVndPerUsd: Number(cfg.sapLocalRate), validOn: yesterday }]
      : $('08 - SAP TCURR API').all().map(i=>i.json);
  const fxObj=rateItems.length===1?rateItems[0]:rateItems[0]||{};
  if(fxObj?.error||fxObj?.statusCode>=400)log('ERROR','SAP_TCURR',`API tỷ giá trả lỗi: ${txt(fxObj.error?.message||fxObj.message||fxObj.statusCode)}`);
  fxRate=num(pick(fxObj,cfg.fx.rate));fxDate=date(pick(fxObj,cfg.fx.validOn));
  if(!Number.isFinite(fxRate)||fxRate<=0)log('ERROR','SAP_TCURR','Thiếu tỷ giá USD/VND dương/hợp lệ trong response SAP.');
  if(cfg.requireFXValidOn&&fxDate!==yesterday)log('ERROR','SAP_TCURR',`TCURR ngày hiệu lực cần ${yesterday}, nhận ${fxDate||'trống'}.`);
  log('INFO','SAP_TCURR',`Tỷ giá USD/VND=${fxRate}, ngày hiệu lực=${fxDate}.`);
  const map=cfg.sap,group=new Map();let sapDelta=0,deltaN=0;
  for(const d of docs.values())if(!group.has(d.groupKey))group.set(d.groupKey,[]);
  stats.sapRows=sapRows.length;
  for(let j=0;j<sapRows.length;j++){
    const x=sapRows[j],bill=date(pick(x,map.date));
    const inv=invoice(pick(x,map.invoice));
    const inCustoms=inv&&expand(inv).some(id=>docs.has(id));
    if(bill!==cfg.date&&!inCustoms)continue;
    stats.sapRowsForDate++;
    const qty=num(pick(x,map.L));
    if(!inv){log('ERROR','SAP_YINVLST',`SAP record ${j+1} thiếu invoice.`);continue}
    if(!Number.isFinite(qty)){log('ERROR','SAP_YINVLST',`SAP invoice ${inv} số lượng không hợp lệ.`);continue}
    if(qty===0){stats.droppedZero++;continue}
    const delta=num(pick(x,map.reconcileDifference));
    if(cfg.requireSAPDifference){if(!Number.isFinite(delta))log('ERROR','RECONCILE',`SAP invoice ${inv} thiếu reconciliationDifference (kiểm soát tương đương AC1).`);else{sapDelta+=delta;deltaN++;}}
    const keys=expand(inv).filter(id=>docs.has(id));
    if(!keys.length){log('ERROR','RECONCILE',`SAP hóa đơn ${inv} không có trong tờ khai ngày ${cfg.date}.`);continue}
    stats.matched++;
    const groupKey=docs.get(keys[0]).groupKey||inv;
    const decl=[...new Set(keys.flatMap(k=>[...docs.get(k).declarations]))];
    const note=txt(pick(x,map.P))||`Gia công xuất khẩu theo tờ khai số ${decl.join(' và ')}`;
    if(!decl.length)log('ERROR','RECONCILE',`SAP hóa đơn ${inv} không có số tờ khai để viết ghi chú.`);
    const m=pick(x,map.M);
    if(m==null||txt(m)===''||/^#(N\/A|REF!|VALUE!|DIV\/0!)/i.test(txt(m)))log('ERROR','RECONCILE',`SAP hóa đơn ${inv} thiếu/lỗi trường M.`);
    const row={invoice:groupKey,rawInv:inv,rawRow:x.rawRow||null,colIndices:x.colIndices||null,F:pick(x,map.F),G:pick(x,map.G),K:pick(x,map.K),L:qty,M:m,N:pick(x,map.N),note,decl};
    if(!group.has(groupKey))group.set(groupKey,[]);
    group.get(groupKey).push(row);
  }
  stats.sapDifference=Number(sapDelta.toFixed(8));
  if(cfg.requireSAPDifference&&deltaN!==stats.sapRowsForDate-stats.droppedZero)log('ERROR','RECONCILE','Chưa kiểm đủ chênh lệch tài chính từ SAP.');
  if(cfg.requireSAPDifference&&Math.abs(sapDelta)>cfg.toleranceSAP)log('ERROR','RECONCILE',`Tổng chênh lệch SAP ${sapDelta.toFixed(5)} khác 0. Phải xử lý trước khi xuất.`);
  if(!stats.sapRowsForDate)log('ERROR','SAP_YINVLST',`SAP không có hóa đơn ngày ${cfg.date}.`);
  for(const [k,v] of group){
    if(!v.length){
      log('ERROR','RECONCILE',`Tờ khai ngày ${cfg.date} có hóa đơn ${k} nhưng không có dữ liệu khớp trong SAP.`);
      group.delete(k);
    }else{
      const docBucket=docs.get(expand(k)[0]);
      const tkSum=(docBucket?.amounts||[]).reduce((a,b)=>a+b,0);
      const hasRaw=v.some(item=>item.rawRow);
      if(hasRaw&&docBucket?.amounts?.length){
        const sapSum=v.reduce((a,item)=>{
          const ci=item.colIndices||{iQty:34,iPrice:38,iAmount:40};
          const bs=num(item.rawRow?.[ci.iAmount]);
          return a+(Number.isFinite(bs)?bs:0);
        },0);
        const grpDelta=Number((tkSum-sapSum).toFixed(4));
        if(Math.abs(grpDelta)>cfg.toleranceCustoms){
          log('ERROR','RECONCILE',`Nhóm hóa đơn ${k}: tổng tiền tờ khai (${tkSum.toFixed(2)}) lệch tổng SAP (${sapSum.toFixed(2)}) là ${grpDelta.toFixed(2)} vượt ngưỡng ±${cfg.toleranceCustoms}.`);
        }
      }
    }
  }
  if(!group.size)log('ERROR','RECONCILE','Không có nhóm invoice hợp lệ sau đối chiếu.');
  stats.invoiceGroups=group.size;
  log('INFO','RECONCILE',`Tổng dữ liệu SAP hợp lệ: ${stats.sapRows} dòng; thuộc kỳ đang xử lý: ${stats.sapRowsForDate} dòng (khớp tờ khai: ${stats.matched} dòng, bỏ ${stats.droppedZero} dòng số lượng = 0, chênh lệch: ${stats.sapDifference}).`);
  if(localSapData){
    if(localSapData.skippedCounts){
      const sc=localSapData.skippedCounts;
      const parts=[];
      if(sc.canceled)parts.push(`${sc.canceled} dòng đã hủy`);
      if(sc.wrongBillingType)parts.push(`${sc.wrongBillingType} dòng khác loại hóa đơn (không phải ZEX)`);
      if(sc.wrongPlant)parts.push(`${sc.wrongPlant} dòng khác nhà máy (không phải VN20)`);
      if(sc.wrongCustomer)parts.push(`${sc.wrongCustomer} dòng khách hàng khác`);
      if(sc.wrongSalesOrg)parts.push(`${sc.wrongSalesOrg} dòng khác mã công ty`);
      if(parts.length)log('INFO','SAP_FILTER',`Đã lọc bỏ các dòng SAP không thuộc phạm vi xuất hóa đơn: ${parts.join(', ')}.`);
    }
    const zeroForDate=(localSapData.droppedZeroRows||[]).filter(z=>z.billingDate===cfg.date||docs.has(z.invoice));
    if(zeroForDate.length){
      log('INFO','EXCEPTION_QTY_ZERO',`Đã tự động loại bỏ ${zeroForDate.length} dòng có Số lượng (Quantity) = 0 theo quy trình.`,'Ví dụ: '+zeroForDate.slice(0,8).map(z=>`${z.invoice} (${z.material})`).join(', '));
    }
    const pirMissing=(localSapData.missingPirMaterials||[]).filter(p=>p.billingDate===cfg.date||docs.has(p.invoice));
    if(pirMissing.length){
      const uniqMat=[...new Set(pirMissing.map(p=>`${p.invoice} (${p.material})`))];
      log('ERROR','EXCEPTION_PIR_NA',`Thiếu đơn giá trong sheet PIR price (${uniqMat.length} mã bị lỗi #N/A ở cột M) — Đã dừng để bổ sung giá PIR trước khi xuất.`,uniqMat.slice(0,15).join(', '));
    }
  }
  for(const [,items] of group){
    for(const item of items){
      if(!item.rawRow)continue;
      const ci=item.colIndices||{iQty:34,iPrice:38,iAmount:40};
      const bq=num(item.rawRow[ci.iPrice]),q0=num(item.rawRow[ci.iQty]),bs=num(item.rawRow[ci.iAmount]);
      if(Number.isFinite(bq)&&Number.isFinite(q0)&&Number.isFinite(bs)){
        const diffDiv10=Math.abs(Math.round((bq/10)*q0*100)/100-bs);
        const diffNoDiv10=Math.abs(Math.round(bq*q0*100)/100-bs);
        if(diffDiv10>0.03&&diffNoDiv10<=0.03){
          item.fixAA=true;
          log('WARN','EXCEPTION_AA_DIV10',`Hóa đơn ${item.rawInv} (${item.F}): Thành tiền bị lệch khi chia 10 (lệch ${diffDiv10.toFixed(2)}), hệ thống đã tự động bỏ "/10" ở cột đơn giá (cột AA).`,`Đơn giá SAP=${bq}, Số lượng=${q0}, Tổng tiền SAP=${bs}`);
        }else if(diffDiv10>0.03&&diffNoDiv10>0.03){
          log('ERROR','EXCEPTION_AC_UNRESOLVED',`Hóa đơn ${item.rawInv} (${item.F}): Cột kiểm tra chênh lệch (AC) vẫn lệch sau khi thử bỏ "/10" — Đã dừng để kế toán kiểm tra lại.`, `Đơn giá SAP=${bq}, Số lượng=${q0}, Tổng tiền SAP=${bs}`);
        }
      }
    }
  }
  if(stats.sapRows>cfg.maxRows)log('ERROR','LIMIT',`SAP có ${stats.sapRows} dòng, vượt giới hạn ${cfg.maxRows}.`);
  const sample=vs.getRow(11);
  let populated=0;
  for(let c=4;c<=25;c++)if(sample.getCell(c).value!==null&&sample.getCell(c).value!==undefined)populated++;
  if(populated<4)log('ERROR','TEMPLATE','Dòng mẫu 11 (D:Y) không có đủ ô được định nghĩa.');
  const last=vs.lastRow?.number||11;
  for(let r=12;r<=last;r++)if(vs.getRow(r).hasValues){log('ERROR','TEMPLATE',`Template có dữ liệu ở dòng ${r}. Phải dùng bản trống sau dòng 11.`);break}
  if(defects.length)throw new Error('RECONCILIATION_BLOCKED');
  // Write only after all strict checks pass. Never upload or issue invoice.
  vs.getCell('V11').value=String(fxRate);
  const copyVal=v=>{if(v==null||typeof v!=='object')return v;if(v instanceof Date)return new Date(v.getTime());return JSON.parse(JSON.stringify(v))};
  const shift=(formula,off)=>String(formula).replace(/(?<![\w])([\$]?)([A-Z]{1,3})([\$]?)(\d+)(?![\w])/g,(m,a,c,b,r)=>a+c+b+(b?r:String(Number(r)+off)));
  const tplSnap=Array.from({length:25},(_,idx)=>{const cell=sample.getCell(idx+1);return {style:copyVal(cell.style||{}),value:copyVal(cell.value)};});
  const tplHeight=sample.height;
  const copySnapToRow=(dst,from,to,off)=>{
    for(let c=from;c<=to;c++){const s=tplSnap[c-1],t=dst.getCell(c);if(s.style&&Object.keys(s.style).length)t.style=copyVal(s.style);t.value=copyVal(s.value);
      if(t.value&&typeof t.value==='object'&&t.value.formula){t.value={...t.value,formula:shift(t.value.formula,off),result:undefined};}
    }
    if(tplHeight)dst.height=tplHeight;
  };
  let rowNumber=11,hasFormula=false,groupNum=Number(tplSnap[1].value)||1;
  for(const [inv,items] of group){
    for(const item of items){
      const r=vs.getRow(rowNumber++);
      if(r.number>11)copySnapToRow(r,2,25,r.number-11);
      if(Array.from({length:22},(_,k)=>r.getCell(k+4).value).some(v=>v&&typeof v==='object'&&v.formula))hasFormula=true;
      r.getCell('B').value=groupNum;r.getCell('C').value=null;r.getCell('AA').value=item.F;r.getCell('AB').value=item.G;
      r.getCell('AF').value=item.K;r.getCell('AG').value=item.L;r.getCell('AH').value=item.M;r.getCell('AI').value=item.N;
      stats.details++;
    }
    const notes=[...new Set(items.map(r=>r.note).filter(Boolean))];
    if(items.length){const nr=vs.getRow(rowNumber++);copySnapToRow(nr,2,22,nr.number-11);
      nr.getCell('B').value=groupNum;nr.getCell('C').value=null;nr.getCell('Y').value=2;nr.getCell('AA').value=null;nr.getCell('AB').value=notes.join(' và ');
      nr.getCell('AF').value=null;nr.getCell('AG').value=null;nr.getCell('AH').value=null;nr.getCell('AI').value=null;
      stats.noteRows++;groupNum++;
    }
  }
  tw.calcProperties.fullCalcOnLoad=true;
  if(hasFormula)log('WARN','TEMPLATE','D:Y có công thức sao chép. ExcelJS không tính công thức; phải mở và cho Excel recalculate trước khi upload Viettel.');
  xlsx=await tw.xlsx.writeBuffer();
  log('INFO','OUTPUT',`Đã tạo ${stats.details} dòng chi tiết và ${stats.noteRows} dòng ghi chú, chờ kiểm tra trước upload Viettel.`);
  if(localSapData&&localSapData.workFilePath){
    const fs=require('fs');
    if(fs.existsSync(localSapData.workFilePath)){
      const workWb=new ExcelJS.Workbook();
      await workWb.xlsx.readFile(localSapData.workFilePath);
      const tkWs=workWb.getWorksheet('Tờ khai');
      if(tkWs){
        const existingTkDecls=new Set();
        let lastTkRow=10;
        tkWs.eachRow((r,rn)=>{
          if(rn<11)return;
          const declVal=txt(r.getCell('I').value);
          if(declVal)existingTkDecls.add(declVal);
          if(txt(r.getCell('C').value)||declVal)lastTkRow=Math.max(lastTkRow,rn);
        });
        // Append BaoCaoToKhai rows for cfg.date if not yet present in sheet Tờ khai
        let tkAppended=0;
        const tplTkRow=tkWs.getRow(lastTkRow);
        const resolveTkFormula=(cellVal,dstRow)=>{
          if(cellVal&&typeof cellVal==='object'){
            if(cellVal.formula)return {formula:shift(cellVal.formula,dstRow-tplTkRow.number)};
            if(cellVal.sharedFormula){
              const mCell=tkWs.getCell(cellVal.sharedFormula).value;
              const mRow=Number(String(cellVal.sharedFormula).replace(/^[A-Z]+/,''))||11;
              if(mCell&&mCell.formula)return {formula:shift(mCell.formula,dstRow-mRow)};
            }
          }
          return copyVal(cellVal);
        };
        const declToCustoms=new Map(rawCustomsRows.map(cr=>[cr.decl,cr]));
        for(const cr of rawCustomsRows){
          if(!cr.decl||existingTkDecls.has(cr.decl))continue;
          const nr=++lastTkRow;
          const dstTk=tkWs.getRow(nr);
          for(let c=1;c<=6;c++){
            const sv=tplTkRow.getCell(c).value;
            const ss=tplTkRow.getCell(c).style;
            if(ss&&Object.keys(ss).length)dstTk.getCell(c).style=copyVal(ss);
            dstTk.getCell(c).value=resolveTkFormula(sv,nr);
          }
          for(let c=0;c<Math.min(106,cr.arr.length);c++){
            const v=cr.arr[c];
            dstTk.getCell(8+c).value=(v===undefined||v==='')?null:copyVal(v);
          }
          existingTkDecls.add(cr.decl);
          tkAppended++;
        }
        if(tkAppended>0)log('INFO','TO_KHAI_APPEND',`Đã thêm ${tkAppended} dòng tờ khai mới vào sheet Tờ khai và điền công thức cột A–F.`);
        const seenTkInv=new Set();
        const clearedDupTk=[];
        tkWs.eachRow((r,rn)=>{
          if(rn<11)return;
          const declVal=txt(r.getCell('I').value);
          const fromCustoms=declToCustoms.get(declVal);
          const dVal=fromCustoms?cfg.date:date(r.getCell('D').value);
          const invVal=fromCustoms?fromCustoms.inv:invoice(r.getCell('C').value);
          if(dVal===cfg.date&&invVal){
            if(seenTkInv.has(invVal)){
              if(r.getCell('B').value!==null){r.getCell('B').value=null;clearedDupTk.push(`Dòng ${rn} (${invVal})`);}
            }else seenTkInv.add(invVal);
          }
        });
        if(clearedDupTk.length)log('INFO','EXCEPTION_TK_DUP_INVOICE',`Đã xóa công thức cột Số tiền (cột B) tại ${clearedDupTk.length} dòng trùng số hóa đơn trên sheet Tờ khai để tránh cộng trùng.`,clearedDupTk.join(', '));
      }
      const msWs=workWb.getWorksheet('Ms Tiên invoices');
      if(msWs){
        let lastMsRow=3;
        const existingMsInvs=new Set();
        msWs.eachRow((r,rn)=>{
          if(rn<4)return;
          const rInv=txt(r.getCell('R').value)||txt(r.getCell('T').value);
          if(rInv)existingMsInvs.add(rInv);
          if(txt(r.getCell('B').value)||rInv||txt(r.getCell('AE').value))lastMsRow=Math.max(lastMsRow,rn);
        });
        const tplMsRow=msWs.getRow(lastMsRow);
        const tplD=txt(tplMsRow.getCell('D').value)||cfg.companyCode;
        const tplE=txt(tplMsRow.getCell('E').value)||'';
        let msAdded=0,msFixedAA=0,msMultiDecl=0;
        const multiDeclLogged=new Set();
        for(const [,items] of group){
          for(const item of items){
            if(!item.rawRow||existingMsInvs.has(item.rawInv))continue;
            const rn=++lastMsRow;
            const dst=msWs.getRow(rn);
            for(let c=2;c<=30;c++){
              const sc=tplMsRow.getCell(c).style;
              if(sc&&Object.keys(sc).length)dst.getCell(c).style=copyVal(sc);
            }
            for(let i=0;i<item.rawRow.length;i++)dst.getCell(31+i).value=item.rawRow[i]===''?null:item.rawRow[i];
            dst.getCell('B').value={formula:`VLOOKUP(R${rn},'Tờ khai'!C:D,2,0)`};
            dst.getCell('C').value={formula:`IF(O${rn}=O${rn-1},C${rn-1},C${rn-1}+1)`};
            dst.getCell('D').value=tplD;
            dst.getCell('E').value=tplE;
            dst.getCell('F').value={formula:`W${rn}`};
            dst.getCell('G').value={formula:`X${rn}`};
            dst.getCell('K').value={formula:`Y${rn}`};
            dst.getCell('L').value={formula:`Z${rn}`};
            dst.getCell('M').value={formula:`_xlfn.XLOOKUP(F${rn},'PIR price'!B:B,'PIR price'!N:N)`};
            dst.getCell('N').value={formula:`ROUND(M${rn}*L${rn},4)`};
            dst.getCell('O').value={formula:`_xlfn.XLOOKUP(R${rn},'Tờ khai'!C:C,'Tờ khai'!I:I)`};
            if(item.decl.length>1){
              msMultiDecl++;
              if(!multiDeclLogged.has(item.rawInv)){
                multiDeclLogged.add(item.rawInv);
                log('INFO','EXCEPTION_MULTI_DECL_P',`Hóa đơn ${item.rawInv} đi kèm ${item.decl.length} tờ khai → đã tự động nối thêm số tờ khai "${item.decl.slice(1).join(' và ')}" vào cột Diễn giải (cột P).`);
              }
              dst.getCell('P').value={formula:`"Gia công xuất khẩu theo tờ khai số"&" "&O${rn}&" và ${item.decl.slice(1).join(' và ')}"`};
            }else{
              dst.getCell('P').value={formula:`"Gia công xuất khẩu theo tờ khai số"&" "&O${rn}`};
            }
            dst.getCell('R').value=item.invoice;
            dst.getCell('T').value={formula:`AF${rn}`};
            dst.getCell('U').value={formula:`AH${rn}`};
            dst.getCell('V').value={formula:`AP${rn}`};
            dst.getCell('W').value={formula:`AI${rn}`};
            dst.getCell('X').value={formula:`BI${rn}`};
            dst.getCell('Y').value={formula:`BN${rn}`};
            dst.getCell('Z').value={formula:`BM${rn}`};
            if(item.fixAA){
              dst.getCell('AA').value={formula:`BQ${rn}`};
              msFixedAA++;
            }else{
              dst.getCell('AA').value={formula:`BQ${rn}/10`};
            }
            dst.getCell('AB').value={formula:`ROUND(AA${rn}*Z${rn},2)`};
            dst.getCell('AC').value={formula:`AB${rn}-BS${rn}`};
            msAdded++;
          }
        }
        workWb.calcProperties.fullCalcOnLoad=true;
        const outWorkPath=`${cfg.outputFolder}/${localSapData.workFilePath.split(/[\\/]/).pop()}`;
        fs.writeFileSync(outWorkPath,Buffer.from(await workWb.xlsx.writeBuffer()));
        log('INFO','MS_TIEN_INVOICE',`Đã cập nhật ${msAdded} dòng vào sheet Ms Tiên invoices và điền đầy đủ công thức (${msMultiDecl} dòng nối nhiều tờ khai ở cột P, ${msFixedAA} dòng tự bỏ /10 ở cột AA).`,`File đã lưu: ${outWorkPath}`);
      }
    }
  }
} catch(e){
  if(!defects.length||e.message!=='RECONCILIATION_BLOCKED')log('ERROR','ENGINE',e.message||String(e));
}
const status=defects.length?'BLOCKED':'READY_FOR_REVIEW';
const declSummary=processedDecls.length?processedDecls.join(', '):'không có';
if(status==='READY_FOR_REVIEW'){
  const fs=require('fs');
  const monthFolder=`${cfg.date.slice(5,7)}.${cfg.date.slice(0,4)}`;
  const movedDirs=[];
  for(const dir of processedCustomsDirs){
    const normDir=String(dir).replace(/\\/g,'/').replace(/\/$/,'');
    const folderName=normDir.split('/').pop();
    const parentDir=normDir.slice(0,normDir.lastIndexOf('/'));
    if(!folderName||!/^TK/i.test(folderName)||normDir.includes(`/${monthFolder}/`))continue;
    if(fs.existsSync(normDir)){
      const archiveRoot=`${parentDir}/${monthFolder}`;
      fs.mkdirSync(archiveRoot,{recursive:true});
      let destDir=`${archiveRoot}/${folderName}`;
      if(fs.existsSync(destDir))destDir=`${archiveRoot}/${folderName}_${Date.now()}`;
      try{
        fs.renameSync(normDir,destDir);
      }catch{
        fs.cpSync(normDir,destDir,{recursive:true});
        fs.rmSync(normDir,{recursive:true,force:true});
      }
      movedDirs.push(`${normDir} -> ${destDir}`);
    }
  }
  if(movedDirs.length){
    log('INFO','TO_KHAI_ARCHIVE',`Đã xử lý ${processedDecls.length} số tờ khai ngày ${cfg.date} (${declSummary}) và đã chuyển thư mục vào "${monthFolder}" để tránh xử lý lại.`,`Trạng thái chuyển: ĐÃ CHUYỂN (${movedDirs.join('; ')})`);
  }else{
    log('INFO','TO_KHAI_ARCHIVE',`Đã xử lý ${processedDecls.length} số tờ khai ngày ${cfg.date} (${declSummary}).`,`Trạng thái chuyển: KHÔNG CẦN CHUYỂN (đã nằm trong thư mục tháng hoặc chạy qua bộ nhớ test)`);
  }
}else{
  log('WARN','TO_KHAI_ARCHIVE',`Chưa chuyển thư mục tờ khai ngày ${cfg.date} (${declSummary}) do đối soát bị chặn (BLOCKED).`,`Trạng thái chuyển: CHƯA CHUYỂN — giữ nguyên tại thư mục gốc để kiểm tra và chạy lại.`);
}
const report={workflow:'InvoiceFlow n8n E2E', status, startedAt, finishedAt:new Date().toISOString(), requestedBy:cfg.requestedBy,
  cfg:{profile:cfg.profile,company:cfg.companyCode,date:cfg.date,customsPath:cfg.customsPath,templatePath:cfg.templatePath,outputFolder:cfg.outputFolder,
    SAPInvoiceEndpoint:cfg.sapInvoiceUrl,SAPRateEndpoint:cfg.sapRateUrl},
  sources:{customsFile:params.customs.fileName,customsSHA256:customsHash,customsSelection:params.customs.selection,templateFile:params.template.fileName},
  stats, rate:{USD_VND:fxRate,validOn:fxDate}, errors:defects,warnings,steps,
  control:{sapApiContractVerified:!!cfg.sapContractVerified,customsMappingVerified:!!cfg.customsMappingVerified,templateVerified:!!cfg.templateMappingVerified},
  publication:'NOT_UPLOADED_TO_VIETTEL',reviewRequired:true};
const reportJSON=JSON.stringify(report,null,2);
const reportCSV=['time,level,step,message,details',...steps.map(x=>[x.time,x.level,x.step,x.message,x.details].map(csv).join(','))].join('\r\n');
const reportHTML=`<!doctype html><html lang="vi"><head><meta charset="utf-8"><title>InvoiceFlow Audit ${escape(cfg.date)}</title><style>body{font-family:Arial,sans-serif;background:#f7f9fc;padding:28px;color:#10233f}main{max-width:1100px;margin:auto;background:white;padding:30px;border-radius:14px}table{width:100%;border-collapse:collapse}th,td{border-bottom:1px solid #e1e6ee;text-align:left;padding:9px;vertical-align:top}th{background:#f1f4fa}.ERROR{color:#bb172a}.WARN{color:#a76a10}.INFO{color:#126b3d}.summary{display:flex;gap:16px;flex-wrap:wrap}.summary>div{background:#eef3fa;padding:12px;border-radius:9px}code{overflow-wrap:anywhere}h1{margin-top:0}</style></head><body><main><h1>InvoiceFlow — Audit ${escape(status)}</h1><p><b>Kỳ:</b> ${escape(cfg.date)} · <b>Company:</b> ${escape(cfg.companyCode)} · <b>User:</b> ${escape(cfg.requestedBy)}</p><div class="summary"><div>Hải quan: <b>${stats.customsRows}</b></div><div>SAP: <b>${stats.sapRowsForDate}</b></div><div>Chi tiết: <b>${stats.details}</b></div><div>Ghi chú: <b>${stats.noteRows}</b></div><div>Lỗi: <b>${defects.length}</b></div><div>Cảnh báo: <b>${warnings.length}</b></div></div><h2>Thông số</h2><p>Tờ khai: <code>${escape(params.customs.fileName)}</code> · FX USD/VND: ${escape(fxRate)} · FX date: ${escape(fxDate)}</p><p><b>Không phát hành Viettel.</b> Phải kiểm tra và phê duyệt file Excel trước khi tải lên.</p><h2>Audit trail</h2><table><thead><tr><th>Thời gian</th><th>Mức</th><th>Bước</th><th>Thông tin</th><th>Chi tiết</th></tr></thead><tbody>${steps.map(s=>`<tr><td>${escape(s.time)}</td><td class="${escape(s.level)}">${escape(s.level)}</td><td>${escape(s.step)}</td><td>${escape(s.message)}</td><td>${escape(s.details)}</td></tr>`).join('')}</tbody></table><h2>Nguồn</h2><p>SHA-256 tờ khai: <code>${escape(customsHash)}</code></p><p>Run time: ${escape(startedAt)} → ${escape(report.finishedAt)}</p></main></body></html>`;
const binary={};
if(ExcelJS){
  const stepNames={
    SOURCE:'1. Đọc file đầu vào',
    CUSTOMS:'2. Kiểm tra tờ khai HQ',
    SAP_TCURR:'3. Kiểm tra tỷ giá USD/VND',
    SAP_FILTER:'4. Lọc dữ liệu SAP',
    RECONCILE:'5. Đối chiếu HQ & SAP',
    EXCEPTION_QTY_ZERO:'Ngoại lệ: Số lượng = 0',
    EXCEPTION_PIR_NA:'Lỗi: Thiếu giá PIR (#N/A)',
    EXCEPTION_AA_DIV10:'Ngoại lệ: Sửa đơn giá (/10)',
    EXCEPTION_AC_UNRESOLVED:'Lỗi: Lệch tiền chưa xử lý',
    TO_KHAI_APPEND:'6. Cập nhật sheet Tờ khai',
    EXCEPTION_TK_DUP_INVOICE:'Ngoại lệ: Trùng hóa đơn Tờ khai',
    EXCEPTION_MULTI_DECL_P:'Ngoại lệ: Hóa đơn nhiều tờ khai',
    MS_TIEN_INVOICE:'7. Cập nhật Ms Tiên invoices',
    OUTPUT:'8. Tạo file Upload Viettel',
    TO_KHAI_ARCHIVE:'9. Lưu trữ thư mục Tờ khai',
    TEMPLATE:'Lưu ý khi mở Excel',
    ENGINE:'Lỗi hệ thống'
  };
  const levelNames={INFO:'Thông tin',WARN:'Cảnh báo',ERROR:'Lỗi (Dừng)'};
  const fmtVnTime=(iso)=>{
    const d=new Date(Date.parse(iso)+7*3600000);
    if(!Number.isFinite(d.getTime()))return iso;
    const pad=n=>String(n).padStart(2,'0');
    return `${pad(d.getUTCDate())}/${pad(d.getUTCMonth()+1)}/${d.getUTCFullYear()} ${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}`;
  };
  const logWb=new ExcelJS.Workbook();
  const sumWs=logWb.addWorksheet('Tom_tat');
  sumWs.columns=[{header:'Mục',key:'k',width:32},{header:'Giá trị',key:'v',width:68}];
  const sumHdr=sumWs.getRow(1);
  if(sumHdr&&sumHdr.eachCell)sumHdr.eachCell(c=>{c.font={bold:true};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFEEF3FA'}};});
  const statusLabel=status==='READY_FOR_REVIEW'?'ĐẠT - Sẵn sàng kiểm tra & xuất Viettel':'BỊ CHẶN - Vui lòng xem chi tiết bên sheet Nhat_ky_xu_ly';
  for(const [k,v] of [
    ['Kết quả xử lý',statusLabel],['Ngày hóa đơn',cfg.date],['Mã công ty',cfg.companyCode],['Người thực hiện',cfg.requestedBy],
    ['File tờ khai HQ',params.customs.fileName],['Tỷ giá USD/VND',fxRate??''],['Ngày hiệu lực tỷ giá',fxDate??''],
    ['Số dòng tờ khai HQ',stats.customsRows],['Số dòng SAP trong ngày',stats.sapRowsForDate],
    ['Số dòng chi tiết (Upload Viettel)',stats.details],['Số dòng ghi chú (Upload Viettel)',stats.noteRows],
    ['Số lỗi cần xử lý',defects.length],['Số cảnh báo',warnings.length],
    ['Bắt đầu lúc',fmtVnTime(startedAt)],['Kết thúc lúc',fmtVnTime(report.finishedAt)]
  ]) sumWs.addRow({k,v});
  const logWs=logWb.addWorksheet('Nhat_ky_xu_ly');
  logWs.columns=[{header:'Thời gian',key:'time',width:20},{header:'Mức độ',key:'level',width:14},{header:'Bước thực hiện',key:'step',width:30},{header:'Nội dung chi tiết',key:'message',width:68},{header:'Ghi chú thêm',key:'details',width:52}];
  const hdr=logWs.getRow(1);
  if(hdr&&hdr.eachCell)hdr.eachCell(c=>{c.font={bold:true};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFEEF3FA'}};});
  for(const s of steps){
    const r=logWs.addRow({time:fmtVnTime(s.time),level:levelNames[s.level]||s.level,step:stepNames[s.step]||s.step,message:s.message,details:s.details});
    const lvlColor=s.level==='ERROR'?'FFBB172A':s.level==='WARN'?'FFA76A10':'FF126B3D';
    if(r&&r.getCell){
      r.getCell(2).font={bold:true,color:{argb:lvlColor}};
      r.getCell(4).alignment={wrapText:true,vertical:'top'};
      r.getCell(5).alignment={wrapText:true,vertical:'top'};
    }
  }
  const auditXlsxBuf=await logWb.xlsx.writeBuffer();
  try{const fs=require('fs');if(fs.existsSync(cfg.outputFolder))fs.writeFileSync(`${cfg.outputFolder}/${filenames.audit_xlsx}`,Buffer.from(auditXlsxBuf));}catch{}
  binary.audit_xlsx=await this.helpers.prepareBinaryData(Buffer.from(auditXlsxBuf),filenames.audit_xlsx,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}
if(status==='READY_FOR_REVIEW'&&xlsx)binary.excel=await this.helpers.prepareBinaryData(Buffer.from(xlsx),filenames.excel,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
return [{json:{status,summary:stats,errors:defects,warnings,report,filenames,outputPath:`${cfg.outputFolder}/${filenames.excel}`},binary}];
