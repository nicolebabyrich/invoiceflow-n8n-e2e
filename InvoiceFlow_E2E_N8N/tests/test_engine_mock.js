// Runtime-free CONTRACT tests. Fake ExcelJS is intentional; NOT a test with SAP, n8n or actual XLSX formatting.
const fs=require('fs'),assert=require('assert'),crypto=require('crypto');
const AsyncFunction=Object.getPrototypeOf(async function(){}).constructor;
const code=fs.readFileSync(__dirname+'/../code/09_engine.js','utf8');
function colNum(x){if(typeof x==='number')return x;return String(x).toUpperCase().split('').reduce((n,c)=>n*26+c.charCodeAt(0)-64,0)}
class Cell{constructor(row,col){this.row=row;this.col=col;this.style={}}get value(){return this.row.data[this.col]??null}set value(v){this.row.data[this.col]=v}}
class Row{constructor(n,vals={}){this.number=n;this.data={...vals};this.cells={};this.height=20}getCell(col){col=colNum(col);return this.cells[col]||(this.cells[col]=new Cell(this,col))}get hasValues(){return Object.values(this.data).some(v=>v!=null&&v!=='')}}
class Sheet{constructor(data){this.rows={};for(const [r,values] of Object.entries(data))this.rows[r]=new Row(+r,values)}getRow(r){return this.rows[r]||(this.rows[r]=new Row(r))}addRow(vals){const next=(this.lastRow?this.lastRow.number:0)+1;return this.rows[next]=new Row(next,vals)}getCell(addr){const m=String(addr).match(/^([A-Z]+)(\d+)$/);return this.getRow(+m[2]).getCell(m[1])}eachRow(opts,fn){if(typeof opts==='function'){fn=opts;}for(const r of Object.values(this.rows).sort((a,b)=>a.number-b.number))if(r.hasValues)fn(r,r.number)}get lastRow(){const a=Object.values(this.rows).sort((a,b)=>b.number-a.number)[0];return a}}
class Workbook{constructor(){this.calcProperties={};this.sheets=[]}get worksheets(){return this.sheets||[]}addWorksheet(name){const s=Object.assign(new Sheet({}),{name});this.sheets.push(s);return s}getWorksheet(name){return this.worksheets.find(s=>s.name===name)}xlsx={load:async(b)=>{const raw=JSON.parse(Buffer.from(b).toString());this.sheets=Object.entries(raw).map(([name,rows])=>Object.assign(new Sheet(rows),{name}));},writeBuffer:async()=>Buffer.from('XLSX_TEST_BYTES')}}
const fakeRequire=(id)=>id==='exceljs'?{Workbook}:id==='crypto'?crypto:require(id);
function customRoot(delta){return {'Sheet1':{'1':{'1':'DIFF','2':'DECL','3':'INV','4':'DATE'},'2':{'1':delta,'2':'308853198750','3':'2300009848/9915','4':'2026-11-11'}}}}
function template(){return {'Sheet1':{'11':{'4':'D','5':'E','6':'F','7':'G','8':'H','22':'V'}}}}
const params={cfg:{companyCode:'HK13',date:'2026-11-11',requestedBy:'test',profile:'FINANCE_VN',customsPath:'/files/finance/source',templatePath:'/files/finance/template',outputFolder:'/files/finance/results',sapInvoiceUrl:'https://sap.local/invoices',sapRateUrl:'https://sap.local/rate',
 customsMappingVerified:true,sapContractVerified:true,fxContractVerified:true,templateMappingVerified:true,
 customs:{sheetName:'',headerRow:1,colInvoice:'C',colDate:'D',colDeclaration:'B',colCheckDifference:'A',colAmount:''},
 sap:{rowsPath:'value',invoice:'invoice',date:'billingDate',F:'contentF',G:'contentG',K:'fieldK',L:'quantity',M:'fieldM',N:'fieldN',P:'declarationText',reconcileDifference:'reconcileDifference'},
 fx:{rate:'rateVndPerUsd',validOn:'validOn'},templateSheet:'',toleranceCustoms:.03,toleranceSAP:0.000001,requireCustomsDifference:true,requireSAPDifference:true,requireCustomsDate:true,requireFXValidOn:true,maxRows:20000},
 customs:{fileName:'Baocaotokhai20261111.xlsx',base64:Buffer.from(JSON.stringify(customRoot(0))).toString('base64'),candidates:1,selection:'ONLY_MATCH'},
 template:{fileName:'template.xlsx',base64:Buffer.from(JSON.stringify(template())).toString('base64')}};
const sap={value:[{invoice:'2300009848',billingDate:'2026-11-11',contentF:'F content',contentG:'G content',fieldK:'SKU-01',quantity:10,fieldM:120,fieldN:1200,reconcileDifference:0},
{invoice:'2300009915',billingDate:'2026-11-11',contentF:'F2',contentG:'G2',fieldK:'SKU-02',quantity:0,fieldM:120,fieldN:0,reconcileDifference:0}]};
async function run(localParams=JSON.parse(JSON.stringify(params))){
 const all={'06 - Buffer template':[{json:localParams}],'07 - SAP YINVLST API':[{json:sap}],'08 - SAP TCURR API':[{json:{rateVndPerUsd:25890,validOn:'2026-11-10'}}]};
 const dollar=(name)=>({first:()=>all[name][0],all:()=>all[name]});
 const self={helpers:{prepareBinaryData:async(buffer,name,mime)=>({data:Buffer.from(buffer).toString('base64'),fileName:name,mimeType:mime})}};
 const fn=new AsyncFunction('$','require','Buffer',code);
 return (await fn.call(self,dollar,fakeRequire,Buffer))[0];
}
(async()=>{
 const good=await run();assert.equal(good.json.status,'READY_FOR_REVIEW');assert.equal(good.json.summary.details,1);assert.equal(good.json.summary.noteRows,1);assert.ok(good.binary.excel);assert.ok(good.binary.audit_xlsx);
 console.log('PASS valid customs+SAP+FX, one note row, Excel emitted');
 const wrong=JSON.parse(JSON.stringify(params));wrong.customs.base64=Buffer.from(JSON.stringify(customRoot(0.15))).toString('base64');
 const bad=await run(wrong);assert.equal(bad.json.status,'BLOCKED');assert.ok(!bad.binary.excel);assert.ok(bad.binary.audit_xlsx);
 console.log('PASS customs delta > 0.03 blocks Excel and emits logs');
 const unapproved=JSON.parse(JSON.stringify(params));unapproved.cfg.customsMappingVerified=false;
 const blocked=await run(unapproved);assert.equal(blocked.json.status,'BLOCKED');assert.ok(blocked.json.errors.some(e=>e.includes('CONFIG_NOT_APPROVED')));
 console.log('PASS unverified mapping fails closed');
})().catch(e=>{console.error(e.stack);process.exit(1)});
