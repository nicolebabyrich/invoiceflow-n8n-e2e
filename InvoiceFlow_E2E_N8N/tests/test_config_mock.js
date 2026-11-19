const assert=require('assert'),fs=require('fs');
const AsyncFunction=Object.getPrototypeOf(async function(){}).constructor;
const fn=new AsyncFunction('$input',fs.readFileSync(__dirname+'/../code/02_config.js','utf8'));
(async()=>{
 const run=async fields=>(await fn({first:()=>({json:fields})}))[0].json;
 const good=await run({Profile:'FINANCE_VN',Ngay_xu_ly:'2026-11-11',Ma_cong_ty:'HK13',Ghi_de_path_to_khai:'customs/2026/Baocaotokhai20261111.xlsx',Ghi_de_thu_muc_ket_qua:'results/2026'});
 assert.equal(good.cfg.monthStart,'2026-11-01');assert.equal(good.cfg.monthEnd,'2026-11-30');assert.ok(good.cfg.customsPath.endsWith('Baocaotokhai20261111.xlsx'));
 console.log('PASS date range, custom paths, company profile');
 await assert.rejects(run({Profile:'FINANCE_VN',Ngay_xu_ly:'2026-11-11',Ma_cong_ty:'HK13',Ghi_de_path_to_khai:'../../etc/passwd'}),/PATH_REJECTED/);
 console.log('PASS path traversal rejected');
 await assert.rejects(run({Profile:'FINANCE_VN',Ngay_xu_ly:'2026-02-30',Ma_cong_ty:'HK13'}),/DATE_INVALID/);
 console.log('PASS invalid date rejected');
})().catch(e=>{console.error(e);process.exit(1)});
