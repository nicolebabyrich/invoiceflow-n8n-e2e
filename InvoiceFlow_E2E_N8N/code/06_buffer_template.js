const it=$input.first();
if(!it.binary?.data) throw new Error('TEMPLATE_NOT_FOUND: Không đọc được mẫu Excel Viettel.');
const b=await this.helpers.getBinaryDataBuffer(0,'data');
return [{json:{...$('04 - Chon file to khai').first().json, template:{fileName:it.binary.data.fileName||'Viettel.xlsx',base64:b.toString('base64')}}}];
