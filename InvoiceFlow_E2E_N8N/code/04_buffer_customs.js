/** Select newest by timestamp/name per TK directory (supports multiple TK folders for the same date). */
const a = $input.all();
if (!a.length) throw new Error('CUSTOMS_NOT_FOUND: Không có file tờ khai trong path đã cấu hình.');
const byDir = new Map();
for (let idx = 0; idx < a.length; idx++) {
  const item = a[idx];
  const b = item.binary?.data;
  if (!b) throw new Error('CUSTOMS_MISSING_BINARY: Read file node chưa trả binary.data');
  const name = String(b.fileName || item.json?.fileName || `file_${idx}`);
  const dir = String(b.directory || item.json?.directory || '');
  const m = name.match(/20\d{2}[-_.]?\d{2}[-_.]?\d{2}(?:[-_.]?\d{2})?(?:[-_.]?\d{2})?(?:[-_.]?\d{2})?/);
  const stamp = m ? m[0].replace(/\D/g, '') : '';
  if (!byDir.has(dir)) byDir.set(dir, []);
  byDir.get(dir).push({ item, idx, name, dir, stamp });
}
const customsFiles = [];
for (const [dir, list] of byDir) {
  list.sort((x, y) => (y.stamp || y.name).localeCompare(x.stamp || x.name));
  const chosen = list[0];
  const buf = await this.helpers.getBinaryDataBuffer(chosen.idx, 'data');
  customsFiles.push({
    fileName: chosen.name,
    directory: dir,
    selection: list.length === 1 ? 'ONLY_MATCH' : chosen.stamp ? 'FILENAME_TIMESTAMP' : 'FILENAME_LEXICAL',
    candidates: list.length,
    base64: buf.toString('base64')
  });
}
const previous = $('02 - Config va bao ve path').first().json;
return [{ json: { ...previous, customs: customsFiles[0], customsFiles } }];
