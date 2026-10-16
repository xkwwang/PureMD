const fs = require('fs');
const file = process.argv[2];
const fd = fs.openSync(file, 'r');
const magic = Buffer.alloc(4);
fs.readSync(fd, magic, 0, 4, 0);

let prefix, hsOff;
if (magic.toString('latin1') === 'ASAR') { prefix = 8; hsOff = 4; }
else { prefix = 16; hsOff = 4; }

const sizeBuf = Buffer.alloc(4);
fs.readSync(fd, sizeBuf, 0, 4, hsOff);
const headerSize = sizeBuf.readUInt32LE(0);
const headerBuf = Buffer.alloc(headerSize);
fs.readSync(fd, headerBuf, 0, headerSize, prefix);
const header = JSON.parse(headerBuf.toString('utf8'));
const dataStart = Math.ceil((prefix + headerSize) / 4096) * 4096;

const hits = [];
function walk(node, p) {
  for (const [name, info] of Object.entries(node)) {
    const full = p ? p + '/' + name : name;
    if (/(markdown|preview).*\.(css|js)$/i.test(full)) hits.push(full);
    if (info.files) walk(info.files, full);
  }
}
walk(header.files, '');
fs.writeFileSync('D:/00. APP/notepad++/hits.txt', hits.join('\n'));
console.log('TOTAL_HITS', hits.length);
console.log(hits.slice(0, 60).join('\n'));

if (process.argv[3] && process.argv[4]) {
  const target = process.argv[3];
  let cur = header.files;
  const parts = target.split('/');
  let info = null;
  for (const part of parts) {
    if (!cur[part]) { console.log('NOT_FOUND', target); process.exit(0); }
    info = cur[part];
    if (info.files) cur = info.files;
  }
  const offset = dataStart + Number(info.offset || 0);
  const size = Number(info.size || 0);
  const out = Buffer.alloc(size);
  fs.readSync(fd, out, 0, size, offset);
  fs.writeFileSync(process.argv[4], out);
  console.log('EXTRACTED', target, '->', process.argv[4], size, 'bytes');
}
