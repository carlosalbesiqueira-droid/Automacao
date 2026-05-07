const fs = require('node:fs/promises');
const path = require('node:path');

let lastTimestampBase = '';
let lastTimestampSequence = 0;

async function ensureDirectory(directoryPath) {
  await fs.mkdir(directoryPath, { recursive: true });
  return directoryPath;
}

function safeFileName(fileName) {
  return String(fileName ?? 'arquivo')
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, '_')
    .replace(/\s+/g, ' ')
    .trim();
}

function makeTimestamp(date = new Date()) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  const milliseconds = String(date.getMilliseconds()).padStart(3, '0');
  const base = `${year}${month}${day}_${hours}${minutes}${seconds}_${milliseconds}`;

  if (base === lastTimestampBase) {
    lastTimestampSequence += 1;
  } else {
    lastTimestampBase = base;
    lastTimestampSequence = 0;
  }

  return lastTimestampSequence
    ? `${base}_${String(lastTimestampSequence).padStart(2, '0')}`
    : base;
}

async function writeTextFile(filePath, content) {
  await ensureDirectory(path.dirname(filePath));
  await fs.writeFile(filePath, content, 'utf8');
  return filePath;
}

module.exports = {
  ensureDirectory,
  makeTimestamp,
  safeFileName,
  writeTextFile
};
