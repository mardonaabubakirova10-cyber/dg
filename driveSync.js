// Синхронизация папки Google Drive (через сервисный аккаунт) в локальный кеш.
// Включается, если задан GOOGLE_DRIVE_FOLDER_ID и ключ (GOOGLE_SERVICE_ACCOUNT_JSON или GOOGLE_KEY_FILE).
const fs = require('fs');
const os = require('os');
const path = require('path');

const FOLDER_ID = process.env.GOOGLE_DRIVE_FOLDER_ID;
const CACHE_DIR = path.join(os.tmpdir(), 'dg_drive_cache');
const WORD_DIR = path.join(CACHE_DIR, 'knowledge_base_word');
const INTERVAL_MS = 30 * 1000;
const GDOC = 'application/vnd.google-apps.document';
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

let drive = null;
function getDrive() {
  if (drive) return drive;
  const { google } = require('googleapis');
  const opts = { scopes: ['https://www.googleapis.com/auth/drive.readonly'] };
  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) opts.credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  else opts.keyFile = path.resolve(process.env.GOOGLE_KEY_FILE);
  drive = google.drive({ version: 'v3', auth: new google.auth.GoogleAuth(opts) });
  return drive;
}

const enabled = Boolean(FOLDER_ID && (process.env.GOOGLE_SERVICE_ACCOUNT_JSON || process.env.GOOGLE_KEY_FILE));
const listOpts = { supportsAllDrives: true, includeItemsFromAllDrives: true };

async function list(parentId) {
  const r = await getDrive().files.list({
    q: `'${parentId}' in parents and trashed=false`,
    fields: 'files(id,name,mimeType,modifiedTime)',
    ...listOpts,
  });
  return r.data.files;
}

// Скачивает файл, только если он изменился (сравнение по modifiedTime Drive).
const seen = {};
async function fetchFile(file, dest) {
  if (seen[file.id] === file.modifiedTime && fs.existsSync(dest)) return;
  const req = file.mimeType === GDOC
    ? getDrive().files.export({ fileId: file.id, mimeType: DOCX }, { responseType: 'arraybuffer' })
    : getDrive().files.get({ fileId: file.id, alt: 'media', supportsAllDrives: true }, { responseType: 'arraybuffer' });
  const res = await req;
  fs.mkdirSync(path.dirname(dest), { recursive: true });
  fs.writeFileSync(dest, Buffer.from(res.data));
  seen[file.id] = file.modifiedTime;
}

async function syncOnce() {
  const files = await list(FOLDER_ID);
  const xlsx = files.find(f => f.name === 'courses_data.xlsx');
  if (xlsx) await fetchFile(xlsx, path.join(CACHE_DIR, 'courses_data.xlsx'));

  const wordFolder = files.find(f => f.name === 'knowledge_base_word' && f.mimeType.endsWith('folder'));
  if (wordFolder) {
    const docs = (await list(wordFolder.id)).filter(f => f.name.endsWith('.docx') || f.mimeType === GDOC);
    const names = new Set();
    for (const f of docs) {
      const name = f.name.endsWith('.docx') ? f.name : `${f.name}.docx`;
      names.add(name);
      await fetchFile(f, path.join(WORD_DIR, name));
    }
    // удалённые в Drive документы убираем из кеша
    if (fs.existsSync(WORD_DIR)) {
      for (const f of fs.readdirSync(WORD_DIR)) if (!names.has(f)) fs.unlinkSync(path.join(WORD_DIR, f));
    }
  }
}

// Не чаще раза в 30 секунд; при ошибке остаётся прежний кеш.
let last = 0;
let running = null;
function refresh() {
  if (!enabled) return Promise.resolve();
  if (running) return running;
  if (Date.now() - last < INTERVAL_MS) return Promise.resolve();
  running = syncOnce()
    .catch(err => console.error('[DriveSync] Ошибка синхронизации:', err.message))
    .finally(() => { last = Date.now(); running = null; });
  return running;
}

module.exports = { enabled, refresh, CACHE_DIR };
