// Сохранение заявок с сайта. Если задан GOOGLE_LEADS_SHEET_ID — строка добавляется в Google Таблицу
// (сервисный аккаунт должен быть её редактором). Иначе заявки пишутся в leads.jsonl рядом с проектом.
const fs = require('fs');
const path = require('path');

const SHEET_ID = process.env.GOOGLE_LEADS_SHEET_ID;
const LOCAL_FILE = path.join(__dirname, 'leads.jsonl');

async function appendToSheet(row) {
  const { google } = require('googleapis');
  const opts = { scopes: ['https://www.googleapis.com/auth/spreadsheets'] };
  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) opts.credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  else opts.keyFile = path.resolve(process.env.GOOGLE_KEY_FILE);
  const sheets = google.sheets({ version: 'v4', auth: new google.auth.GoogleAuth(opts) });
  // RAW: значения хранятся как текст, «=…» и «+998…» не превращаются в формулы
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: 'A1',
    valueInputOption: 'RAW',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [row] },
  });
}

async function saveLead({ type, name, phone, course }) {
  const row = [new Date().toISOString(), type, name, phone, course || ''];
  if (SHEET_ID) return appendToSheet(row);
  fs.appendFileSync(LOCAL_FILE, JSON.stringify(row) + '\n');
}

module.exports = { saveLead };
