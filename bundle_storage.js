// Упаковывает storage_drive/ в storage_bundle.json, чтобы данные гарантированно попадали в деплой на Vercel.
// Запуск: npm run bundle-storage (на Vercel выполняется автоматически как build).
const fs = require('fs');
const path = require('path');

const dir = path.join(__dirname, 'storage_drive');
const wordDir = path.join(dir, 'knowledge_base_word');

const bundle = {
  excel: fs.readFileSync(path.join(dir, 'courses_data.xlsx')).toString('base64'),
  docs: {}
};

for (const file of fs.readdirSync(wordDir)) {
  if (file.endsWith('.docx') && !file.startsWith('~$')) {
    bundle.docs[file] = fs.readFileSync(path.join(wordDir, file)).toString('base64');
  }
}

fs.writeFileSync(path.join(__dirname, 'storage_bundle.json'), JSON.stringify(bundle));
console.log(`[bundle] Excel + ${Object.keys(bundle.docs).length} docx упакованы в storage_bundle.json`);
