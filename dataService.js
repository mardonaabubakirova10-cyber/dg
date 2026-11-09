const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const mammoth = require('mammoth');

// DRIVE_FOLDER — путь к папке, синхронизируемой Google Drive для компьютера
// (внутри: courses_data.xlsx и папка knowledge_base_word/). По умолчанию — локальная storage_drive.
const STORAGE_DIR = process.env.DRIVE_FOLDER
  ? path.resolve(process.env.DRIVE_FOLDER)
  : path.join(__dirname, 'storage_drive');
const EXCEL_PATH = path.join(STORAGE_DIR, 'courses_data.xlsx');
const WORD_DIR = path.join(STORAGE_DIR, 'knowledge_base_word');

// Запасной источник для Vercel: данные, упакованные в JSON (см. bundle_storage.js).
// Используется, только если на диске нет файлов (нет локальной папки или Drive).
let bundle = null;
try {
  bundle = require('./storage_bundle.json');
} catch (err) {
  // файла нет — работаем только с диска
}

// Кеш для данных и документов
let dataCache = {
  courses: [],
  instructors: [],
  lastLoaded: null
};

let knowledgeDocs = [];

/**
 * Чтение Excel файла (courses_data.xlsx)
 * При каждом чтении актуализирует данные из листов Courses и Instructors
 */
function loadExcelData() {
  try {
    let workbook;
    if (fs.existsSync(EXCEL_PATH)) {
      workbook = xlsx.readFile(EXCEL_PATH);
    } else if (bundle) {
      workbook = xlsx.read(Buffer.from(bundle.excel, 'base64'), { type: 'buffer' });
    } else {
      console.warn(`[DataService] Excel файл не найден по пути: ${EXCEL_PATH}`);
      return dataCache;
    }

    // Лист Courses
    if (workbook.SheetNames.includes('Courses')) {
      const sheet = workbook.Sheets['Courses'];
      dataCache.courses = xlsx.utils.sheet_to_json(sheet, { defval: "" });
    }

    // Лист Instructors
    if (workbook.SheetNames.includes('Instructors')) {
      const sheet = workbook.Sheets['Instructors'];
      dataCache.instructors = xlsx.utils.sheet_to_json(sheet, { defval: "" });
    }

    dataCache.lastLoaded = new Date().toISOString();
    console.log(`[DataService] Данные из Excel обновлены: ${dataCache.courses.length} курсов, ${dataCache.instructors.length} преподавателей`);
    return dataCache;
  } catch (err) {
    console.error("[DataService] Ошибка чтения Excel файла:", err.message);
    return dataCache; // возвращаем последний рабочий кеш
  }
}

/**
 * Чтение Word (.docx) файлов из папки knowledge_base_word для RAG
 */
async function loadWordDocs() {
  try {
    const onDisk = fs.existsSync(WORD_DIR);
    if (!onDisk && !bundle) return [];
    const files = onDisk
      ? fs.readdirSync(WORD_DIR).filter(f => f.endsWith('.docx') && !f.startsWith('~$'))
      : Object.keys(bundle.docs);
    const parsed = [];

    for (const file of files) {
      const filePath = path.join(WORD_DIR, file);
      const buffer = onDisk ? fs.readFileSync(filePath) : Buffer.from(bundle.docs[file], 'base64');
      const result = await mammoth.extractRawText({ buffer });
      const text = result.value || "";

      parsed.push({
        fileName: file,
        filePath,
        text,
        paragraphs: text.split('\n').filter(p => p.trim().length > 0)
      });
    }

    knowledgeDocs = parsed;
    console.log(`[DataService] Загружено Word документов базы знаний: ${knowledgeDocs.length}`);
    return knowledgeDocs;
  } catch (err) {
    console.error("[DataService] Ошибка чтения Word файлов:", err.message);
    return knowledgeDocs;
  }
}

// Проверка даты изменения файла для автоматического обновления кеша
let lastMtime = null;

function getFreshExcelData() {
  try {
    if (!fs.existsSync(EXCEL_PATH)) return dataCache;
    const stats = fs.statSync(EXCEL_PATH);
    if (!lastMtime || stats.mtimeMs > lastMtime) {
      lastMtime = stats.mtimeMs;
      loadExcelData();
    }
  } catch (err) {
    // игнорируем кратковременные ошибки чтения при перезаписи
  }
  return dataCache;
}

// Начальная загрузка
loadExcelData();
loadWordDocs();

module.exports = {
  getExcelData: getFreshExcelData,
  getKnowledgeDocs: async () => {
    await loadWordDocs();
    return knowledgeDocs;
  },
  EXCEL_PATH,
  WORD_DIR
};
