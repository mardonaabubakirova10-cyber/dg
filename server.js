const express = require('express');
const cors = require('cors');
const path = require('path');
require('dotenv').config({ quiet: true });
const driveSync = require('./driveSync');
const dataService = require('./dataService');
const leads = require('./leads');
const ragService = require('./ragService');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Подтягиваем свежие файлы с Google Drive (не чаще раза в 30 с)
app.use('/api', async (req, res, next) => { await driveSync.refresh(); next(); });

// API: Получить курсы с опциональной фильтрацией
app.get('/api/courses', (req, res) => {
  try {
    const { courses, instructors } = dataService.getExcelData();
    const { category, level, search } = req.query;

    let filtered = [...courses];

    if (category && category !== 'Все') {
      filtered = filtered.filter(c => c.category && c.category.toLowerCase() === category.toLowerCase());
    }

    if (level && level !== 'Все') {
      filtered = filtered.filter(c => c.level && c.level.toLowerCase() === level.toLowerCase());
    }

    if (search) {
      const q = search.toLowerCase();
      filtered = filtered.filter(c => 
        (c.title && c.title.toLowerCase().includes(q)) ||
        (c.description && c.description.toLowerCase().includes(q)) ||
        (c.tags && c.tags.toLowerCase().includes(q))
      );
    }

    // Соединяем с данными преподавателя
    const enriched = filtered.map(course => {
      const instructor = instructors.find(i => String(i.id) === String(course.instructor_id));
      return {
        ...course,
        instructor: instructor || null
      };
    });

    res.json({
      success: true,
      count: enriched.length,
      data: enriched
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// API: Получить преподавателей
app.get('/api/instructors', (req, res) => {
  try {
    const { instructors, courses } = dataService.getExcelData();
    
    // Добавляем список курсов каждого преподавателя
    const enriched = instructors.map(inst => {
      const taughtCourses = courses.filter(c => String(c.instructor_id) === String(inst.id));
      return {
        ...inst,
        courses: taughtCourses.map(c => ({ id: c.id, title: c.title, category: c.category }))
      };
    });

    res.json({
      success: true,
      count: enriched.length,
      data: enriched
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// API: Получить категории и метаданные каталога
app.get('/api/meta', (req, res) => {
  try {
    const { courses } = dataService.getExcelData();
    const categories = Array.from(new Set(courses.map(c => c.category).filter(Boolean)));
    const levels = Array.from(new Set(courses.map(c => c.level).filter(Boolean)));

    res.json({
      success: true,
      categories: ['Все', ...categories],
      levels: ['Все', ...levels]
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// API: заявка на курс или консультацию (из модалки на сайте)
const leadHits = new Map();
app.post('/api/leads', async (req, res) => {
  const ip = (req.headers['x-forwarded-for'] || req.ip || '').split(',')[0].trim();
  const now = Date.now();
  const hits = (leadHits.get(ip) || []).filter(t => now - t < 60 * 1000);
  if (hits.length >= 5) return res.status(429).json({ success: false, error: 'Слишком много заявок. Подождите минуту.' });
  leadHits.set(ip, [...hits, now]);

  const name = String(req.body.name || '').trim().slice(0, 100);
  const phone = String(req.body.phone || '').trim().slice(0, 30);
  const digits = phone.replace(/\D/g, '');
  if (!name || digits.length < 9 || digits.length > 15) {
    return res.status(400).json({ success: false, error: 'Укажите имя и корректный номер телефона.' });
  }
  const course = dataService.getExcelData().courses.find(c => String(c.id) === String(req.body.course_id));
  try {
    await leads.saveLead({ type: course ? 'Запись на курс' : 'Консультация', name, phone, course: course ? course.title : '' });
    res.json({ success: true });
  } catch (err) {
    console.error('[Leads] Не удалось сохранить заявку:', err.message);
    res.status(500).json({ success: false, error: 'Не удалось отправить заявку. Попробуйте позже.' });
  }
});

// API: RAG Query (поиск по Word-документам и Excel)
// Возвращает ответ в структурированном виде шагов:
// 1 Описание шага 1.1 1.2 1.3 1.4
// Лимит: не более RAG_LIMIT запросов с одного IP за минуту (защита квоты Gemini).
// Счётчик в памяти: на Vercel он отдельный для каждого экземпляра функции.
const RAG_LIMIT = 10;
const RAG_WINDOW_MS = 60 * 1000;
const ragHits = new Map();

function ragRateLimit(req, res, next) {
  const ip = (req.headers['x-forwarded-for'] || req.ip || '').split(',')[0].trim();
  const now = Date.now();
  const hits = (ragHits.get(ip) || []).filter(t => now - t < RAG_WINDOW_MS);
  if (hits.length >= RAG_LIMIT) {
    return res.status(429).json({ success: false, error: 'Слишком много запросов. Подождите минуту.' });
  }
  hits.push(now);
  ragHits.set(ip, hits);
  if (ragHits.size > 1000) {
    for (const [key, times] of ragHits) {
      if (!times.some(t => now - t < RAG_WINDOW_MS)) ragHits.delete(key);
    }
  }
  next();
}

app.get('/api/rag/query', ragRateLimit, async (req, res) => {
  try {
    const query = (req.query.q || "").toLowerCase().trim();
    const docs = await dataService.getKnowledgeDocs();
    const { courses, instructors } = dataService.getExcelData();

    if (!query) {
      return res.json({
        success: true,
        answer: "1 Описание шага\n1.1 Задайте вопрос по интересующей профессии или навыкам (например: «Хочу войти в IT с нуля» или «Сколько длится рассрочка?»).\n1.2 Консультант ответит по документам школы и каталогу курсов.\n1.3 Вы получите пошаговый план обучения с точной стоимостью и менторами.\n1.4 Мы подберем подходящую программу с гарантией трудоустройства.",
        sources: []
      });
    }

    const index = ragService.buildIndex(docs, courses, instructors);
    const result = await ragService.answer(query, index);

    if (!result.text) {
      return res.json({
        success: true,
        mode: 'none',
        answer: "1 Не нашёл ответа в базе знаний\n1.1 В каталоге представлены направления: Программирование, Дизайн, Маркетинг, Управление и Игры.\n1.2 Уточните вопрос: назовите профессию, навык, цену или условия обучения.\n1.3 Можно также задать вопрос о рассрочке, гарантии трудоустройства или программе обучения.",
        sources: [],
        citations: [],
        courses: []
      });
    }

    res.json({
      success: true,
      mode: result.mode,
      answer: result.text,
      sources: [...new Set(result.chunks.map(c => c.source))],
      citations: result.chunks.map((c, i) => ({ n: i + 1, source: c.source, text: c.text.slice(0, 220) })),
      courses: result.chunks.filter(c => c.type === 'course').map(c => ({ id: c.courseId, title: c.courseTitle }))
    });

  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// Локально слушаем порт; на Vercel приложение подключается через api/index.js
if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`[Server] Сервер digital курсов запущен на http://localhost:${PORT}`);
  });
}

module.exports = app;
