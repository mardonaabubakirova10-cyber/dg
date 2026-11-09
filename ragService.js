const MAX_CHUNK = 600;
const TOP_K = 4;
const GEMINI_MODEL = process.env.GEMINI_MODEL || 'gemini-3.5-flash-lite';

// ---------- 5.2 Индексация: чанки из Word и из Excel ----------

// Короткие строки (заголовки, вопросы) склеиваются со следующим абзацем,
// чтобы чанк содержал и вопрос, и ответ.
function chunkDoc(doc) {
  const chunks = [];
  let pending = '';
  const flush = () => {
    if (pending.trim()) chunks.push({ source: doc.fileName, type: 'doc', text: pending.trim() });
    pending = '';
  };
  for (const p of doc.paragraphs) {
    const isHeading = p.length < 120;
    if (!isHeading && pending && pending.length + p.length > MAX_CHUNK) flush();
    pending += (pending ? '\n' : '') + p;
    if (!isHeading) flush();
  }
  flush();
  return chunks;
}

const fmt = (n) => Number(n).toLocaleString('ru-RU');

function courseChunk(c, instructors) {
  const inst = instructors.find(i => String(i.id) === String(c.instructor_id));
  const text = [
    `Курс «${c.title}» (${c.category}, уровень: ${c.level}, длительность: ${c.duration}).`,
    `Цена со скидкой ${c.discount_percent}%: ${fmt(c.price_discount)} сум (полная ${fmt(c.price_full)} сум), рассрочка от ${fmt(c.price_monthly)} сум/мес.`,
    `Навыки: ${c.tags}.`,
    c.description,
    inst ? `Наставник: ${inst.name}, ${inst.role}, ${inst.company}, ${inst.experience}.` : ''
  ].filter(Boolean).join(' ');
  return { source: 'courses_data.xlsx', type: 'course', courseId: c.id, courseTitle: c.title, text };
}

function buildIndex(docs, courses, instructors) {
  return [
    ...docs.flatMap(chunkDoc),
    ...courses.map(c => courseChunk(c, instructors))
  ];
}

// ---------- Поиск ----------

// Грубый «стемминг»: первые 5 букв слова, чтобы «рассрочка» находило «рассрочки».
// Стоп-слова: «курс» есть в каждом чанке-курсе и только шумит.
const STOP = new Set(['хочу', 'как', 'что', 'для', 'это', 'или', 'при', 'про', 'курс', 'курсы', 'есть', 'можно', 'меня', 'расскажи', 'подробно',
  'кто', 'какое', 'какой', 'какая', 'какие', 'какого', 'какую', 'сколько', 'будет', 'если', 'после', 'через', 'где', 'когда', 'зачем', 'почему', 'мне', 'нас', 'вас', 'они', 'этот', 'такое', 'ли']);

function stems(text) {
  return (text.toLowerCase().match(/[a-zа-яё0-9+#.]+/g) || [])
    .filter(w => w.length > 2 && !STOP.has(w))
    .map(w => w.slice(0, 5));
}

function retrieve(index, query) {
  const q = stems(query);
  if (!q.length) return [];
  return index
    .map(chunk => {
      const words = stems(chunk.text);
      const score = q.reduce((s, w) => s + (words.includes(w) ? 1 : 0), 0) / q.length;
      return { chunk, score };
    })
    .filter(r => r.score >= 0.34)
    .sort((a, b) => b.score - a.score)
    .slice(0, TOP_K)
    .map(r => r.chunk);
}

// ---------- 5.3 Генерация ответа ----------

function buildPrompt(query, chunks) {
  const context = chunks.map((c, i) => `[${i + 1}] (${c.source}) ${c.text}`).join('\n\n');
  return `Ты — консультант образовательной платформы Digital. Отвечай ТОЛЬКО по контексту ниже: не добавляй фактов, цифр и сроков, которых в нём нет, и не делай выводов сверх написанного. Если ответа в контексте нет, так и скажи.
Формат ответа строго такой, без markdown:
1 Краткий заголовок ответа
1.1 первый шаг или факт
1.2 второй шаг или факт
1.3 третий шаг или факт
1.4 итог или следующий шаг
В конце строк добавляй номер источника в квадратных скобках, например [1]. Если фактов меньше четырёх, используй меньше строк.

Контекст:
${context}

Вопрос: ${query}`;
}

async function askGemini(query, chunks) {
  const key = process.env.GEMINI_API_KEY;
  if (!key) return null;
  const request = () => fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-goog-api-key': key },
      body: JSON.stringify({
        contents: [{ role: 'user', parts: [{ text: buildPrompt(query, chunks) }] }],
        generationConfig: { temperature: 0.1 }
      }),
      signal: AbortSignal.timeout(20000)
    }
  );
  let res = await request();
  if (res.status === 503) res = await request(); // временная перегрузка Gemini
  if (!res.ok) throw new Error(`Gemini HTTP ${res.status}`);
  const data = await res.json();
  return data.candidates?.[0]?.content?.parts?.map(p => p.text).join('').trim() || null;
}

// Ответ без LLM: берём найденные чанки как есть, в формате шагов.
function extractiveAnswer(chunks) {
  const lines = chunks.slice(0, 4).map((c, i) => `1.${i + 1} ${c.text.replace(/\n+/g, ' — ')} [${i + 1}]`);
  return `1 Найдено в базе знаний\n${lines.join('\n')}`;
}

async function answer(query, index) {
  const chunks = retrieve(index, query);
  if (!chunks.length) return { chunks, text: null, mode: 'none' };

  try {
    const text = await askGemini(query, chunks);
    if (text) return { chunks, text, mode: 'gemini' };
  } catch (err) {
    console.error('[RAG] Gemini недоступен, использую ответ без LLM:', err.message);
  }
  return { chunks, text: extractiveAnswer(chunks), mode: 'extractive' };
}

module.exports = { buildIndex, retrieve, answer };
