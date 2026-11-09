const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const docx = require('docx');
const { Document, Packer, Paragraph, TextRun, HeadingLevel } = docx;

const storageDir = path.join(__dirname, 'storage_drive');
const knowledgeDir = path.join(storageDir, 'knowledge_base_word');

if (!fs.existsSync(storageDir)) fs.mkdirSync(storageDir, { recursive: true });
if (!fs.existsSync(knowledgeDir)) fs.mkdirSync(knowledgeDir, { recursive: true });

// 1. Создание courses_data.xlsx
const courses = [
  {
    id: "course-1",
    title: "Профессия Fullstack-разработчик на JavaScript и Python",
    category: "Программирование",
    level: "С нуля",
    duration: "12 месяцев",
    price_full: 27600000,
    price_discount: 16560000,
    price_monthly: 690000,
    discount_percent: 40,
    instructor_id: "inst-1",
    badge: "Хит продаж",
    tags: "JavaScript, React, Node.js, Python, PostgreSQL, Docker",
    rating: 4.9,
    reviews_count: 842,
    description: "Комплексная программа подготовки востребованного веб-разработчика: фронтенд на React и бэкенд на Python/Node.js с гарантией трудоустройства."
  },
  {
    id: "course-2",
    title: "UX/UI-дизайнер с нуля до PRO",
    category: "Дизайн",
    level: "Для начинающих",
    duration: "9 месяцев",
    price_full: 24000000,
    price_discount: 14400000,
    price_monthly: 600000,
    discount_percent: 40,
    instructor_id: "inst-2",
    badge: "Топ выбор",
    tags: "Figma, UX Research, Design Systems, Прототипирование, CJM",
    rating: 4.95,
    reviews_count: 620,
    description: "Научитесь проектировать продуманные мобильные приложения и сложные веб-сервисы, создавать дизайн-системы и тестировать интерфейсы."
  },
  {
    id: "course-3",
    title: "Data Scientist и Инженер искусственного интеллекта",
    category: "Программирование",
    level: "С нуля",
    duration: "14 месяцев",
    price_full: 31500000,
    price_discount: 18900000,
    price_monthly: 790000,
    discount_percent: 40,
    instructor_id: "inst-3",
    badge: "Высокая зарплата",
    tags: "Machine Learning, LLM, Neural Networks, PyTorch, SQL, Pandas",
    rating: 4.88,
    reviews_count: 415,
    description: "Освойте машинное обучение, глубокие нейросети, работу с генеративными LLM и RAG-системами на реальных прикладных бизнес-кейсах."
  },
  {
    id: "course-4",
    title: "Интернет-маркетолог: трафик, сквозная аналитика и бренд",
    category: "Маркетинг",
    level: "С нуля",
    duration: "8 месяцев",
    price_full: 21000000,
    price_discount: 12600000,
    price_monthly: 530000,
    discount_percent: 40,
    instructor_id: "inst-4",
    badge: "Быстрый старт",
    tags: "Контекст, Таргет, Яндекс.Директ, SEO, Воронки продаж",
    rating: 4.86,
    reviews_count: 531,
    description: "Научитесь привлекать целевых клиентов, настраивать рекламу в ключевых каналах и выстраивать окупаемые digital-стратегии."
  },
  {
    id: "course-5",
    title: "Product Manager: управление IT-продуктами",
    category: "Управление",
    level: "Продвинутый",
    duration: "10 месяцев",
    price_full: 29300000,
    price_discount: 17580000,
    price_monthly: 730000,
    discount_percent: 40,
    instructor_id: "inst-5",
    badge: "Премиум",
    tags: "CustDev, Unit-экономика, Agile, Scrum, Roadmaps, A/B тесты",
    rating: 4.92,
    reviews_count: 290,
    description: "Полный цикл вывода цифровых продуктов на рынок: проверка гипотез, расчёт юнит-экономики и лидерство в кросс-функциональных командах."
  },
  {
    id: "course-6",
    title: "Разработчик игр на Unity и C#",
    category: "Игры",
    level: "С нуля",
    duration: "11 месяцев",
    price_full: 26300000,
    price_discount: 15780000,
    price_monthly: 660000,
    discount_percent: 40,
    instructor_id: "inst-1",
    badge: "Популярно",
    tags: "Unity 3D, C#, 2D/3D физика, Шейдеры, Мультиплеер",
    rating: 4.91,
    reviews_count: 380,
    description: "Создайте 4 полноценных игровых проекта для ПК и смартфонов с собственной механикой, физикой и сетевым кодом."
  },
  {
    id: "course-7",
    title: "Python-разработчик: backend на FastAPI",
    category: "Программирование",
    level: "С нуля",
    duration: "8 месяцев",
    price_full: 18400000,
    price_discount: 11040000,
    price_monthly: 460000,
    discount_percent: 40,
    instructor_id: "inst-6",
    badge: "Востребовано",
    tags: "Python, FastAPI, PostgreSQL, Docker, REST API, Git",
    rating: 4.87,
    reviews_count: 312,
    description: "Научитесь писать серверную часть приложений на Python: API, базы данных, авторизация и деплой на сервер."
  },
  {
    id: "course-8",
    title: "Frontend-разработчик на React",
    category: "Программирование",
    level: "С нуля",
    duration: "7 месяцев",
    price_full: 16100000,
    price_discount: 10470000,
    price_monthly: 440000,
    discount_percent: 35,
    instructor_id: "inst-1",
    badge: "Новый поток",
    tags: "HTML, CSS, JavaScript, React, TypeScript, Git",
    rating: 4.84,
    reviews_count: 268,
    description: "Освойте вёрстку и современный JavaScript, соберите несколько интерфейсов на React и портфолио для первого трудоустройства."
  },
  {
    id: "course-9",
    title: "QA-инженер: ручное и автоматизированное тестирование",
    category: "Программирование",
    level: "Для начинающих",
    duration: "6 месяцев",
    price_full: 13800000,
    price_discount: 9660000,
    price_monthly: 400000,
    discount_percent: 30,
    instructor_id: "inst-6",
    badge: "Быстрый старт",
    tags: "Тест-дизайн, Postman, SQL, Selenium, Jira",
    rating: 4.81,
    reviews_count: 196,
    description: "Научитесь находить ошибки в продуктах, писать тест-кейсы и автотесты. Один из самых коротких путей в IT."
  },
  {
    id: "course-10",
    title: "Мобильный разработчик на Flutter",
    category: "Программирование",
    level: "Для начинающих",
    duration: "8 месяцев",
    price_full: 18400000,
    price_discount: 11960000,
    price_monthly: 500000,
    discount_percent: 35,
    instructor_id: "inst-6",
    badge: "Хит продаж",
    tags: "Dart, Flutter, Firebase, REST API, Git",
    rating: 4.86,
    reviews_count: 174,
    description: "Создавайте приложения для Android и iOS на одной кодовой базе и публикуйте их в магазинах приложений."
  },
  {
    id: "course-11",
    title: "Графический дизайнер с нуля",
    category: "Дизайн",
    level: "С нуля",
    duration: "6 месяцев",
    price_full: 13800000,
    price_discount: 8280000,
    price_monthly: 350000,
    discount_percent: 40,
    instructor_id: "inst-7",
    badge: "Популярно",
    tags: "Photoshop, Illustrator, Типографика, Брендинг, Композиция",
    rating: 4.9,
    reviews_count: 354,
    description: "Освойте композицию, цвет и типографику, делайте логотипы, фирменный стиль и макеты для печати и соцсетей."
  },
  {
    id: "course-12",
    title: "Motion-дизайнер: анимация и видеографика",
    category: "Дизайн",
    level: "Для начинающих",
    duration: "5 месяцев",
    price_full: 11500000,
    price_discount: 7480000,
    price_monthly: 310000,
    discount_percent: 35,
    instructor_id: "inst-7",
    badge: "Новый курс",
    tags: "After Effects, Анимация, Монтаж, Видеографика",
    rating: 4.83,
    reviews_count: 142,
    description: "Научитесь делать анимированные заставки, рекламные ролики и анимацию интерфейсов в After Effects."
  },
  {
    id: "course-13",
    title: "3D-художник в Blender",
    category: "Дизайн",
    level: "С нуля",
    duration: "7 месяцев",
    price_full: 16100000,
    price_discount: 11270000,
    price_monthly: 470000,
    discount_percent: 30,
    instructor_id: "inst-7",
    badge: "Творческий",
    tags: "Blender, 3D-моделирование, Текстуры, Рендер, Анимация",
    rating: 4.88,
    reviews_count: 121,
    description: "Моделируйте персонажей и предметы, настраивайте свет и материалы, готовьте работы для игр, рекламы и портфолио."
  },
  {
    id: "course-14",
    title: "SMM-менеджер: продвижение в Instagram и Telegram",
    category: "Маркетинг",
    level: "С нуля",
    duration: "5 месяцев",
    price_full: 11500000,
    price_discount: 6900000,
    price_monthly: 290000,
    discount_percent: 40,
    instructor_id: "inst-8",
    badge: "Быстрый старт",
    tags: "SMM, Контент-план, Reels, Telegram, Аналитика",
    rating: 4.85,
    reviews_count: 423,
    description: "Научитесь вести соцсети бизнеса: стратегия, контент, съёмка Reels, работа с блогерами и отчёты по результатам."
  },
  {
    id: "course-15",
    title: "SEO-специалист",
    category: "Маркетинг",
    level: "Для начинающих",
    duration: "5 месяцев",
    price_full: 11500000,
    price_discount: 8050000,
    price_monthly: 340000,
    discount_percent: 30,
    instructor_id: "inst-8",
    badge: "Востребовано",
    tags: "SEO, Семантика, Google Search Console, Аналитика, Контент",
    rating: 4.8,
    reviews_count: 167,
    description: "Поднимайте сайты в поисковой выдаче: семантическое ядро, техническая оптимизация, контент и ссылки."
  },
  {
    id: "course-16",
    title: "Таргетолог: реклама в Instagram и Telegram Ads",
    category: "Маркетинг",
    level: "С нуля",
    duration: "4 месяца",
    price_full: 9200000,
    price_discount: 5980000,
    price_monthly: 250000,
    discount_percent: 35,
    instructor_id: "inst-8",
    badge: "Практика",
    tags: "Таргетинг, Meta Ads, Telegram Ads, Креативы, Воронки",
    rating: 4.82,
    reviews_count: 289,
    description: "Запускайте рекламу, которая приводит клиентов: аудитории, креативы, тесты и расчёт окупаемости."
  },
  {
    id: "course-17",
    title: "Бизнес-аналитик",
    category: "Управление",
    level: "Продвинутый",
    duration: "6 месяцев",
    price_full: 13800000,
    price_discount: 9660000,
    price_monthly: 400000,
    discount_percent: 30,
    instructor_id: "inst-5",
    badge: "Для карьеры",
    tags: "Анализ требований, BPMN, SQL, UML, Документация",
    rating: 4.89,
    reviews_count: 138,
    description: "Научитесь собирать требования, моделировать процессы и ставить задачи разработчикам в IT-проектах."
  },
  {
    id: "course-18",
    title: "Scrum-мастер и Agile-коуч",
    category: "Управление",
    level: "Для начинающих",
    duration: "4 месяца",
    price_full: 9200000,
    price_discount: 5980000,
    price_monthly: 250000,
    discount_percent: 35,
    instructor_id: "inst-5",
    badge: "Короткий курс",
    tags: "Scrum, Agile, Kanban, Jira, Фасилитация",
    rating: 4.84,
    reviews_count: 109,
    description: "Освойте гибкие методологии и научитесь выстраивать работу команды: спринты, ретроспективы и метрики."
  },
  {
    id: "course-19",
    title: "Геймдизайнер",
    category: "Игры",
    level: "Для начинающих",
    duration: "7 месяцев",
    price_full: 16100000,
    price_discount: 10470000,
    price_monthly: 440000,
    discount_percent: 35,
    instructor_id: "inst-1",
    badge: "Для творческих",
    tags: "Геймдизайн, Баланс, Прототипирование, Unity, Документация",
    rating: 4.87,
    reviews_count: 156,
    description: "Придумывайте игровые механики, прописывайте баланс и уровни, собирайте прототипы и защищайте идею перед командой."
  },
  {
    id: "course-20",
    title: "Unreal Engine: разработчик игр",
    category: "Игры",
    level: "Продвинутый",
    duration: "10 месяцев",
    price_full: 23000000,
    price_discount: 16100000,
    price_monthly: 670000,
    discount_percent: 30,
    instructor_id: "inst-1",
    badge: "Продвинутый",
    tags: "Unreal Engine, C++, Blueprints, 3D, Мультиплеер",
    rating: 4.9,
    reviews_count: 87,
    description: "Создайте проект на Unreal Engine с нуля: игровая логика на Blueprints и C++, графика, оптимизация и сетевой режим."
  }
];

// Примеры преподавателей: вымышленные, для демонстрации. Замените на реальных.
const instructors = [
  {
    id: "inst-1",
    name: "Тимур Каримов",
    role: "Lead Fullstack & Game Dev",
    company: "Fintech-компания, Ташкент",
    experience: "11 лет в индустрии",
    avatar: "/avatars/inst-1.svg",
    bio: "Архитектор высоконагруженных веб-сервисов и создатель популярных open-source решений. Обучил свыше 2500 разработчиков."
  },
  {
    id: "inst-2",
    name: "Мадина Рахимова",
    role: "Principal Product Designer",
    company: "Продуктовая студия, Ташкент",
    experience: "8 лет в продуктовом дизайне",
    avatar: "/avatars/inst-2.svg",
    bio: "Спикер профильных дизайн-конференций, автор гайдлайнов по дизайн-системам. Специализируется на комплексных fintech и e-commerce интерфейсах."
  },
  {
    id: "inst-3",
    name: "Жасур Юсупов",
    role: "Head of AI & Data Science",
    company: "AI-лаборатория, Ташкент",
    experience: "10 лет в ML и Big Data",
    avatar: "/avatars/inst-3.svg",
    bio: "Руководит внедрением генеративных LLM и рекомендательных систем. Кандидат физико-математических наук."
  },
  {
    id: "inst-4",
    name: "Нодира Абдуллаева",
    role: "CMO & Growth Lead",
    company: "Маркетинговое агентство, Ташкент",
    experience: "9 лет в digital-маркетинге",
    avatar: "/avatars/inst-4.svg",
    bio: "Эксперт по перформанс-маркетингу и автоматизации сквозной аналитики. Вела рекламные кампании для крупных компаний Узбекистана."
  },
  {
    id: "inst-5",
    name: "Бобур Хасанов",
    role: "CPO / Senior Product Director",
    company: "Маркетплейс, Ташкент",
    experience: "12 лет в продуктовом менеджменте",
    avatar: "/avatars/inst-5.svg",
    bio: "Запустил более 15 успешных цифровых сервисов. Ментор стартапов и преподаватель программ для руководителей цифровой трансформации."
  },
  {
    id: "inst-6",
    name: "Рустам Ташматов",
    role: "Senior Backend & QA Lead",
    company: "IT-аутсорсинг, Ташкент",
    experience: "9 лет в разработке",
    avatar: "/avatars/inst-6.svg",
    bio: "Строит серверные системы и выстраивает тестирование в продуктовых командах. Наставник джуниоров, помог выйти на первую работу более чем 300 студентам."
  },
  {
    id: "inst-7",
    name: "Шахноза Исмоилова",
    role: "Art Director & Motion Designer",
    company: "Креативное агентство, Ташкент",
    experience: "10 лет в дизайне",
    avatar: "/avatars/inst-7.svg",
    bio: "Руководит визуальными проектами для брендов Узбекистана: айдентика, реклама, анимация и 3D. Автор курсов по графике и motion-дизайну."
  },
  {
    id: "inst-8",
    name: "Азиз Нурматов",
    role: "Head of SMM & Performance",
    company: "Digital-агентство, Ташкент",
    experience: "8 лет в маркетинге",
    avatar: "/avatars/inst-8.svg",
    bio: "Ведёт рекламу и соцсети для локального бизнеса: розница, образование, e-commerce. Практик в таргетированной рекламе, SMM и SEO."
  }
];

const wb = xlsx.utils.book_new();
const wsCourses = xlsx.utils.json_to_sheet(courses);
const wsInstructors = xlsx.utils.json_to_sheet(instructors);

xlsx.utils.book_append_sheet(wb, wsCourses, "Courses");
xlsx.utils.book_append_sheet(wb, wsInstructors, "Instructors");

const excelPath = path.join(storageDir, 'courses_data.xlsx');
xlsx.writeFile(wb, excelPath);
console.log(`Excel создан успешно: ${excelPath}`);

// 2. Создание Word (.docx) файлов базы знаний для RAG
async function createWordFiles() {
  // Файл 1: FAQ и поступление
  const docFaq = new Document({
    sections: [{
      properties: {},
      children: [
        new Paragraph({ text: "База знаний: Поступление, рассрочка и гарантия трудоустройства", heading: HeadingLevel.HEADING_1 }),
        new Paragraph({ text: "1. Как работает программа рассрочки и льготного обучения?" }),
        new Paragraph({ text: "Все курсы платформы можно оплачивать в беспроцентную рассрочку сроком до 24-36 месяцев. Первый платеж взимается только через 3 месяца после старта обучения (период льготных каникул)." }),
        new Paragraph({ text: "2. Что представляет собой гарантия трудоустройства?" }),
        new Paragraph({ text: "Если после успешного прохождения всех практических заданий и дипломной работы выпускник не находит работу при поддержке нашего Центра Карьеры в течение 6 месяцев — мы полностью возвращаем 100% стоимости обучения по договору." }),
        new Paragraph({ text: "3. Преподаватели и практика:" }),
        new Paragraph({ text: "Все преподаватели — действующие тимлиды и сеньоры из IT-компаний Ташкента. Проверка домашних работ осуществляется персональным ментором в течение 24 часов." }),
        new Paragraph({ text: "4. Как поступить на курс?" }),
        new Paragraph({ text: "Поступление проходит в четыре шага: оставьте заявку на сайте, пройдите бесплатную консультацию с менеджером (15 минут), подпишите договор и внесите первый платёж либо оформите рассрочку. Вступительных экзаменов нет, для курсов уровня «С нуля» знания не требуются." }),
        new Paragraph({ text: "5. Какие документы нужны для поступления?" }),
        new Paragraph({ text: "Для заключения договора нужен паспорт или ID-карта (для студентов младше 18 лет — документ одного из родителей или законного представителя). Для оформления рассрочки дополнительно указывается ПИНФЛ и номер телефона, привязанный к банковской карте. Справки об образовании и аттестат не требуются." }),
        new Paragraph({ text: "6. Какие требования к студентам и оборудованию?" }),
        new Paragraph({ text: "Минимальный возраст — 16 лет. Нужен компьютер или ноутбук с доступом в интернет и камерой для онлайн-занятий. Для курсов по дизайну и разработке игр рекомендуется устройство с 16 ГБ оперативной памяти." })
      ]
    }]
  });

  const bufferFaq = await Packer.toBuffer(docFaq);
  const faqPath = path.join(knowledgeDir, 'admission_and_career_faq.docx');
  fs.writeFileSync(faqPath, bufferFaq);
  console.log(`Word FAQ создан: ${faqPath}`);

  // Файл 2: Программа Fullstack и AI
  const docPrograms = new Document({
    sections: [{
      properties: {},
      children: [
        new Paragraph({ text: "Учебные программы: IT, Web и Искусственный Интеллект", heading: HeadingLevel.HEADING_1 }),
        new Paragraph({ text: "Модуль 1. Основы веб-технологий и алгоритмов: HTML5, CSS3, современный JavaScript (ES6+), архитектура браузера, Git и совместная разработка." }),
        new Paragraph({ text: "Модуль 2. Frontend-инженерия: React 19, TypeScript, управление состоянием (Zustand/Redux Toolkit), оптимизация рендеринга и доступность UI." }),
        new Paragraph({ text: "Модуль 3. Backend и Базы Данных: Серверы на Node.js и Python FastAPI, проектирование реляционных баз PostgreSQL, ORM, миграции, Redis кэширование." }),
        new Paragraph({ text: "Модуль 4. AI и Data Science интеграции: Архитектура Retrieval-Augmented Generation (RAG), интеграция с Gemini API, векторные эмбеддинги, разработка интеллектуальных ассистентов." }),
        new Paragraph({ text: "Модуль 5. DevOps и дипломный проект: Docker контейнеризация, CI/CD пайплайны, деплой на облачные серверы и защита реального проекта перед комиссией работодателей." })
      ]
    }]
  });

  const bufferProg = await Packer.toBuffer(docPrograms);
  const progPath = path.join(knowledgeDir, 'it_and_ai_programs.docx');
  fs.writeFileSync(progPath, bufferProg);
  console.log(`Word Программы создан: ${progPath}`);
}

createWordFiles().catch(err => {
  console.error("Ошибка при генерации Word файлов:", err);
});
