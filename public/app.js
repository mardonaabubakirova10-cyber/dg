// Клиентская логика для Digital Courses (Стиль Bitwip с интерактивным кубиком)

let allCourses = [];
let allInstructors = [];
let currentCategory = 'Все';
let searchQuery = '';

document.addEventListener('DOMContentLoaded', () => {
  initBitwipApp();
  initCubeInteractions();
  initRagAssistant();
  initModals();
  startSyncPolling();
});

async function initBitwipApp() {
  await fetchCourses();
  await fetchInstructors();
  setupFilterPills();
  setupSearch();
}

/**
 * Получение данных о курсах из Excel (courses_data.xlsx)
 */
async function fetchCourses() {
  const grid = document.getElementById('coursesGrid');
  try {
    const res = await fetch('/api/courses');
    const data = await res.json();
    if (data.success) {
      allCourses = data.data;
      renderCourses();
    } else {
      grid.innerHTML = `<div style="grid-column: 1/-1; text-align: center; padding: 32px;">Ошибка: ${data.error}</div>`;
    }
  } catch (err) {
    grid.innerHTML = `<div style="grid-column: 1/-1; text-align: center; padding: 32px;">Не удалось загрузить курсы. Попробуйте обновить страницу.</div>`;
  }
}

/**
 * Получение преподавателей
 */
async function fetchInstructors() {
  const grid = document.getElementById('instructorsGrid');
  try {
    const res = await fetch('/api/instructors');
    const data = await res.json();
    if (data.success) {
      allInstructors = data.data;
      renderInstructors();
    }
  } catch (err) {
    console.error('Ошибка загрузки преподавателей:', err);
  }
}

/**
 * Рендеринг карточек курсов в премиальном тёмном стиле
 */
function renderCourses() {
  const grid = document.getElementById('coursesGrid');

  const filtered = allCourses.filter(course => {
    const matchCat = currentCategory === 'Все' || course.category.toLowerCase() === currentCategory.toLowerCase();
    const matchSearch = !searchQuery || 
      course.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
      (course.tags && course.tags.toLowerCase().includes(searchQuery.toLowerCase())) ||
      (course.description && course.description.toLowerCase().includes(searchQuery.toLowerCase()));
    return matchCat && matchSearch;
  });

  if (filtered.length === 0) {
    grid.innerHTML = `
      <div style="grid-column: 1/-1; text-align: center; padding: 60px 20px; background: #18191e; border-radius: 20px; border: 1px solid #272a33;">
        <h3 style="font-size: 20px; margin-bottom: 8px;">Курсов по выбранному фильтру не найдено</h3>
        <p style="color: #9aa1b0; font-size: 14px;">Попробуйте нажать другую плитку на кубике или сбросить фильтр</p>
      </div>
    `;
    return;
  }

  grid.innerHTML = filtered.map(course => {
    const tagsArr = course.tags ? course.tags.split(',').map(t => t.trim()) : [];
    const monthlyFormatted = Number(course.price_monthly || 0).toLocaleString('ru-RU');
    const discountFormatted = Number(course.price_discount || 0).toLocaleString('ru-RU');
    const fullFormatted = Number(course.price_full || 0).toLocaleString('ru-RU');

    return `
      <div class="course-card-bitwip" data-id="${course.id}">
        <div class="c-meta-top">
          <span class="c-badge">${course.badge || 'Профессия'}</span>
          <span class="c-duration">${course.duration}</span>
        </div>

        <h3 class="c-title">${course.title}</h3>
        <p class="c-desc">${course.description}</p>

        <div class="c-tags">
          ${tagsArr.map(t => `<span class="c-tag">${t}</span>`).join('')}
        </div>

        ${course.instructor ? `
          <div class="c-instructor">
            <img src="${course.instructor.avatar}" alt="${course.instructor.name}" class="c-inst-img" />
            <div>
              <div class="c-inst-name">${course.instructor.name}</div>
              <div class="c-inst-comp">${course.instructor.company}</div>
            </div>
          </div>
        ` : ''}

        <div class="c-price-box">
          <div>
            <div class="c-price-month">от ${monthlyFormatted} сум/мес</div>
            <div class="c-month-sub">Рассрочка без %</div>
          </div>
          <div class="c-full-wrap">
            <div class="c-price-old">${fullFormatted} сум</div>
            <div class="c-price-discount">${discountFormatted} сум (-${course.discount_percent}%)</div>
          </div>
        </div>

        <div class="c-btn-row">
          <button class="btn-card-buy" onclick="openEnroll('${course.id}')">Записаться</button>
          <button class="btn-card-ai" onclick="askAiCourse('${course.title.replace(/'/g, "\\'")}')">О курсе ИИ</button>
        </div>
      </div>
    `;
  }).join('');
}

/**
 * Рендеринг карточек преподавателей
 */
function renderInstructors() {
  const grid = document.getElementById('instructorsGrid');
  grid.innerHTML = allInstructors.map(inst => `
    <div class="inst-card-dark">
      <img src="${inst.avatar}" alt="${inst.name}" class="inst-img-big" />
      <h3 class="inst-title">${inst.name}</h3>
      <div class="inst-work">${inst.company}</div>
      <p class="inst-bio">${inst.bio}</p>
    </div>
  `).join('');
}

/**
 * Интерактивный кубик курсов (клик по плиткам фильтрует каталог и скроллит к нему)
 */
function initCubeInteractions() {
  const tiles = document.querySelectorAll('.cube-grid-3x3 .tile');
  tiles.forEach(tile => {
    tile.addEventListener('click', (e) => {
      e.stopPropagation();
      const cat = tile.getAttribute('data-cat');
      if (cat) {
        selectCategory(cat);
        // Скроллим к каталогу
        const catalogEl = document.getElementById('catalog');
        if (catalogEl) {
          catalogEl.scrollIntoView({ behavior: 'smooth' });
        }
      }
    });
  });
}

/**
 * Выбор категории (синхронизирует табы и кубик)
 */
function selectCategory(category) {
  currentCategory = category;
  const pills = document.querySelectorAll('.cat-pill');
  pills.forEach(p => {
    if (p.getAttribute('data-cat') === category) {
      p.classList.add('active');
    } else {
      p.classList.remove('active');
    }
  });
  renderCourses();
}

function setupFilterPills() {
  const pills = document.querySelectorAll('.cat-pill');
  pills.forEach(pill => {
    pill.addEventListener('click', () => {
      const cat = pill.getAttribute('data-cat');
      selectCategory(cat);
    });
  });
}

function setupSearch() {
  const input = document.getElementById('courseSearchInput');
  input.addEventListener('input', (e) => {
    searchQuery = e.target.value.trim();
    renderCourses();
  });
}

/**
 * Автоматический опрос обновлений Excel таблицы
 */
function startSyncPolling() {
  setInterval(async () => {
    try {
      const res = await fetch('/api/courses');
      const data = await res.json();
      if (data.success && JSON.stringify(data.data) !== JSON.stringify(allCourses)) {
        allCourses = data.data;
        renderCourses();
      }
    } catch (e) {}
  }, 4000);
}

/**
 * Логика RAG Ассистента
 */
function initRagAssistant() {
  const trigger = document.getElementById('ragTrigger');
  const popup = document.getElementById('ragPopup');
  const close = document.getElementById('ragClose');
  const form = document.getElementById('ragForm');
  const input = document.getElementById('ragUserInput');
  const dialog = document.getElementById('ragDialog');
  const loading = document.getElementById('ragLoading');
  const heroAiBtn = document.getElementById('heroAiTrigger');
  const openChatBtn = document.getElementById('openAiChatBtn');

  const toggle = () => {
    popup.classList.toggle('active');
    if (popup.classList.contains('active')) input.focus();
  };

  trigger.addEventListener('click', toggle);
  close.addEventListener('click', toggle);
  if (heroAiBtn) heroAiBtn.addEventListener('click', () => { popup.classList.add('active'); input.focus(); });
  if (openChatBtn) openChatBtn.addEventListener('click', () => { popup.classList.add('active'); input.focus(); });

  document.querySelectorAll('.sug-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      input.value = btn.getAttribute('data-q');
      form.dispatchEvent(new Event('submit'));
    });
  });

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const q = input.value.trim();
    if (!q) return;

    addChatMsg(q, 'user');
    input.value = '';
    loading.style.display = 'flex';
    dialog.scrollTop = dialog.scrollHeight;

    try {
      const res = await fetch(`/api/rag/query?q=${encodeURIComponent(q)}`);
      const data = await res.json();
      loading.style.display = 'none';

      if (data.success) {
        addChatMsg(data.answer, 'ai', data.sources, data);
      } else {
        addChatMsg('Произошла ошибка при обработке запроса.', 'ai');
      }
    } catch (err) {
      loading.style.display = 'none';
      addChatMsg('Не удалось получить ответ. Попробуйте ещё раз.', 'ai');
    }
  });

  const esc = (s) => String(s).replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));

  const SOURCE_NAMES = {
    'admission_and_career_faq.docx': 'Поступление, рассрочка и гарантия',
    'it_and_ai_programs.docx': 'Учебные программы',
    'courses_data.xlsx': 'Каталог курсов'
  };
  const sourceName = (file) => SOURCE_NAMES[file] || file;

  function addChatMsg(text, sender, sources = [], extra = {}) {
    const el = document.createElement('div');
    el.className = `chat-message ${sender}-message`;

    let html = '';
    if (sender === 'ai') {
      const lines = esc(text).split('\n');
      html += '<div class="step-format">';
      lines.forEach((line, idx) => {
        if (idx === 0) {
          html += `<div class="step-title">${line}</div>`;
        } else if (line.match(/^\d+\.\d+/)) {
          const m = line.match(/^(\d+\.\d+)(.*)/);
          html += `<div class="step-line"><strong>${m[1]}</strong>${m[2]}</div>`;
        } else {
          html += `<div class="step-line">${line}</div>`;
        }
      });
      if (extra.courses && extra.courses.length > 0) {
        html += '<div style="margin-top: 8px; display: flex; flex-wrap: wrap; gap: 6px;">';
        extra.courses.forEach(c => {
          html += `<button type="button" class="rag-course-link" data-title="${esc(c.title)}" style="font-size: 11px; padding: 4px 8px; border-radius: 999px; border: 1px solid #e8607e; background: transparent; color: #e8607e; cursor: pointer;">→ ${esc(c.title)}</button>`;
        });
        html += '</div>';
      }
      if (extra.citations && extra.citations.length > 0) {
        html += '<div style="font-size: 11px; opacity: 0.7; margin-top: 6px;">';
        extra.citations.forEach(c => {
          html += `<div style="margin-top: 2px;">[${c.n}] «${esc(c.text)}…» — ${esc(sourceName(c.source))}</div>`;
        });
        html += '</div>';
      }
      if (sources && sources.length > 0) {
        html += `<div style="font-size: 11px; color: #e8607e; margin-top: 6px; border-top: 1px dashed rgba(255,255,255,0.1); padding-top: 4px;">
          📄 Источники: ${sources.map(sourceName).map(esc).join(', ')}
        </div>`;
      }
      html += '</div>';
    } else {
      html = text;
    }

    el.innerHTML = `<div class="chat-bubble">${sender === 'ai' ? html : esc(html)}</div>`;
    el.querySelectorAll('.rag-course-link').forEach(btn => btn.addEventListener('click', () => {
      const search = document.getElementById('courseSearchInput');
      search.value = btn.dataset.title;
      search.dispatchEvent(new Event('input'));
      popup.classList.remove('active');
      document.getElementById('catalog').scrollIntoView({ behavior: 'smooth' });
    }));
    dialog.appendChild(el);
    dialog.scrollTop = dialog.scrollHeight;
  }

  window.askAiCourse = function(title) {
    popup.classList.add('active');
    input.value = `Расскажи подробно про курс «${title}»`;
    form.dispatchEvent(new Event('submit'));
  };
}

/**
 * Модальные окна
 */
function initModals() {
  const modal = document.getElementById('enrollModal');
  const dismiss = document.getElementById('modalDismiss');
  const body = document.getElementById('enrollModalBody');
  const consultBtn = document.getElementById('consultBtn');

  dismiss.addEventListener('click', () => modal.classList.remove('open'));
  modal.addEventListener('click', (e) => {
    if (e.target === modal) modal.classList.remove('open');
  });

  if (consultBtn) {
    consultBtn.addEventListener('click', () => {
      openConsultModal();
    });
  }

  window.openEnroll = function(courseId) {
    const c = allCourses.find(item => String(item.id) === String(courseId));
    if (!c) return;

    body.innerHTML = `
      <h3 style="font-size: 22px; font-weight: 800; margin-bottom: 8px;">Запись на курс</h3>
      <p style="color: #9aa1b0; font-size: 14px; margin-bottom: 20px;">${c.title}</p>
      
      <div style="background: rgba(176, 38, 76, 0.1); border: 1px solid rgba(176, 38, 76, 0.2); padding: 14px; border-radius: 12px; margin-bottom: 20px;">
        <div style="font-size: 12px; color: #e8607e; font-weight: 700; text-transform: uppercase;">Рассрочка без %:</div>
        <div style="font-size: 20px; font-weight: 900; color: #fff;">от ${Number(c.price_monthly).toLocaleString('ru-RU')} сум/мес</div>
        <div style="font-size: 12px; color: #9aa1b0; margin-top: 4px;">Первые 3 месяца каникул. Гарантия трудоустройства по договору.</div>
      </div>

      <form onsubmit="handleFormSubmit(event, '${c.id}')">
        <div style="margin-bottom: 14px;">
          <label style="display: block; font-size: 13px; font-weight: 600; margin-bottom: 6px;">Ваше имя</label>
          <input type="text" required placeholder="Константин" style="width: 100%; padding: 12px 14px; background: #242730; border: 1px solid #272a33; border-radius: 8px; color: #fff; font-size: 14px;" />
        </div>
        <div style="margin-bottom: 20px;">
          <label style="display: block; font-size: 13px; font-weight: 600; margin-bottom: 6px;">Номер телефона</label>
          <input type="tel" required placeholder="+998 90 123 45 67" style="width: 100%; padding: 12px 14px; background: #242730; border: 1px solid #272a33; border-radius: 8px; color: #fff; font-size: 14px;" />
        </div>
        <button type="submit" class="btn-hero-white" style="width: 100%; border: none; cursor: pointer;">Зафиксировать скидку -${c.discount_percent}%</button>
      </form>
    `;
    modal.classList.add('open');
  };

  function openConsultModal() {
    body.innerHTML = `
      <h3 style="font-size: 22px; font-weight: 800; margin-bottom: 8px;">Консультация с ментором</h3>
      <p style="color: #9aa1b0; font-size: 14px; margin-bottom: 20px;">Поможем определиться со специальностью и рассчитать условия рассрочки</p>
      
      <form onsubmit="handleFormSubmit(event)">
        <div style="margin-bottom: 14px;">
          <label style="display: block; font-size: 13px; font-weight: 600; margin-bottom: 6px;">Имя</label>
          <input type="text" required placeholder="Ваше имя" style="width: 100%; padding: 12px 14px; background: #242730; border: 1px solid #272a33; border-radius: 8px; color: #fff; font-size: 14px;" />
        </div>
        <div style="margin-bottom: 20px;">
          <label style="display: block; font-size: 13px; font-weight: 600; margin-bottom: 6px;">Телефон</label>
          <input type="tel" required placeholder="+998 90 123 45 67" style="width: 100%; padding: 12px 14px; background: #242730; border: 1px solid #272a33; border-radius: 8px; color: #fff; font-size: 14px;" />
        </div>
        <button type="submit" class="btn-hero-white" style="width: 100%; border: none; cursor: pointer;">Получить консультацию</button>
      </form>
    `;
    modal.classList.add('open');
  }

  window.handleFormSubmit = async function(e, courseId) {
    e.preventDefault();
    const form = e.target;
    const [nameInput, phoneInput] = form.querySelectorAll('input');
    const button = form.querySelector('button[type="submit"]');
    const label = button.textContent;
    button.disabled = true;
    button.textContent = 'Отправляем…';
    let error = '';
    try {
      const res = await fetch('/api/leads', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: nameInput.value, phone: phoneInput.value, course_id: courseId })
      });
      const data = await res.json();
      if (!data.success) error = data.error || 'Не удалось отправить заявку.';
    } catch (err) {
      error = 'Нет связи с сервером. Попробуйте позже.';
    }
    if (error) {
      button.disabled = false;
      button.textContent = label;
      let box = form.querySelector('.form-error');
      if (!box) {
        box = document.createElement('div');
        box.className = 'form-error';
        box.style.cssText = 'color: #e8607e; font-size: 13px; margin-top: 12px;';
        form.appendChild(box);
      }
      box.textContent = error;
      return;
    }
    body.innerHTML = `
      <div style="text-align: center; padding: 24px 0;">
        <div style="font-size: 44px; margin-bottom: 12px;">🎉</div>
        <h3 style="font-size: 20px; font-weight: 800; margin-bottom: 8px;">Заявка принята</h3>
        <p style="color: #9aa1b0; font-size: 14px;">Консультант свяжется с вами в течение 15 минут.</p>
        <button class="btn-hero-white" style="margin-top: 20px; border: none; cursor: pointer;" onclick="document.getElementById('enrollModal').classList.remove('open')">Понятно</button>
      </div>
    `;
  };
}
