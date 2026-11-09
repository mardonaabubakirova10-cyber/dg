# Digital Курсы: каталог с RAG-ассистентом

Сайт каталога курсов, у которого вместо базы данных Excel-таблица, а ассистент отвечает по Word-документам через Gemini.
Данные лежат в папке Google Drive и обновляются без перезапуска сервера.

## Запуск

```bash
npm install
cp .env.example .env     # затем заполните .env, см. ниже
npm start                # http://localhost:3000
```

## Что где лежит

| Файл | Назначение |
|---|---|
| `server.js` | Express-сервер и API (`/api/courses`, `/api/instructors`, `/api/meta`, `/api/rag/query`) |
| `dataService.js` | Чтение `courses_data.xlsx` и `.docx`, кеш по времени изменения файла |
| `ragService.js` | Нарезка текста на фрагменты, поиск, запрос к Gemini |
| `public/` | Фронтенд: `index.html`, `style.css`, `app.js` |
| `storage_drive/` | Локальные данные по умолчанию (копия структуры папки Drive) |
| `init_storage.js` | Пересоздаёт тестовые Excel и Word в `storage_drive/` (`npm run init-storage`) |

## Подключение Google Drive

Используется синхронизированная папка (Google Drive для компьютера), а не Drive API.

1. Установите [Google Drive для компьютера](https://www.google.com/drive/download/) и войдите в аккаунт.
2. В «Мой диск» создайте папку (например, `digital`) со структурой:
   ```
   digital/
   ├── courses_data.xlsx          # листы Courses и Instructors
   └── knowledge_base_word/       # любые .docx
   ```
3. В `.env` укажите путь к ней:
   ```
   DRIVE_FOLDER=G:\Мой диск\digital
   ```
4. Перезапустите сервер один раз. Дальше правки в Excel и Word подхватываются сами,
   как только Drive синхронизирует файл на компьютер.

Если `DRIVE_FOLDER` пуст, используется `storage_drive/`.

## Подключение Gemini

1. Создайте ключ на https://aistudio.google.com/apikey.
2. Впишите его в `.env`:
   ```
   GEMINI_API_KEY=ваш_ключ
   GEMINI_MODEL=gemini-3.5-flash-lite
   ```
3. Перезапустите сервер.

Файл `.env` содержит секрет и в `.gitignore`. Не отправляйте его и не публикуйте.

Без ключа, при ошибке или таймауте Gemini (20 секунд) ассистент не падает: он показывает найденные
фрагменты документов в том же формате шагов, но без пересказа.

## Формат таблицы `courses_data.xlsx`

Лист `Courses`: `id, title, category, level, duration, price_full, price_discount, price_monthly,
discount_percent, instructor_id, badge, tags, rating, reviews_count, description`.
Лист `Instructors`: `id, name, role, company, experience, avatar, bio`.
`instructor_id` курса должен совпадать с `id` преподавателя.

## Как работает ассистент

1. Word-документы режутся на фрагменты (заголовок или вопрос склеивается с абзацем после него),
   каждый курс из Excel становится отдельным фрагментом.
2. Для вопроса ищутся до 4 лучших фрагментов по совпадению слов (без эмбеддингов).
3. Фрагменты передаются Gemini с требованием отвечать только по ним и ставить номера источников.
4. В чате показываются ответ, цитаты `[1]`, `[2]`, файлы-источники и кнопки найденных курсов.

Ограничение: поиск по словам не понимает синонимов («стоимость» и «цена» для него разные слова).
Если ответы начнут промахиваться на таких вопросах, следующий шаг — эмбеддинги Gemini.

## Проверка обновления данных

Измените цену в `courses_data.xlsx` и сохраните: на сайте она изменится через несколько секунд
(страница раз в несколько секунд опрашивает `/api/courses`).
