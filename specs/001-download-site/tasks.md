# Tasks: Сайт скачивания WorkSpace

**Feature**: `001-download-site` · **Plan**: [plan.md](./plan.md) · **Spec**: [spec.md](./spec.md)

## Phase 1. Подготовка

- [x] T001 Инициализировать spec-kit в репозитории (`.specify/`), описать конституцию проекта
- [x] T002 Зафиксировать спецификацию и план (`specs/001-download-site/`)
- [x] T003 Прочитать факты о продукте и релизе: `README.md`, `package.json`, GitHub API `releases/latest`

## Phase 2. Демо-данные без личных файлов (FR-007)

- [x] T004 `scripts/site/make-demo-data.py`: генерация DOCX, XLSX, PPTX, Markdown, Python-файла, лога, конфига и PNG
- [x] T005 Нейтральные имена и содержимое: «Квартальный отчёт», «Бюджет 2026», «Регламент согласования»; без имён, почты и переписки
- [x] T006 Проверка: скрипт создаёт папку с нуля, ни один файл не берётся с диска пользователя

## Phase 3. Скриншоты (FR-006, FR-007, FR-008)

- [x] T007 Запуск приложения с чистым `--user-data-dir`, отдельным CDP-портом и демо-папкой в качестве рабочей
- [x] T008 `scripts/site/shoot-app.py`: эмуляция 1440 × 900 при deviceScaleFactor 2, тёмная тема, закрытый AI-док
- [x] T009 Съёмка семи экранов: Word, Excel, Markdown в разделённом режиме, код, презентации, лог, настройки провайдеров
- [x] T010 Страховка `assert_clean`: перед каждым кадром проверяются маркеры личных данных (`Dmitriy`, `Builders`, `secrets.local`, `Telegram`, `Notion`)
- [x] T011 Скрыть на кадре настроек плашку «найден файл с ключами» (артефакт машины автора: путь к репозиторию и имена посторонних секретов)
- [x] T012 Постобработка: 1600 px по ширине, WebP q88 (7 файлов, около 470 КБ суммарно)

## Phase 4. Сайт (FR-001—FR-005, FR-009—FR-012)

- [x] T013 Тема Hallmark Console (`theme.mjs deal` + `scaffold`), `docs/tokens.css` переведён с `@theme` на `:root`
- [x] T014 Шрифты: `scripts/site/fetch-fonts.mjs` скачивает Golos Text и JetBrains Mono (кириллица) в `docs/assets/fonts`, `docs/fonts.css` — без обращений к CDN
- [x] T015 `docs/index.html`: первый экран с кнопкой скачивания, bento из экранов, провайдеры, таблица форматов, обновления, требования, вопросы, финальный блок с измеренной строкой релиза, подвал
- [x] T016 `docs/styles.css`: штамп Hallmark, сетка 12 колонок, стек 60rem/40rem, таблица перестраивается в блоки на телефоне
- [x] T017 Только относительные пути, `.nojekyll`, `lang="ru"`, `title`, `description`, Open Graph, favicon, `preload` первого экрана
- [x] T018 Ссылки на скачивание: точный ассет `v1.0.15` и страница релизов; размер и версия сверены с GitHub API

## Phase 5. Проверка

- [x] T019 `sloplint.mjs --render` на 1280×800, 1440×900, 1920×1080 и 375 px: 0 FAIL, 0 REFLEX, 0 WARN
- [x] T020 Локальный просмотр всех брейкпоинтов, проверка скриншотов на личные данные
- [x] T021 Проверка работы из подпути GitHub Pages (относительные пути) и отсутствия 404 по ресурсам
- [x] T022 Pages включён (Settings → Pages → GitHub Actions), публикация прошла, страница открывается по `https://disemay12.github.io/workspace/`

## Обновление после релиза

1. `python scripts/site/make-demo-data.py %TEMP%\ws-site-demo` — при необходимости обновить демо-файлы.
2. Запустить приложение с CDP (команда в шапке `scripts/site/shoot-app.py`) и выполнить
   `python scripts/site/shoot-app.py --workspace <демо-папка> --out docs/assets/screens`.
3. Обновить версию, размер архива и ссылки в `docs/index.html` (три кнопки скачивания и строка релиза).
4. `node scripts/site/fetch-fonts.mjs` нужен только при смене шрифтов.
