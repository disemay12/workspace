# Implementation Plan: Сайт скачивания WorkSpace

**Branch**: `001-download-site` | **Date**: 2026-11-11 | **Spec**: [spec.md](./spec.md)

**Input**: Feature specification from `/specs/001-download-site/spec.md`

## Summary

Добавить в репозиторий статический сайт продукта в каталоге `docs/` (публикуется
GitHub Pages), с реальными скриншотами интерфейса, снятыми с чистого профиля
приложения поверх синтетической демо-папки. Съёмка и генерация демо-данных —
скрипты в `scripts/site/`, чтобы кадры обновлялись после изменения UI. Дизайн — по
Hallmark v2.0.0 (токены, Floor, sloplint), без внешних шрифтов и CDN, только
относительные пути (подпуть `/workspace/` на GitHub Pages).

## Technical Context

**Language/Version**: HTML5 + CSS (custom properties), без сборки; вспомогательные
скрипты — Python 3.12 (вендорный интерпретатор DSH) и Node 24.

**Primary Dependencies**: ядро — ни одной; съёмка кадров — `patchright`/Playwright
(уже установлен в `~/.dsh/scrapling/venv`), генерация демо-документов —
`python-docx`, `openpyxl`, `python-pptx` (бандл DSH).

**Storage**: файлы `docs/assets/screens/*.png` в репозитории; демо-папка
генерируется в `%TEMP%\workspace-site-demo` и не коммитится.

**Testing**: `sloplint.mjs --render` (Hallmark), локальный статический сервер +
съёмка страницы на 390/768/1280/1920, проверка отсутствия личных данных в кадрах,
`npm run build`/`npm run typecheck` для приложения.

**Target Platform**: GitHub Pages, ветка `main`, каталог `/docs`; браузеры
Windows/Android/iOS последних версий; работает по `file://`.

**Project Type**: desktop-app + статический лендинг в том же репозитории.

**Performance Goals**: один запрос HTML + один CSS + кадры с `loading="lazy"`;
Lighthouse mobile Performance ≥ 90, Accessibility ≥ 95.

**Constraints**: без внешних CDN/шрифтов (офлайн-политика продукта), без сборки,
относительные пути, `.nojekyll`, размер страницы без кадров < 60 КБ.

**Scale/Scope**: 1 страница, 9 разделов, 6–8 кадров, RU.

## Constitution Check

*GATE: проверяется до Phase 0 и повторно после Phase 1.*

| Принцип | Как соблюдается |
|---|---|
| I. Продукт — источник правды | Тексты собраны из `README.md`; версия `1.0.15`, ассет `WorkSpace-1.0.15-win-x64.zip`, размер 264 177 605 Б — из GitHub API релиза |
| II. Никаких личных данных | Чистый `--user-data-dir`, синтетическая демо-папка, проверка каждого кадра |
| III. Статика без сборки | `docs/` = HTML + CSS + PNG, все пути относительные |
| IV. Hallmark | токены `:root`, Floor-правила, `sloplint --render` |
| V. Доступность | `focus-visible`, `prefers-reduced-motion`, контраст, alt, no-JS |
| VI. Воспроизводимость | `scripts/site/make-demo-data.py`, `scripts/site/shoot-app.py`, `scripts/site/build-screens.ps1` |

Нарушений нет.

## Project Structure

### Documentation (this feature)

```text
specs/001-download-site/
├── spec.md
├── plan.md
├── tasks.md
└── checklists/requirements.md
```

### Source Code (repository root)

```text
docs/                          # артефакт GitHub Pages, публикуется как есть
├── index.html
├── styles.css                 # @import tokens.css + стили страницы
├── tokens.css                 # токены темы Hallmark (scaffold)
├── .nojekyll
└── assets/
    ├── screens/               # кадры приложения (2x PNG, webp при необходимости)
    └── icon.png               # иконка приложения (из public/icon.png)

scripts/site/
├── make-demo-data.py          # синтетическая демо-папка (docx, xlsx, pptx, md, …)
├── shoot-app.py               # запуск Electron с CDP и съёмка экранов
└── README.md                  # как пересоздать кадры

.github/workflows/pages.yml    # публикация docs/ через actions/deploy-pages
```

**Structure Decision**: лендинг живёт в `docs/` того же репозитория — публикация
включается одним переключателем (Deploy from branch → `main` → `/docs`) или
готовым workflow; каталог не пересекается с кодом приложения.

## Порядок работ

1. **Демо-данные** — `make-demo-data.py` создаёт папку с документами без личных
   данных (названия нейтральные: «Квартальный отчёт», «Заметки по релизу»).
2. **Съёмка** — `shoot-app.py` поднимает `dist-electron/main.js` с
   `--user-data-dir=%TEMP%\ws-site-profile`, подложив `workspace-session.json` с
   `lastPath` = демо-папка; по CDP снимает: проводник + Word, Excel с формулами,
   Markdown в режиме «разделённый», код в Monaco, PowerPoint, AI-панель, настройки.
3. **Сайт** — тема Hallmark, `docs/index.html` + `docs/styles.css`; первый экран
   с основным действием, далее возможности с кадрами, требования, ограничения,
   FAQ, футер со ссылками на репозиторий и релизы.
4. **Проверка** — локальный сервер, съёмка 390/768/1280/1920, `sloplint --render`,
   просмотр кадров на личные данные.
5. **Публикация** — `.nojekyll`, workflow Pages, короткая инструкция в README.

## Complexity Tracking

Нарушений Constitution Check нет; дополнительных компонентов не требуется.
