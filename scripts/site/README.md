# Скриншоты и шрифты сайта

Всё, что попадает в `docs/`, пересоздаётся этими скриптами. Личных файлов на кадрах
нет: съёмка идёт на чистом профиле приложения поверх синтетической демо-папки,
а `shoot-app.py` перед каждым кадром проверяет маркеры личных данных и падает,
если они попали в кадр.

## 1. Демо-папка

```powershell
python scripts/site/make-demo-data.py C:\Demo
```

Создаёт `Квартальный отчёт.docx`, `Бюджет 2026.xlsx`, `Итоги квартала.pptx`,
`Регламент.md`, `summary.py`, `server.log`, `config.json`, `Диаграмма.png` и папку
`Архив`. Все данные синтетические, имена нейтральные.

## 2. Запуск приложения с DevTools-портом

```powershell
Remove-Item Env:\ELECTRON_RUN_AS_NODE -ErrorAction SilentlyContinue
node scripts/run-electron.js dist-electron/main.js `
  --remote-debugging-port=9333 `
  --user-data-dir="$env:LOCALAPPDATA\Temp\ws-site-profile" `
  --force-device-scale-factor=2
```

`%LOCALAPPDATA%\Temp\ws-site-profile\workspace-session.json` заранее содержит
`{"lastPath": "<демо-папка>"}`, поэтому приложение открывает демо-папку, а не
`Documents`.

Если процесс запущен от администратора (UAC выключен), Chromium уходит в режим
низкой целостности и не может писать в профиль. Лечится меткой:

```powershell
icacls <демо-папка> /setintegritylevel "(OI)(CI)L"
icacls <профиль>   /setintegritylevel "(OI)(CI)L"
```

## 3. Съёмка

```powershell
python scripts/site/shoot-app.py --workspace C:/Demo --out docs/assets/screens
```

Скрипт подключается по CDP, ставит эмуляцию 1440 × 900 при `deviceScaleFactor 2`,
включает тёмную тему, закрывает правый AI-док, открывает файлы двойным кликом и
сохраняет семь кадров (Word, Excel, Markdown в разделённом режиме, код,
презентации, лог, настройки провайдеров) в WebP 1600 px по ширине.

На кадре настроек скрипт скрывает плашку «найден файл с ключами»: она показывает
путь к репозиторию автора и имена посторонних секретов, то есть является
артефактом машины, а не интерфейсом продукта.

## 4. Шрифты

```powershell
node scripts/site/fetch-fonts.mjs
```

Скачивает Golos Text и JetBrains Mono (кириллица + латиница) в
`docs/assets/fonts/` и пишет `docs/fonts.css`. Тема Hallmark Console называет
Host Grotesk и Atkinson Hyperlegible Mono, но у них нет кириллицы, поэтому для
русской страницы взяты гротеск и моноширинный того же регистра. Страница не
обращается к CDN: шрифты лежат в репозитории.

## 5. Проверка

```powershell
python -m http.server 8137 --directory docs
node "$env:USERPROFILE\.dsh\skills\hallmark\scripts\sloplint.mjs" docs\index.html docs\styles.css --genre modern-minimal --url "http://127.0.0.1:8137/"
```

Ожидаемый итог: `0 FAIL, 0 REFLEX, 0 WARN`. Кадры проверки складываются в
`docs/.hallmark/shots` (каталог в `.gitignore`).
