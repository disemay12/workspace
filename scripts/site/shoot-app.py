"""Съёмка экранов WorkSpace для сайта.

Приложение уже должно быть запущено с DevTools-портом:

    node scripts/run-electron.js dist-electron/main.js \
        --remote-debugging-port=9333 \
        --user-data-dir=%LOCALAPPDATA%\\Temp\\ws-site-profile

Затем:

    python scripts/site/shoot-app.py --workspace C:/Demo --out docs/assets/screens

Кадры снимаются с эмуляцией 1440 x 900 при deviceScaleFactor 2, уменьшаются до
1600 px по ширине и сохраняются в WebP.
"""

from __future__ import annotations

import argparse
import base64
import io
import json
import os
import sys
from pathlib import Path

from patchright.sync_api import sync_playwright
from PIL import Image


class App:
    def __init__(self, session, page):
        self.s = session
        self.page = page

    def ev(self, expression: str):
        result = self.s.send(
            "Runtime.evaluate",
            {"expression": expression, "returnByValue": True, "awaitPromise": True},
        )
        if "exceptionDetails" in result:
            raise RuntimeError(str(result["exceptionDetails"])[:300])
        return result.get("result", {}).get("value")

    def wait(self, ms: int = 1200) -> None:
        self.page.wait_for_timeout(ms)

    def reload(self, settle: int = 3000) -> None:
        self.page.reload()
        self.wait(settle)

    def find(self, predicate: str):
        """Центр первого элемента, подходящего под JS-предикат (return element)."""
        return self.ev(
            """(() => {
              const hit = (() => { %s })();
              if (!hit) return null;
              const b = hit.getBoundingClientRect();
              if (b.width < 2 || b.height < 2) return null;
              return [b.x + b.width / 2, b.y + b.height / 2];
            })()"""
            % predicate
        )

    def click(self, pos, clicks: int = 1) -> bool:
        if not pos:
            return False
        x, y = pos
        for count in range(1, clicks + 1):
            self.s.send("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": x, "y": y})
            self.s.send("Input.dispatchMouseEvent", {"type": "mousePressed", "x": x, "y": y, "button": "left", "clickCount": count})
            self.s.send("Input.dispatchMouseEvent", {"type": "mouseReleased", "x": x, "y": y, "button": "left", "clickCount": count})
            self.wait(40)
        return True

    def click_text(self, text: str, clicks: int = 1) -> bool:
        pos = self.find(
            """const t = %s;
            return Array.from(document.querySelectorAll('div,span,button,li'))
              .filter(e => e.children.length === 0)
              .find(e => (e.textContent || '').trim() === t);"""
            % repr(text)
        )
        return self.click(pos, clicks)

    def click_label(self, label: str) -> bool:
        pos = self.find(
            """const l = %s.toLowerCase();
            return Array.from(document.querySelectorAll('[aria-label],[title]'))
              .find(e => ((e.getAttribute('aria-label') || '') + ' ' + (e.getAttribute('title') || ''))
                .toLowerCase().includes(l));"""
            % repr(label)
        )
        return self.click(pos)

    def has_text(self, text: str) -> bool:
        return bool(self.ev("(() => document.body.innerText.includes(%s))()" % repr(text)))

    def visible_text(self) -> str:
        return self.ev(
            """(() => {
              const vh = innerHeight, vw = innerWidth, out = [];
              for (const el of document.querySelectorAll('body *')) {
                if (el.children.length > 0) continue;
                const r = el.getBoundingClientRect();
                if (r.bottom < 0 || r.top > vh || r.right < 0 || r.left > vw || r.height < 2 || r.width < 2) continue;
                out.push(el.textContent || '');
                if (el.value) out.push(String(el.value));
              }
              return out.join('\\n');
            })()"""
        )

    def assert_clean(self) -> None:
        """Страховка принципа «никаких личных данных на кадрах»."""
        text = self.visible_text()
        leaked = [marker for marker in PERSONAL_MARKERS if marker.lower() in text.lower()]
        if leaked:
            raise SystemExit(f"на экране личные данные: {leaked}")

    def shot(self, path: Path) -> None:
        self.assert_clean()
        data = self.s.send("Page.captureScreenshot", {"format": "png"})["data"]
        image = Image.open(io.BytesIO(base64.b64decode(data))).convert("RGB")
        width = 1600
        height = round(image.height * width / image.width)
        image = image.resize((width, height), Image.LANCZOS)
        path.parent.mkdir(parents=True, exist_ok=True)
        image.save(path, "WEBP", quality=88, method=6)
        print(f"  {path.name}  {width}x{height}  {path.stat().st_size // 1024} КБ")


FILES = (
    ("word", "Квартальный отчёт.docx"),
    ("excel", "Бюджет 2026.xlsx"),
    ("markdown", "Регламент.md"),
    ("code", "summary.py"),
    ("slides", "Итоги квартала.pptx"),
    ("log", "server.log"),
)

# Маркеры личных данных: путь пользователя, имена файлов секретов, шлюз, мессенджеры.
PERSONAL_MARKERS = ("Dmitriy", "Builders", "secrets.local", ".local.ps1", "Telegram", "Notion")


def capture(port: int, workspace: str, out: Path) -> None:
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{port}")
        context = browser.contexts[0]
        page = next((pg for pg in context.pages if "WorkSpace" in (pg.title() or "")), context.pages[0])
        app = App(context.new_cdp_session(page), page)

        app.s.send(
            "Emulation.setDeviceMetricsOverride",
            {"width": 1440, "height": 900, "deviceScaleFactor": 2, "mobile": False},
        )
        app.wait(600)

        current = app.ev("(async () => await window.electronAPI.getCurrentWorkspace())()")
        print("workspace:", current)
        if os.path.normcase(str(current)) != os.path.normcase(workspace):
            app.ev("(async () => await window.electronAPI.setWorkspaceRoot(%s))()" % json.dumps(workspace))
            app.reload(4000)
            print("workspace ->", app.ev("(async () => await window.electronAPI.getCurrentWorkspace())()"))

        app.ev("(() => localStorage.setItem('workspace_theme', 'dark'))()")
        app.reload(3500)

        # без правого дока редакторы шире; сам док нужен только на кадре про AI
        if app.has_text("Чем помочь?"):
            app.click_label("AI-ассистент")
            app.wait(800)
            print("ai panel:", "closed" if not app.has_text("Чем помочь?") else "STILL OPEN")

        for name, target in FILES:
            print(name, "->", target)
            if not app.click_text(target, clicks=2):
                raise SystemExit(f"нет файла в дереве: {target}")
            app.wait(4200)
            if name == "markdown":
                if not app.click_label("разделённый режим"):
                    raise SystemExit("не найдена кнопка режима «разделённый режим»")
                app.wait(1800)
            app.shot(out / f"{name}.webp")

        if app.click_label("Настройки"):
            app.wait(2600)
            # Плашка «найден файл с ключами» — артефакт этой машины: путь к репозиторию и
            # имена посторонних секретов. На кадре её быть не должно.
            print("  settings banner:", app.ev(
                """(() => {
                  const deepest = (re) => {
                    const hits = Array.from(document.querySelectorAll('div,p,span'))
                      .filter(e => re.test(e.textContent || ''));
                    return hits.filter(e => !hits.some(o => o !== e && e.contains(o)))[0] || null;
                  };
                  const head = deepest(/файл с ключами/i);
                  const warn = deepest(/другие ключи/i);
                  if (!head && !warn) return 'плашки нет';
                  if (!head || !warn) {
                    (head || warn).parentElement.style.display = 'none';
                    return 'скрыта часть';
                  }
                  let node = head;
                  while (node && !node.contains(warn)) node = node.parentElement;
                  if (!node) return 'нет общего предка';
                  node.style.display = 'none';
                  return 'скрыта плашка';
                })()"""))
            app.wait(500)
            app.shot(out / "settings.webp")
        else:
            print("  ! кнопка настроек не найдена")

        browser.close()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=9333)
    parser.add_argument("--workspace", default="C:/Demo")
    parser.add_argument("--out", default="docs/assets/screens")
    args = parser.parse_args()
    capture(args.port, args.workspace, Path(args.out))


if __name__ == "__main__":
    sys.exit(main())
