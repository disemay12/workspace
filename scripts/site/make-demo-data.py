"""Синтетическая демо-папка для скриншотов сайта WorkSpace.

Все документы генерируются с нуля: ни личных файлов, ни переписки, ни ключей.
Запуск:
    python scripts/site/make-demo-data.py <путь-демо-папки>
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Inches, Pt as PPt

TITLE = "Квартальный отчёт"
SUBTITLE = "Отдел аналитики, четвёртый квартал"


def make_chart(path: Path) -> None:
    width, height = 960, 540
    image = Image.new("RGB", (width, height), "#0f141d")
    draw = ImageDraw.Draw(image)
    values = [64, 78, 71, 92, 88, 103]
    labels = ["июл", "авг", "сен", "окт", "ноя", "дек"]
    base = height - 90
    bar_w = 74
    gap = 54
    x = 120
    for value, label in zip(values, labels):
        top = base - value * 3
        draw.rectangle([x, top, x + bar_w, base], fill="#e2a33c")
        draw.text((x + bar_w // 2 - 14, base + 16), label, fill="#9aa4b2")
        draw.text((x + bar_w // 2 - 16, top - 26), str(value), fill="#e6eaf0")
        x += bar_w + gap
    draw.line([100, base, width - 100, base], fill="#33404f", width=2)
    image.save(path)


def make_docx(path: Path, picture: Path) -> None:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    head = doc.add_heading(TITLE, level=1)
    head.alignment = WD_ALIGN_PARAGRAPH.LEFT

    intro = doc.add_paragraph(
        "Сводка по загрузке команды и расходам на инфраструктуру за четвёртый квартал. "
        "Данные приведены по внутренним отчётам, без персональных сведений."
    )
    intro.paragraph_format.space_after = Pt(10)

    doc.add_heading("Что изменилось", level=2)
    for item in (
        "перевели сборку на один архив вместо набора папок;",
        "сократили время согласования макетов на два рабочих дня;",
        "добавили проверку кодировок при сохранении CSV.",
    ):
        doc.add_paragraph(item, style="List Bullet")

    doc.add_heading("Показатели", level=2)
    table = doc.add_table(rows=1, cols=3)
    table.style = "Light Grid Accent 1"
    header = table.rows[0].cells
    for cell, text in zip(header, ("Направление", "План", "Факт")):
        cell.text = text
    for row in (
        ("Документы", "120", "118"),
        ("Таблицы", "90", "94"),
        ("Презентации", "40", "37"),
    ):
        cells = table.add_row().cells
        for cell, text in zip(cells, row):
            cell.text = text

    doc.add_heading("Динамика обращений", level=2)
    doc.add_picture(str(picture), width=Inches(5.4))
    doc.add_paragraph("Источник: журнал обращений, ручная выборка по месяцам.").runs[0].italic = True

    doc.save(path)


def make_xlsx(path: Path) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "Свод"

    sheet["A1"] = "Бюджет проекта"
    sheet["A1"].font = Font(size=14, bold=True)
    sheet.merge_cells("A1:E1")

    heads = ["Статья", "План", "Факт", "Отклонение"]
    for col, text in enumerate(heads, start=1):
        cell = sheet.cell(row=3, column=col, value=text)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2F3B4A")
        cell.alignment = Alignment(horizontal="center")

    rows = [
        ("Инфраструктура", 180000, 172400),
        ("Лицензии", 96000, 96000),
        ("Подряд", 240000, 251300),
        ("Обучение", 60000, 38500),
        ("Резерв", 90000, 0),
    ]
    for offset, (name, plan, fact) in enumerate(rows):
        row = 4 + offset
        sheet.cell(row=row, column=1, value=name)
        sheet.cell(row=row, column=2, value=plan).number_format = "#,##0"
        sheet.cell(row=row, column=3, value=fact).number_format = "#,##0"
        sheet.cell(row=row, column=4, value=f"=C{row}-B{row}").number_format = "#,##0;[Red]-#,##0"

    total = 10
    sheet.cell(row=total, column=1, value="Итого").font = Font(bold=True)
    for col in (2, 3, 4):
        letter = "BCD"[col - 2]
        cell = sheet.cell(row=total, column=col, value=f"=SUM({letter}4:{letter}8)")
        cell.font = Font(bold=True)
        cell.number_format = "#,##0"

    for col, width in zip("ABCD", (22, 14, 14, 16)):
        sheet.column_dimensions[col].width = width
    sheet.freeze_panes = "A4"

    second = book.create_sheet("Заметки")
    second["A1"] = "Дата"
    second["B1"] = "Комментарий"
    second["A1"].font = Font(bold=True)
    second["B1"].font = Font(bold=True)
    for offset, text in enumerate(
        (
            "Перенос резерва на следующий период.",
            "Сверка по подрядчику закончена.",
            "Проверить акты в январе.",
        ),
        start=2,
    ):
        second.cell(row=offset, column=1, value=f"2026-12-0{offset}")
        second.cell(row=offset, column=2, value=text)
    second.column_dimensions["A"].width = 14
    second.column_dimensions["B"].width = 44

    book.save(path)


def make_pptx(path: Path, picture: Path) -> None:
    deck = Presentation()
    deck.slide_width = Inches(13.333)
    deck.slide_height = Inches(7.5)
    blank = deck.slide_layouts[6]

    def title(slide, text: str) -> None:
        box = slide.shapes.add_textbox(Inches(0.9), Inches(0.75), Inches(11.5), Inches(1.1))
        paragraph = box.text_frame.paragraphs[0]
        paragraph.text = text
        paragraph.font.size = PPt(40)
        paragraph.font.bold = True

    slide = deck.slides.add_slide(blank)
    title(slide, "Итоги квартала")
    box = slide.shapes.add_textbox(Inches(0.95), Inches(2.0), Inches(9.5), Inches(0.9))
    box.text_frame.paragraphs[0].text = "Отдел аналитики, четвёртый квартал"
    box.text_frame.paragraphs[0].font.size = PPt(20)

    slide = deck.slides.add_slide(blank)
    title(slide, "Что сделали")
    box = slide.shapes.add_textbox(Inches(0.95), Inches(2.1), Inches(11), Inches(3.2))
    frame = box.text_frame
    frame.text = "Собрали единый формат отчётов"
    for line in ("Убрали ручные переносы таблиц", "Добавили проверку формул", "Сократили путь согласования"):
        paragraph = frame.add_paragraph()
        paragraph.text = line
        paragraph.font.size = PPt(20)

    slide = deck.slides.add_slide(blank)
    title(slide, "Динамика обращений")
    slide.shapes.add_picture(str(picture), Inches(1.6), Inches(2.0), width=Inches(10.0))

    deck.save(path)


MARKDOWN = """# Регламент согласования

Документ описывает, как макет проходит путь от автора до публикации. Правки
вносятся в одном месте, история хранится рядом с файлом.

## Порядок

1. Автор заполняет карточку и прикладывает макет.
2. Редактор проверяет формулировки.
3. Ответственный подтверждает срок.

## Таблица ролей

| Роль | Что проверяет | Срок |
|---|---|---|
| Автор | полноту данных | 1 день |
| Редактор | формулировки | 2 дня |
| Ответственный | сроки и бюджет | 1 день |

## Служебная часть

```json
{
  "revision": 14,
  "status": "approved",
  "checks": ["spelling", "tables"]
}
```

- [x] вычитать введение
- [x] сверить таблицу ролей
- [ ] приложить финальный макет

> Замечания, которые не вошли в регламент, переносятся в следующий выпуск.
"""

CODE = '''"""Сводка по бюджету из журнала операций."""

from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path


@dataclass
class Row:
    article: str
    plan: int
    fact: int

    @property
    def deviation(self) -> float:
        if not self.plan:
            return 0.0
        return (self.fact - self.plan) / self.plan * 100


def load(path: Path) -> list[Row]:
    with path.open(encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle, delimiter=";")
        return [
            Row(item["article"], int(item["plan"]), int(item["fact"]))
            for item in reader
            if item.get("article")
        ]


def summary(rows: list[Row]) -> str:
    worst = max(rows, key=lambda row: abs(row.deviation), default=None)
    if worst is None:
        return "нет данных"
    return f"{worst.article}: {worst.deviation:+.1f}%"
'''

LOG = """2026-12-01 09:14:02 INFO  server  запуск, конфигурация прочитана
2026-12-01 09:14:02 INFO  server  слушаем 127.0.0.1:8080
2026-12-01 09:14:31 INFO  queue   задача 1041 принята, размер 2,4 МБ
2026-12-01 09:14:35 INFO  queue   задача 1041 завершена за 3,8 с
2026-12-01 09:21:07 WARN  cache   запись вытеснена, ключ отчёта 2026-11
2026-12-01 09:33:12 INFO  sync    файлов в очереди: 3
2026-12-01 09:33:19 INFO  sync    отправлено 3 из 3
2026-12-01 09:47:55 WARN  parser  строка 118: неожиданный разделитель, пропущена
2026-12-01 10:02:41 INFO  report  собран отчёт за ноябрь, строк 412
2026-12-01 10:11:03 ERROR export  не удалось записать файл: каталог только для чтения
2026-12-01 10:11:03 INFO  export  повтор после смены каталога
2026-12-01 10:11:04 INFO  export  отчёт сохранён, 412 строк
2026-12-01 10:26:18 INFO  queue   задача 1042 принята, размер 118 КБ
2026-12-01 10:26:19 INFO  queue   задача 1042 завершена за 0,9 с
2026-12-01 10:40:00 WARN  auth    истёк срок токена, запрошен новый
2026-12-01 10:40:01 INFO  auth    токен обновлён
"""

CONFIG = {
    "profile": "default",
    "workspace": {
        "autosave": True,
        "intervalSeconds": 30,
        "encoding": "utf-8",
        "newline": "crlf",
    },
    "excel": {"recalculateOnOpen": True, "showFormulas": False, "maxRows": 20000},
    "updates": {"channel": "stable", "checkOnStart": True},
}


def main() -> None:
    root = Path(sys.argv[1] if len(sys.argv) > 1 else "workspace-demo").resolve()
    root.mkdir(parents=True, exist_ok=True)

    picture = root / "Диаграмма.png"
    make_chart(picture)
    make_docx(root / f"{TITLE}.docx", picture)
    make_xlsx(root / "Бюджет 2026.xlsx")
    make_pptx(root / "Итоги квартала.pptx", picture)
    (root / "Регламент.md").write_text(MARKDOWN, encoding="utf-8")
    (root / "summary.py").write_text(CODE, encoding="utf-8")
    (root / "server.log").write_text(LOG, encoding="utf-8")
    (root / "config.json").write_text(json.dumps(CONFIG, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    archive = root / "Архив"
    archive.mkdir(exist_ok=True)
    (archive / "Итоги 2025.md").write_text(
        "# Итоги 2025\n\nГод закрыт, отчёт сдан, замечаний нет.\n", encoding="utf-8"
    )

    files = sorted(p.name for p in root.iterdir())
    print(root)
    for name in files:
        print("  " + name)


if __name__ == "__main__":
    main()
