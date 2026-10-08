"""Сборка «Рапорта по персоналу» из табелей учёта рабочего времени.

Обычный запуск — двойной щелчок по «Собрать рапорт» в папке выше.
Из командной строки: python3 sobrat_raport.py [папка_с_табелями]

Правила подсчёта повторяют то, как рапорт составляется вручную:
  * «в производстве» — операторы и сварщики с часами > 0 за ночь (D-1) и день (D);
  * стажёры и ВЦР входят в «в производстве» и дополнительно пишутся в свои строки;
  * статусы «Не в работе» / «На больничном» — только из дневной смены D.
"""
from __future__ import annotations

import datetime as dt
import os
import re
import subprocess
import sys
import traceback
import warnings
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

if getattr(sys, "frozen", False):  # собранный .exe лежит в корне папки программы
    ROOT_DIR = Path(sys.executable).resolve().parent
    APP_DIR = ROOT_DIR / "app"
else:
    APP_DIR = Path(__file__).resolve().parent
    ROOT_DIR = APP_DIR.parent
sys.path.insert(0, str(APP_DIR / "lib"))

import openpyxl  # noqa: E402
from openpyxl.styles import Font, PatternFill  # noqa: E402

warnings.filterwarnings("ignore", module="openpyxl")

TEMPLATE_PATH = APP_DIR / "Шаблон рапорта.xlsx"
TIMESHEETS_DIR = ROOT_DIR / "Табели"
REPORTS_DIR = ROOT_DIR / "Рапорты"
ERROR_LOG = ROOT_DIR / "ошибка.txt"

# ------------------------------------------------------------------ настройки
# Площадка определяется по словам в имени файла табеля (регистр не важен).
# Проверяются по порядку; если ничего не нашлось — DEFAULT_SITE.
SITE_KEYWORDS = [
    ("СПБ ЛАХТА", ("лахта",)),
    ("СПБ ВОЙСКОВИЦЫ", ("войсковиц", "спб")),
    ("ХЛМ", ("хлм",)),
    ("ДМД", ("дмд",)),
]
DEFAULT_SITE = "ДМД"

# Должность оператора -> строка рапорта, куда пишется стажёр этой должности.
OPERATOR_POSITIONS = {
    "РСС": "стажеры на РСС",
    "ПЛПК": "стажеры на ПЛПК",
    "ВЛ": "стажеры на ВЛ",
    "СГА": "стажеры на прочих РЦ",
    "ПЛР": "стажеры на прочих РЦ",
}
WELDER_TRAINEE_ROW = "стажеры-сварщики"

# Отметки в «Справке» у работавших, которые на рапорт не влияют.
NEUTRAL_NOTES = ("логист",)

# Ячейки рапорта, которые программа считает сама (остальные не трогает).
AUTO_ROWS = [
    "в производстве", WELDER_TRAINEE_ROW, "стажеры на РСС", "стажеры на ПЛПК",
    "стажеры на ВЛ", "стажеры на прочих РЦ", "ВЦР",
    "выходной", "простой", "алкоголь", "прогул", "простуда",
]

ALWAYS_MANUAL = [
    "Стропальщики, отъезд штатных, прибытие штатных по графику, кандидаты, "
    "перевахтовка и дата начала вахты — заполните, как обычно.",
    "Комментарии (уехали, нарушения и т.п.). Людей, уехавших после смены, "
    "программа из «в производстве» не вычитает.",
]

DAY, NIGHT = "день", "ночь"
SCAN_SUFFIXES = {".pdf", ".jpg", ".jpeg", ".png", ".heic", ".tif", ".tiff"}


# ------------------------------------------------------------------ данные
@dataclass
class Person:
    name: str
    position: str
    hours: float
    note: str
    kind: str          # "сварщик", "оператор" или "" — не учитывается
    trainee_row: str   # строка рапорта для стажёра этой должности


@dataclass
class Timesheet:
    path: Path
    site: str
    shift: str
    date: dt.date
    people: list = field(default_factory=list)
    warnings: list = field(default_factory=list)


@dataclass
class SiteResult:
    counts: Counter = field(default_factory=Counter)
    arrivals: Counter = field(default_factory=Counter)
    rows: list = field(default_factory=list)       # строки листа «Проверка»
    attention: list = field(default_factory=list)  # что дописать вручную


class UserError(Exception):
    """Понятная пользователю ошибка (без трассировки)."""


# ------------------------------------------------------------------ утилиты
def norm(value) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("ё", "е").replace("Ё", "Е").split()).lower()


def fmt(d: dt.date) -> str:
    return d.strftime("%d.%m.%Y")


def to_date(value, year_hint: int | None = None) -> dt.date | None:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, str):
        m = re.search(r"(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?", value)
        if m:
            day, month, year = int(m.group(1)), int(m.group(2)), m.group(3)
            if year:
                year = int(year) + (2000 if len(year) == 2 else 0)
            elif year_hint:
                year = year_hint
            else:
                return None
            try:
                return dt.date(year, month, day)
            except ValueError:
                return None
    return None


def to_hours(value) -> tuple[float, str]:
    """Часы и текст, если в ячейке часов написано не число (например, «Д/О»)."""
    if isinstance(value, (int, float)):
        return float(value), ""
    text = norm(value)
    if not text:
        return 0.0, ""
    try:
        return float(text.replace(",", ".")), ""
    except ValueError:
        return 0.0, str(value).strip()


def classify_position(position: str) -> tuple[str, str]:
    p = position.upper().replace("Ё", "Е")
    if "СВАРЩИК" in p:
        return "сварщик", WELDER_TRAINEE_ROW
    key = re.sub(r"[^А-ЯA-Z]", "", p)
    if key in OPERATOR_POSITIONS:
        return "оператор", OPERATOR_POSITIONS[key]
    if "ОПЕРАТОР" in p:
        return "оператор", "стажеры на прочих РЦ"
    return "", ""


def site_from_name(filename: str) -> tuple[str, bool]:
    low = filename.lower()
    for site, words in SITE_KEYWORDS:
        if any(w in low for w in words):
            return site, True
    return DEFAULT_SITE, False


def shift_from_name(stem: str) -> str | None:
    low = stem.lower()
    if "ноч" in low:
        return NIGHT
    if "ден" in low:
        return DAY
    m = re.search(r"[\s.](н|д)$", low.strip())
    if m:
        return NIGHT if m.group(1) == "н" else DAY
    return None


def person_key(name: str) -> str:
    return " ".join(norm(name).split()[:2])


# ------------------------------------------------------------------ чтение табеля
def read_timesheet(path: Path) -> Timesheet:
    """Читает табель .xlsx. Бросает ValueError, если файл не похож на табель."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    name_date = to_date(path.stem, year_hint=dt.date.today().year)

    shift, date = None, None
    for row in ws.iter_rows(min_row=1, max_row=8):
        row_shift, row_date = None, None
        for c in row:
            if row_shift is None and norm(c.value) in (DAY, NIGHT):
                row_shift = norm(c.value)
            elif row_date is None and not isinstance(c.value, dt.time):
                row_date = to_date(c.value)
        if row_shift:
            shift, date = row_shift, row_date
            break
    shift = shift or shift_from_name(path.stem)
    if not shift:
        raise ValueError("не найдена смена («День» / «Ночь»)")
    warns = []
    if date is None:
        date = name_date
    elif name_date and name_date != date:
        warns.append(
            f"{path.name}: в имени файла дата {fmt(name_date)}, а в самом табеле "
            f"{fmt(date)} — использована дата из табеля, проверьте"
        )
    if date is None:
        raise ValueError("не найдена дата табеля")

    header_row, tables = find_tables(ws)
    if not tables:
        raise ValueError("не найдена строка заголовков с «ФИО» и «Часы»")

    site, _ = site_from_name(path.name)
    ts = Timesheet(path=path, site=site, shift=shift, date=date, warnings=warns)
    for r in range(header_row + 1, ws.max_row + 1):
        for name_col, pos_col, hours_col, note_col in tables:
            name = str(ws.cell(r, name_col).value or "").strip()
            position = str(ws.cell(r, pos_col).value or "").strip()
            if not name or not position:
                continue
            # ФИО иногда разбито на две ячейки («Аджиев Сергей» | «Сираждинович»)
            tail = ws.cell(r, name_col + 1).value
            if name_col + 1 < pos_col and isinstance(tail, str) and tail.strip():
                name = f"{name} {tail.strip()}"
            hours, hours_text = to_hours(ws.cell(r, hours_col).value)
            note = str(ws.cell(r, note_col).value or "").strip() if note_col else ""
            if hours_text:
                note = f"{hours_text} {note}".strip()
            kind, trainee_row = classify_position(position)
            ts.people.append(Person(name, position, hours, note, kind, trainee_row))
    return ts


def find_tables(ws) -> tuple[int, list]:
    """Строка заголовков и колонки (ФИО, Должность, Часы, Справка) каждой таблицы на листе."""
    for row in ws.iter_rows(min_row=1, max_row=15):
        header = {i: norm(c.value) for i, c in enumerate(row, start=1) if c.value}
        fio_cols = [i for i, v in header.items() if v == "фио"]
        if not fio_cols:
            continue
        tables = []
        for n, start in enumerate(fio_cols):
            end = fio_cols[n + 1] if n + 1 < len(fio_cols) else 10_000
            cols = {i: v for i, v in header.items() if start < i < end}
            pos = next((i for i, v in cols.items() if v.startswith("должн")), None)
            hours = next((i for i, v in cols.items() if v == "часы"), None)
            note = next((i for i, v in cols.items() if v.startswith(("справ", "срав"))), None)
            if pos and hours:
                tables.append((start, pos, hours, note))
        return row[0].row, tables
    return 0, []


# ------------------------------------------------------------------ подсчёт
def absence(note: str, report_date: dt.date) -> tuple[str, object]:
    """Куда отнести человека с 0 часов в дневной смене: (вид, значение)."""
    n = norm(note)
    if not n:
        return "attention", "0 часов и пустая «Справка»"
    if "межвахт" in n:
        return "skip", "межвахта — не на площадке, не учитывается"
    m = re.search(r"\bс\s*(\d{1,2}\.\d{1,2}(?:\.\d{2,4})?)", n)
    if m:
        d = to_date(m.group(1), year_hint=report_date.year)
        if d and d < report_date - dt.timedelta(days=180):
            d = d.replace(year=d.year + 1)
        if d:
            return "arrival", d
    if n in ("д/о", "до", "д.о", "д/о.") or "выходн" in n:
        return "row", "выходной"
    if "бол" in n or "простуд" in n or "б/л" in n:
        return "row", "простуда"
    if "прогул" in n:
        return "row", "прогул"
    if "алкогол" in n:
        return "row", "алкоголь"
    if "простой" in n:
        return "row", "простой"
    return "attention", f"отметка «{note}» — программа не знает, куда её отнести"


def count_site(sheets: list, report_date: dt.date) -> SiteResult:
    res = SiteResult()
    for ts in sorted(sheets, key=lambda t: t.shift != NIGHT):
        for p in ts.people:
            how, flag = count_person(p, ts, report_date, res)
            res.rows.append([ts.site, ts.shift, p.name, p.position, p.hours, p.note, how, ts.path.name, flag])
    return res


def count_person(p: Person, ts: Timesheet, report_date: dt.date, res: SiteResult) -> tuple[str, bool]:
    if not p.kind:
        return "не учитывается (не оператор и не сварщик)", False
    n = norm(p.note)
    if p.hours > 0:
        res.counts["в производстве"] += 1
        how = "в производстве"
        if "стаж" in n:
            res.counts[p.trainee_row] += 1
            how += f" + {p.trainee_row}"
        if "вцр" in n:
            res.counts["ВЦР"] += 1
            how += " + ВЦР"
        known = "стаж" in n or "вцр" in n or any(w in n for w in NEUTRAL_NOTES)
        if n and not known:
            return how + f" (отметка «{p.note}» не распознана)", True
        return how, False
    if ts.shift == NIGHT:
        return "не учитывается: 0 часов в ночную смену", False
    kind, value = absence(p.note, report_date)
    if kind == "row":
        res.counts[value] += 1
        return value, False
    if kind == "arrival":
        res.arrivals[value] += 1
        return f"прибытие штатных {fmt(value)}", False
    if kind == "skip":
        return value, False
    res.attention.append(f"{ts.site}: {p.name} ({p.position}, день {fmt(ts.date)}) — {value}. Впишите вручную.")
    return f"НЕ УЧТЁН: {value}", True


# ------------------------------------------------------------------ запись рапорта
def find_blocks(ws) -> dict:
    blocks = {}
    for r in range(1, ws.max_row + 1):
        if norm(ws.cell(r, 5).value) == "в работе":
            blocks[" ".join(str(ws.cell(r, 1).value).split()).upper()] = r
    return blocks


def fill_block(ws, h: int, res: SiteResult, report_date: dt.date) -> list:
    if norm(ws.cell(h + 7, 11).value) != "далее":
        raise UserError("Шаблон рапорта изменён: не найдена строка «далее» в блоке прибытия.")
    cells = {}
    for r in range(h + 1, h + 9):
        for label_col in (5, 7, 9):  # E, G, I — подписи; значение в соседней колонке
            label = ws.cell(r, label_col).value
            if isinstance(label, str):
                cells.setdefault(label.strip(), ws.cell(r, label_col + 1))
    for label in AUTO_ROWS:
        if label not in cells:
            raise UserError(f"Шаблон рапорта изменён: не найдена строка «{label}».")
        cells[label].value = res.counts.get(label) or None
    notes = []
    for d, n in sorted(res.arrivals.items()):
        delta = (d - report_date).days
        if delta < 0:
            notes.append(f"прибытие {fmt(d)} — дата уже прошла, не внесено")
            continue
        cell = ws.cell(h + 2 + delta if delta <= 4 else h + 7, 12)
        cell.value = (cell.value if isinstance(cell.value, (int, float)) else 0) + n
    return notes


def write_check_sheet(wb, report_date, todo, used, unused, rows):
    ws = wb.create_sheet("Проверка")
    bold = Font(bold=True)
    yellow = PatternFill("solid", fgColor="FFF2CC")

    def section(title):
        ws.append([])
        ws.append([title])
        ws.cell(ws.max_row, 1).font = bold

    ws.append([f"Проверка рапорта за {fmt(report_date)}"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([f"Учтены смены: ночь {fmt(report_date - dt.timedelta(days=1))} и день {fmt(report_date)}. "
               "Перед отправкой рапорта этот лист можно удалить."])
    section("ЗАПОЛНИТЕ ИЛИ ПРОВЕРЬТЕ ВРУЧНУЮ")
    for t in todo:
        ws.append([t])
        ws.cell(ws.max_row, 1).fill = yellow
    section("ИСПОЛЬЗОВАННЫЕ ТАБЕЛИ")
    for ts in used:
        ws.append([f"{ts.site} — {ts.shift} {fmt(ts.date)} — {ts.path.name}"])
    if unused:
        section("НЕ ИСПОЛЬЗОВАНЫ")
        for line in unused:
            ws.append([line])
    section("КАК УЧТЁН КАЖДЫЙ ЧЕЛОВЕК")
    ws.append(["Площадка", "Смена", "ФИО", "Должность", "Часы", "Справка", "Как учтён", "Файл"])
    for c in ws[ws.max_row]:
        c.font = bold
    for row in rows:
        ws.append(row[:-1])
        if row[-1]:
            for c in ws[ws.max_row]:
                c.fill = yellow
    for col, width in zip("ABCDEFGH", (18, 8, 38, 16, 7, 20, 55, 38)):
        ws.column_dimensions[col].width = width


# ------------------------------------------------------------------ основной сценарий
def collect(folder: Path):
    sheets, unreadable = [], []
    files = sorted(p for p in folder.rglob("*") if p.is_file())
    for path in files:
        if path.name.startswith(("~$", ".")):
            continue
        suffix = path.suffix.lower()
        if suffix in SCAN_SUFFIXES:
            unreadable.append((path, "скан или фото — программа такие файлы не читает"))
        elif suffix == ".xls":
            unreadable.append((path, "старый формат Excel — откройте и сохраните как .xlsx"))
        elif suffix in (".xlsx", ".xlsm"):
            try:
                sheets.append(read_timesheet(path))
            except Exception as e:  # битый файл или не табель
                unreadable.append((path, f"не похоже на табель: {e}"))
    return sheets, unreadable


def ask_report_date(default: dt.date) -> dt.date:
    if not sys.stdin.isatty():
        return default
    while True:
        answer = input(
            f"\nДата рапорта: {fmt(default)} (ночь {fmt(default - dt.timedelta(days=1))} + день {fmt(default)}).\n"
            "Нажмите Enter, чтобы продолжить, или введите другую дату (ДД.ММ.ГГГГ): "
        ).strip()
        if not answer:
            return default
        d = to_date(answer, year_hint=default.year)
        if d:
            return d
        print("Не понял дату. Пример: 07.10.2026")


def unique_path(path: Path) -> Path:
    n = 2
    candidate = path
    while candidate.exists():
        candidate = path.with_name(f"{path.stem} ({n}){path.suffix}")
        n += 1
    return candidate


def open_file(path: Path):
    try:
        if sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        elif os.name == "nt":
            os.startfile(str(path))  # noqa: S606
    except Exception:
        pass  # не открылся — не страшно, путь напечатан


def run(folder: Path) -> Path:
    if not TEMPLATE_PATH.exists():
        raise UserError(f"Не найден шаблон рапорта: {TEMPLATE_PATH}")
    if not folder.is_dir():
        raise UserError(f"Не найдена папка с табелями: {folder}")

    sheets, unreadable = collect(folder)
    if not sheets:
        raise UserError(f"В папке «{folder.name}» нет табелей Excel (.xlsx). Положите их туда и запустите снова.")

    day_dates = [t.date for t in sheets if t.shift == DAY]
    default = max(day_dates) if day_dates else max(t.date for t in sheets) + dt.timedelta(days=1)
    report_date = ask_report_date(default)
    night_date = report_date - dt.timedelta(days=1)

    used = [t for t in sheets
            if (t.shift == DAY and t.date == report_date) or (t.shift == NIGHT and t.date == night_date)]
    other_dates = len(sheets) - len(used)

    todo = []
    for t in used:
        todo.extend(t.warnings)
        if not site_from_name(t.path.name)[1]:
            todo.append(f"{t.path.name}: в имени файла нет названия площадки — посчитан как {DEFAULT_SITE}. Если это не так, переименуйте файл.")

    relevant_unreadable = [
        (p, why) for p, why in unreadable
        if to_date(p.stem, year_hint=report_date.year) in (None, report_date, night_date)
    ]
    unused_lines = [f"{p.name} — {why}" for p, why in relevant_unreadable]
    if other_dates:
        unused_lines.append(f"Табели за другие даты: {other_dates} шт. (не нужны для этого рапорта)")

    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb.worksheets[0]
    blocks = find_blocks(ws)
    ws["B1"].value = dt.datetime.combine(report_date, dt.time())

    for site, _ in SITE_KEYWORDS:
        if site not in blocks:
            raise UserError(f"В шаблоне рапорта нет блока площадки «{site}».")
    scans = [p.name for p, _ in relevant_unreadable if p.suffix.lower() in SCAN_SUFFIXES | {".xls"}]

    summary, all_rows, seen = [], [], {}
    for site in sorted((s for s, _ in SITE_KEYWORDS), key=blocks.get):
        site_sheets = [t for t in used if t.site == site]
        if not site_sheets:
            todo.append(f"{site}: нет табелей за эти смены — блок не заполнен, внесите цифры вручную.")
            summary.append(f"{site}: табелей нет")
            continue
        for shift, d in ((NIGHT, night_date), (DAY, report_date)):
            if not any(t.shift == shift for t in site_sheets):
                hint = ""
                if scans:
                    hint = f" Возможно, он среди файлов, которые программа не читает: {', '.join(scans)}."
                todo.append(f"{site}: нет табеля смены «{shift}» за {fmt(d)} — в рапорте учтена только "
                            f"другая смена, добавьте эту вручную.{hint}")
        res = count_site(site_sheets, report_date)
        todo.extend(res.attention)
        todo.extend(f"{site}: {n}" for n in fill_block(ws, blocks[site], res, report_date))
        all_rows.extend(res.rows)
        for t in site_sheets:
            for p in t.people:
                if p.kind and p.hours > 0:
                    seen.setdefault(person_key(p.name), []).append(f"{t.site} ({t.shift})")
        parts = [f"{k} {res.counts[k]}" for k in AUTO_ROWS if res.counts.get(k)]
        parts += [f"прибытие {d.strftime('%d.%m')}: {n}" for d, n in sorted(res.arrivals.items())]
        summary.append(f"{site}: " + ", ".join(parts))

    for key, places in seen.items():
        if len(places) > 1:
            todo.append(f"«{key.title()}» отработал в нескольких табелях: {', '.join(places)} — проверьте, нет ли ошибки.")
    todo.extend(ALWAYS_MANUAL)

    used.sort(key=lambda t: (blocks[t.site], t.shift != NIGHT))
    write_check_sheet(wb, report_date, todo, used, unused_lines, all_rows)
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True

    REPORTS_DIR.mkdir(exist_ok=True)
    out = unique_path(REPORTS_DIR / f"Рапорт по персоналу {fmt(report_date)}.xlsx")
    try:
        wb.save(out)
    except PermissionError:
        raise UserError(f"Не удалось сохранить {out.name}: файл открыт в Excel? Закройте его и запустите снова.")

    print(f"\nРапорт за {fmt(report_date)}")
    for line in summary:
        print("  " + line)
    print("\nЗаполните или проверьте вручную (это же есть на листе «Проверка»):")
    for t in todo:
        print("  • " + t)
    print(f"\nГотово: {out}")
    return out


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    folder = Path(sys.argv[1]) if len(sys.argv) > 1 else TIMESHEETS_DIR
    if folder == TIMESHEETS_DIR:
        folder.mkdir(exist_ok=True)
    try:
        out = run(folder)
    except UserError as e:
        print(f"\n{e}")
        return 1
    except Exception:
        ERROR_LOG.write_text(traceback.format_exc(), encoding="utf-8")
        print(f"\nЧто-то пошло не так. Подробности записаны в файл «{ERROR_LOG.name}» — перешлите его тому, кто настраивал программу.")
        return 1
    if sys.stdin.isatty():
        open_file(out)
    return 0


if __name__ == "__main__":
    exit_code = main()
    if getattr(sys, "frozen", False) and sys.stdin.isatty():  # у .exe нет .bat с pause
        input("\nНажмите Enter, чтобы закрыть окно ")
    sys.exit(exit_code)
