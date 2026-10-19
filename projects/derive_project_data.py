"""Производные файлы для итоговых проектов из витрин data/dataset_student.

Источник только шесть CSV витрин. Скрипт детерминирован (SEED = 42) и ничего не пишет
за пределы каталога teacher/projects-data/.

    python3 teacher/projects-data/derive_project_data.py

Результат, по папке на проект; файлы данных всегда в подпапке data/:
  p2-batch-normalize/  data/registry_01..20.xlsx, data/stations.csv
  p3-telegram-files/   то же, что p2 (копия)
  p1-mail-orders/      только с флагом --with-letters: то же плюс data/letters/letter_01..12.txt (П1 не назначен)
  p4-wagon-cycle/      data/wagon_cycles.csv.gz (вариант А), data/unload_daily.csv.gz (вариант Б)
  Эталоны и статистика (expected.json, FILES.md, stats.json) пишутся только с флагом --with-keys в .scratch/2026-10-19-project-keys/
  p6-forecast-subscription/  data/: копии витрин unload, month_plan, unload_forecast и манифеста
  i3-blogger/, i5-reels/, c1-plpgsql-docs/  скрипт не трогает
"""
from __future__ import annotations

import datetime as dt
import json
import random
import shutil
import sys
from collections import defaultdict
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

SEED = 42
ROOT = Path(__file__).resolve().parents[2]
SRC = ROOT / "data" / "dataset_student"
OUT = Path(__file__).resolve().parent
P1 = OUT / "p2-batch-normalize"          # основной выход реестров (исторически П1; П1 не назначен и папка удалена)
COPIES = [OUT / "p3-telegram-files"]
WITH_LETTERS = "--with-letters" in sys.argv
WITH_KEYS = "--with-keys" in sys.argv        # эталоны и статистика (expected.json, FILES.md, stats.json) — только по флагу, студентам не отдаются
KEYS = ROOT / ".scratch" / "2026-10-19-project-keys"  # письма для П1 собираются только по флагу, в p1-mail-orders/
rng = random.Random(SEED)

ROD_WORDS = {  # код рода ПС -> формы: 1 штука, 2-4 штуки, 5+ штук, именительный множественного
    "ПВ": ("полувагон", "полувагона", "полувагонов", "полувагоны"),
    "КР": ("крытый вагон", "крытых вагона", "крытых вагонов", "крытые вагоны"),
    "ПЛ": ("платформа", "платформы", "платформ", "платформы"),
    "ЦС": ("цистерна", "цистерны", "цистерн", "цистерны"),
    "ЦМВ": ("цельнометаллический вагон", "цельнометаллических вагона", "цельнометаллических вагонов", "цельнометаллические вагоны"),
    "ПЛФИТ": ("фитинговая платформа", "фитинговые платформы", "фитинговых платформ", "фитинговые платформы"),
}


def agree(n: int, forms: tuple) -> str:
    """Согласование существительного с числительным: 1 вагон, 3 вагона, 5 вагонов."""
    if n % 10 == 1 and n % 100 != 11:
        return forms[0]
    if n % 10 in (2, 3, 4) and n % 100 not in (12, 13, 14):
        return forms[1]
    return forms[2]
CARGO_WORDS = {  # группа груза витрины -> разговорная форма в письме
    "ПЕСОК И ГРАВИЙ": "песок",
    "КАУЧУК И РЕЗИНОТЕХНИКА": "каучук",
    "КЕРАМИКА И САНТЕХНИКА": "керамика",
    "СТЕКЛОТАРА": "стеклотара",
    "МАСЛА ТЕХНИЧЕСКИЕ": "техническое масло",
    "ТЕКСТИЛЬ И ВОЛОКНО": "текстиль",
    "ГЛИНОЗЕМ": "глинозём",
    "ТАРА И УПАКОВКА": "упаковка",
    "КРАХМАЛ И ПАТОКА": "крахмал",
    "МЕБЕЛЬ И ДЕРЕВОИЗДЕЛИЯ": "мебель",
}
MONTHS = {1: "января", 2: "февраля", 3: "марта", 4: "апреля", 5: "мая", 6: "июня",
          7: "июля", 8: "августа", 9: "сентября", 10: "октября", 11: "ноября", 12: "декабря"}


# ---------------------------------------------------------------- helpers
def title_case(s: str) -> str:
    return s.capitalize()


def fmt_date(d: dt.date, style: int) -> tuple[str, str]:
    """Возвращает (строка, точность): точность 'day' | 'decade' | 'decade_no_year'."""
    if style == 0:
        return d.strftime("%d.%m.%Y"), "day"
    if style == 1:
        return d.strftime("%Y-%m-%d"), "day"
    decade = "начало" if d.day <= 10 else ("середина" if d.day <= 20 else "конец")
    if style == 2:
        return f"{decade} {MONTHS[d.month]} {d.year}", "decade"
    return f"{decade} {MONTHS[d.month]}", "decade_no_year"


def fmt_wagons(n: int, style: int) -> str:
    return [str(n), f"{n} шт", f"{n} ваг.", ""][style]


def typo(name: str) -> str:
    """Перестановка двух соседних букв около середины слова (как в практике 4)."""
    chars = list(name)
    idx = len(chars) // 2
    for off in range(len(chars) - 1):
        i = idx + (off // 2 if off % 2 == 0 else -(off // 2 + 1))
        if 0 <= i < len(chars) - 1 and chars[i] != chars[i + 1] and chars[i].isalpha() and chars[i + 1].isalpha():
            chars[i], chars[i + 1] = chars[i + 1], chars[i]
            return "".join(chars)
    return name + name[-1]


# ---------------------------------------------------------------- stations
def build_stations() -> pd.DataFrame:
    spec = {
        "V_DM_LOAD.csv": [("ST_CODE", "ST_NAME", "RW_CODE", "RW_NAME"), ("RSV_ST_CODE", "RSV_ST_NAME", "RSV_RW_CODE", "RSV_RW_NAME")],
        "V_DM_UNLOAD.csv": [("ST_CODE", "ST_NAME", "RW_CODE", "RW_NAME"), ("SND_ST_CODE", "SND_ST_NAME", "SND_RW_CODE", "SND_RW_NAME")],
        "V_DM_GU12.csv": [("ST_CODE", "ST_NAME", "RW_CODE", "RW_NAME"), ("RSV_ST_CODE", "RSV_ST_NAME", "RSV_RW_CODE", "RSV_RW_NAME")],
        "V_DM_MPP.csv": [("LOAD_ST_CODE", "LOAD_ST_NAME", "LOAD_RW_CODE", "LOAD_RW_NAME"), ("UNLOAD_ST_CODE", "UNLOAD_ST_NAME", "UNLOAD_RW_CODE", "UNLOAD_RW_NAME")],
        "V_DM_UNLOAD_FORECAST.csv": [("SND_ST_CODE", "SND_ST_NAME", "SND_RW_CODE", "SND_RW_NAME"), ("RSV_ST_CODE", "RSV_ST_NAME", "RSV_RW_CODE", "RSV_RW_NAME")],
    }
    parts = []
    sources = defaultdict(set)
    for fname, groups in spec.items():
        cols = sorted({c for g in groups for c in g})
        df = pd.read_csv(SRC / fname, usecols=cols, dtype=str)
        for g in groups:
            sub = df[list(g)].dropna().drop_duplicates()
            sub.columns = ["st_code", "st_name", "rw_code", "rw_name"]
            parts.append(sub)
            for code in sub["st_code"]:
                sources[code].add(fname.replace("V_DM_", "").replace(".csv", "").lower())
    st = pd.concat(parts).drop_duplicates().sort_values(["rw_code", "st_code"]).reset_index(drop=True)
    conflicts = st.groupby("st_code")["st_name"].nunique()
    assert (conflicts == 1).all(), f"код станции с разными именами: {conflicts[conflicts > 1].index.tolist()}"
    st["sources"] = st["st_code"].map(lambda c: ";".join(sorted(sources[c])))
    (P1 / "data").mkdir(parents=True, exist_ok=True)
    st.to_csv(P1 / "data" / "stations.csv", index=False)
    return st


# ---------------------------------------------------------------- GU-12 -> registries and letters
def load_gu12() -> pd.DataFrame:
    df = pd.read_csv(SRC / "V_DM_GU12.csv", dtype=str)
    df["day"] = df["REPDATE"].str[:10]
    df["GU12_WAG_CNT"] = df["GU12_WAG_CNT"].astype(int)
    keys = ["SND_ORG_NAME", "day", "ST_CODE", "ST_NAME", "RSV_ST_CODE", "RSV_ST_NAME", "GG_NAME", "ROD_NAME"]
    agg = df.groupby(keys, as_index=False)["GU12_WAG_CNT"].sum().rename(columns={"GU12_WAG_CNT": "wagons"})
    return agg


def pick_registry_slices(agg: pd.DataFrame, n: int) -> list[tuple[str, str]]:
    sizes = agg.groupby(["SND_ORG_NAME", "day"]).size()
    cands = sizes[(sizes >= 15) & (sizes <= 45)].reset_index()
    cands = cands.sort_values(["SND_ORG_NAME", "day"]).values.tolist()
    rng.shuffle(cands)
    chosen, per_client = [], defaultdict(int)
    for cap in (1, 2, 3, 4, 5, 6):  # сначала по одному файлу на клиента, потом добираем
        for client, day, _ in cands:
            if len(chosen) == n:
                return chosen
            if per_client[client] < cap and (client, day) not in chosen:
                chosen.append((client, day))
                per_client[client] += 1
    return chosen


# Семь раскладок. header_row None — заголовков нет. col0 — с какой колонки начинается таблица.
LAYOUTS = {
    "L1": dict(desc="шапка документа 4 строки, объединённая ячейка E5:F5 «Маршрут», заголовки в строке 6",
               header_row=6, col0=1, cols=["no", "client", "cargo", "rod", "from", "to", "wagons", "date"],
               headers=["№", "Клиент", "Груз", "Род вагона", "Ст. отправления", "Ст. назначения", "Кол-во вагонов", "Дата"]),
    "L2": dict(desc="заголовки в строке 3 с другими именами колонок, род вагона словами",
               header_row=3, col0=1, cols=["client", "cargo", "rod_words", "from", "to", "wagons", "date"],
               headers=["Заказчик", "Груз", "Тип вагона", "Станция отпр.", "Станция назн.", "Вагонов, шт", "Срок"]),
    "L3": dict(desc="без заголовков, порядок колонок: дата, откуда, куда, груз, род, вагонов, клиент; подпись после таблицы",
               header_row=None, col0=1, cols=["date", "from", "to", "cargo", "rod", "wagons", "client"], headers=None),
    "L4": dict(desc="двухуровневая шапка: строка 1 объединённые группы «Заявка» и «Маршрут», строка 2 имена; пустая колонка-разделитель между группами",
               header_row=2, col0=1, cols=["client", "cargo", "rod", "gap", "from", "to", "wagons", "date"],
               headers=["Клиент", "Груз", "Род ПС", None, "Откуда", "Куда", "Вагонов", "Дата заявки"]),
    "L5": dict(desc="таблица начинается с колонки C, заголовки в строке 1; станции записаны как «код название»; даты — настоящие даты Excel",
               header_row=1, col0=3, cols=["no", "client", "from_coded", "to_coded", "cargo", "rod", "wagons", "date_cell"],
               headers=["п/п", "Грузоотправитель", "Станция отправления", "Станция назначения", "Груз", "Род", "Кол-во", "Дата"]),
    "L6": dict(desc="первый лист «Примечания» без данных, таблица на втором листе с заголовками в строке 2; груз и род в одной ячейке через «/», маршрут одной ячейкой через тире",
               header_row=2, col0=1, cols=["client", "cargo_rod", "route", "wagons", "date"],
               headers=["Клиент", "Груз / род ПС", "Маршрут", "Вагонов", "Дата"], sheet2=True),
    "L7": dict(desc="заголовки строчными буквами с переносами строк внутри ячеек; клиента в таблице нет, он только в заголовке листа; пустая строка после каждых 10 заявок",
               header_row=2, col0=1, cols=["no", "cargo", "rod", "from", "to", "wagons_float", "date"],
               headers=["№", "груз", "род\nвагона", "станция\nотправления", "станция\nназначения", "кол-во\nвагонов", "дата"], blank_every=10),
}
LAYOUT_ORDER = ["L1", "L2", "L3", "L4", "L5", "L6", "L7"]
EXTRA_DEFECTS = ["trailing_header_spaces", "lowercase_station", "nd_wagons", "note_row", "two_digit_year",
                 "cargo_lowercase_all", "client_variant", "dash_variants"]


def write_registry(idx: int, layout: str, client: str, day: str, rows: pd.DataFrame, stations: set[str]) -> dict:
    L = LAYOUTS[layout]
    wb = Workbook()
    ws = wb.active
    if L.get("sheet2"):
        ws.title = "Примечания"
        ws["A1"] = "Реестр сформирован автоматически из системы заявок"
        ws["A2"] = "Вопросы по заявкам — в отдел логистики"
        ws = wb.create_sheet("Данные")
    else:
        ws.title = "Реестр"
    snapshot = dt.date.fromisoformat(day)
    defects: list[str] = []
    clean: list[dict] = []
    recs = rows.sort_values(["ST_NAME", "RSV_ST_NAME", "GG_NAME"]).reset_index(drop=True)
    n = len(recs)
    typo_rows = sorted(rng.sample(range(n), 2))
    space_row, upper_row, blank_after = rng.sample(range(n), 3)
    dup_src = sorted(rng.sample(range(n), 2))
    extras = rng.sample(EXTRA_DEFECTS, rng.randint(1, 3))
    extra_rows = {e: rng.randrange(n) for e in extras}
    c0 = L["col0"]
    hdr_row = L["header_row"]

    # --- шапка документа
    if layout == "L1":
        ws["A1"] = f"{client} — реестр заявок ГУ-12 на подачу вагонов"
        ws["A1"].font = Font(bold=True)
        ws["A2"] = f"По состоянию на {snapshot.strftime('%d.%m.%Y')}"
        ws["A3"] = "Утвердил: начальник отдела перевозок"
        ws.merge_cells("E5:F5")
        ws["E5"] = "Маршрут"
        ws["E5"].alignment = Alignment(horizontal="center")
    elif layout == "L2":
        ws["A1"] = f"Реестр заявок {client}, снимок {snapshot.strftime('%d.%m.%Y')}"
    elif layout == "L4":
        ws.merge_cells("A1:C1"); ws["A1"] = "Заявка"
        ws.merge_cells("E1:H1"); ws["E1"] = "Маршрут и объём"
        for c in ("A1", "E1"):
            ws[c].alignment = Alignment(horizontal="center"); ws[c].font = Font(bold=True)
    elif layout == "L6":
        ws["A1"] = f"{client}, заявки на {snapshot.strftime('%d.%m.%Y')}"
    elif layout == "L7":
        ws["A1"] = f"{client} — заявки на подачу, {snapshot.strftime('%d.%m.%Y')}"
    defects.append(L["desc"])
    if L["headers"]:
        for k, t in enumerate(L["headers"]):
            if t is None:
                continue
            txt = t + "  " if "trailing_header_spaces" in extras else t
            cell = ws.cell(row=hdr_row, column=c0 + k, value=txt)
            cell.font = Font(bold=True)
            if "\n" in t:
                cell.alignment = Alignment(wrap_text=True)
    r = (hdr_row + 1) if hdr_row else 1

    written: list[list] = []
    for i, rec in recs.iterrows():
        date_s, prec = fmt_date(snapshot, (i + idx) % 4)
        wag_s = fmt_wagons(int(rec.wagons), (i + 2 * idx) % 4)
        frm, to, cargo, cl = rec.ST_NAME, rec.RSV_ST_NAME, title_case(rec.GG_NAME), client
        rod = rec.ROD_NAME
        rowdef = []
        if i == typo_rows[0]:
            frm = typo(frm); rowdef.append("опечатка в станции отправления")
        if i == typo_rows[1]:
            to = typo(to); rowdef.append("опечатка в станции назначения")
        if i == space_row:
            cl = f"  {cl}  "; rowdef.append("пробелы вокруг клиента")
        if i == upper_row:
            cargo = cargo.upper(); rowdef.append("груз в верхнем регистре")
        if "cargo_lowercase_all" in extras:
            cargo = cargo.lower()
        if extra_rows.get("lowercase_station") == i:
            frm = frm.lower(); rowdef.append("станция строчными буквами")
        if extra_rows.get("nd_wagons") == i:
            wag_s = "н/д"; rowdef.append("количество вагонов «н/д»")
        if extra_rows.get("two_digit_year") == i:
            date_s, prec = snapshot.strftime("%d.%m.%y"), "day"; rowdef.append("дата с двузначным годом")
        if extra_rows.get("client_variant") == i:
            cl = cl.replace("Компания", "ООО «Компания") + "»"; rowdef.append("клиент с организационной формой")
        if wag_s == "":
            rowdef.append("количество вагонов пусто")
        if prec != "day":
            rowdef.append("дата текстом, точность до декады" + (" без года" if prec == "decade_no_year" else ""))
        dash = " — " if "dash_variants" in extras and i % 2 else " - "
        values = {
            "no": i + 1, "client": cl, "cargo": cargo, "rod": rod, "rod_words": ROD_WORDS[rod][0], "gap": None,
            "from": frm, "to": to, "wagons": wag_s, "date": date_s,
            "from_coded": f"{rec.ST_CODE} {frm}", "to_coded": f"{rec.RSV_ST_CODE} {to}",
            "date_cell": dt.datetime(snapshot.year, snapshot.month, snapshot.day) if prec == "day" else date_s,
            "cargo_rod": f"{cargo} / {rod}", "route": f"{frm}{dash}{to}",
            "wagons_float": (float(rec.wagons) if wag_s not in ("", "н/д") else wag_s),
        }
        if i == blank_after:
            r += 1; defects.append(f"строка {r - 1} — пустая строка внутри таблицы")
        if L.get("blank_every") and i and i % L["blank_every"] == 0:
            r += 1
        if extra_rows.get("note_row") == i:
            ws.cell(row=r, column=c0, value="см. примечание: заявки ниже согласованы по телефону")
            defects.append(f"строка {r} — текстовое примечание внутри таблицы"); r += 1
        vals = [values[c] for c in L["cols"]]
        for k, v in enumerate(vals):
            if v is not None:
                ws.cell(row=r, column=c0 + k, value=v)
        written.append(vals)
        clean.append({"excel_row": r, "client": client, "cargo": rec.GG_NAME, "rod": rec.ROD_NAME,
                      "from": rec.ST_NAME, "to": rec.RSV_ST_NAME, "wagons": int(rec.wagons),
                      "date": day, "date_precision": prec, "defects": rowdef})
        r += 1
    dup_rows = []
    for s_ in dup_src:
        for k, v in enumerate(written[s_]):
            if v is not None:
                ws.cell(row=r, column=c0 + k, value=v)
        dup_rows.append(r); r += 1
    defects.append(f"строки {dup_rows} — точные дубли строк {[clean[s_]['excel_row'] for s_ in dup_src]}")
    if layout in ("L3", "L7"):
        ws.cell(row=r, column=c0, value="Исполнитель: отдел логистики")
        ws.cell(row=r + 1, column=c0, value=f"тел. доб. {100 + idx}")
        defects.append(f"строки {r}-{r + 1} — подпись после таблицы")
    else:
        wcol = L["cols"].index("wagons") if "wagons" in L["cols"] else L["cols"].index("wagons_float")
        ws.cell(row=r, column=c0, value="ИТОГО")
        ws.cell(row=r, column=c0 + wcol, value=int(recs.wagons.sum()))
        defects.append(f"строка {r} — «ИТОГО» внутри табличного диапазона")
    defects.append("даты чередуются по четырём форматам; количество вагонов по четырём стилям («15», «15 шт», «15 ваг.», пусто)")
    names = {"trailing_header_spaces": "пробелы в конце имён заголовков", "cargo_lowercase_all": "весь груз строчными буквами",
             "dash_variants": "в маршруте чередуются дефис и тире"}
    for e in extras:
        if e in names:
            defects.append(names[e])

    fname = f"registry_{idx:02d}.xlsx"
    (P1 / "data").mkdir(parents=True, exist_ok=True)
    wb.save(P1 / "data" / fname)
    assert all(c["from"] in stations and c["to"] in stations for c in clean)
    return {"file": fname, "layout": layout, "sheet": ws.title, "header_row": hdr_row, "first_column": c0, "client": client,
            "snapshot_date": day, "rows_written": len(written) + len(dup_rows), "duplicates": len(dup_rows),
            "clean_rows": len(clean), "wagons_total": int(recs.wagons.sum()), "extra_defects": extras,
            "defects": defects, "clean": clean}


LETTER_TEMPLATES = [
    "Коллеги, добрый день!\nНужно {n} {rod_n} из {frm} на {to}{when}, груз - {cargo}.\nЗаранее спасибо, ждём подтверждения.\n",
    "Здравствуйте!\nПрошу подать {n} {rod_n} под {cargo_acc}, станция погрузки {frm},\nвыгрузка {to}.{when_full}\nС уважением, отдел логистики {client}.\n",
    "добрый день\nнужны {rod_nom} под {cargo_acc} {n} шт, ст. {frm} - {to}{when_short}\nочень срочно, подтвердите пожалуйста\n",
    "Добрый день.\n{Cargo}, {n} {wag_n}, {frm} - {to}{when_short}.\nРод вагона - {rod}. {client}.\n",
]
ACC = {"стеклотара": "стеклотару", "керамика": "керамику", "упаковка": "упаковку", "мебель": "мебель"}


def write_letters(agg: pd.DataFrame, used: set[tuple[str, str]], stations: set[str]) -> list[dict]:
    pool = agg[agg.wagons >= 5].copy()
    pool = pool[~pool.apply(lambda r: (r.SND_ORG_NAME, r.day) in used, axis=1)]
    pool = pool.sample(frac=1, random_state=SEED).drop_duplicates("SND_ORG_NAME").head(12).reset_index(drop=True)
    assert len(pool) == 12
    (OUT / "p1-mail-orders" / "data" / "letters").mkdir(parents=True, exist_ok=True)
    out = []
    for i, rec in pool.iterrows():
        d = dt.date.fromisoformat(rec.day)
        frm, to = rec.ST_NAME.capitalize(), rec.RSV_ST_NAME.capitalize()
        defects = []
        if i in (2, 6, 10):
            if i == 6:
                to = typo(to)
            else:
                frm = typo(frm)
            defects.append("опечатка в станции")
        has_date = i < 9
        date_s, prec = fmt_date(d, i % 4) if has_date else ("", None)
        if not has_date:
            defects.append("дата не указана")
        elif prec != "day":
            defects.append("дата текстом" + (" без года" if prec == "decade_no_year" else ""))
        cargo = CARGO_WORDS[rec.GG_NAME]
        forms = ROD_WORDS[rec.ROD_NAME]
        n = int(rec.wagons)
        body = LETTER_TEMPLATES[i % 4].format(
            n=n, rod_n=agree(n, forms), rod_nom=forms[3], wag_n=agree(n, ("вагон", "вагона", "вагонов")),
            rod=rec.ROD_NAME if i % 2 else forms[0], frm=frm, to=to,
            cargo=cargo, Cargo=cargo.capitalize(), cargo_acc=ACC.get(cargo, cargo), client=rec.SND_ORG_NAME,
            when=f" {date_s}" if has_date and prec != "day" else (f", погрузка {date_s}" if has_date else ""),
            when_full=f" Дата погрузки {date_s}." if has_date else "",
            when_short=f", погрузка {date_s}" if has_date else "",
        )
        fname = f"letter_{i + 1:02d}.txt"
        (OUT / "p1-mail-orders" / "data" / "letters" / fname).write_text(body, encoding="utf-8")
        assert rec.ST_NAME in stations and rec.RSV_ST_NAME in stations
        out.append({"file": fname, "client": rec.SND_ORG_NAME, "cargo": rec.GG_NAME, "rod": rec.ROD_NAME,
                    "from": rec.ST_NAME, "to": rec.RSV_ST_NAME, "wagons": int(rec.wagons),
                    "date": rec.day if has_date else None, "date_precision": prec, "defects": defects})
    return out


# ---------------------------------------------------------------- load -> wagon cycles
def build_wagon_cycles() -> dict:
    cols = ["REPDATE", "WAGNUM", "ST_CODE", "ST_NAME", "RW_NAME", "RSV_ST_CODE", "RSV_ST_NAME", "RSV_RW_NAME",
            "ROD_NAME", "GG_NAME", "FR_CLASS", "C_OWNER_ID", "CLIENT_NAME"]
    df = pd.read_csv(SRC / "V_DM_LOAD.csv", usecols=cols, dtype=str)
    df["REPDATE"] = pd.to_datetime(df["REPDATE"])
    before = len(df)
    df = df.drop_duplicates(["WAGNUM", "REPDATE"]).sort_values(["WAGNUM", "REPDATE"]).reset_index(drop=True)
    dropped = before - len(df)
    g = df.groupby("WAGNUM", sort=False)
    df["next_load_dt"] = g["REPDATE"].shift(-1)
    df["next_st_name"] = g["ST_NAME"].shift(-1)
    df["next_rod_name"] = g["ROD_NAME"].shift(-1)
    df["cycle_days"] = ((df["next_load_dt"] - df["REPDATE"]).dt.total_seconds() / 86400).round(1)
    out = df.rename(columns={"REPDATE": "load_dt", "WAGNUM": "wagnum", "ST_CODE": "st_code", "ST_NAME": "st_name",
                             "RW_NAME": "rw_name", "RSV_ST_CODE": "rsv_st_code", "RSV_ST_NAME": "rsv_st_name",
                             "RSV_RW_NAME": "rsv_rw_name", "ROD_NAME": "rod_name", "GG_NAME": "gg_name",
                             "FR_CLASS": "fr_class", "C_OWNER_ID": "c_owner_id", "CLIENT_NAME": "client_name"})
    order = ["wagnum", "load_dt", "st_code", "st_name", "rw_name", "rsv_st_code", "rsv_st_name", "rsv_rw_name",
             "rod_name", "gg_name", "fr_class", "c_owner_id", "client_name", "next_load_dt", "next_st_name", "next_rod_name", "cycle_days"]
    out = out[order]
    (OUT / "p4-wagon-cycle" / "data").mkdir(parents=True, exist_ok=True)
    out.to_csv(OUT / "p4-wagon-cycle" / "data" / "wagon_cycles.csv.gz", index=False, compression="gzip",
               date_format="%Y-%m-%d %H:%M:%S")
    known = out.dropna(subset=["cycle_days"])
    assert (known.cycle_days >= 0).all()
    cd = known.cycle_days
    m = cd.mean()
    sst = ((cd - m) ** 2).sum()

    def r2_subset(sub, keys):
        m2 = sub.cycle_days.mean()
        sst2 = ((sub.cycle_days - m2) ** 2).sum()
        sse2 = sub.groupby(keys)["cycle_days"].apply(lambda x: ((x - x.mean()) ** 2).sum()).sum()
        return round(float(1 - sse2 / sst2), 3)

    def r2(keys):
        sse = known.groupby(keys)["cycle_days"].apply(lambda s: ((s - s.mean()) ** 2).sum()).sum()
        return round(float(1 - sse / sst), 3)

    at_dest = float((known.next_st_name == known.rsv_st_name).mean())
    top = (known.groupby(["rw_name", "rsv_rw_name"])["cycle_days"]
           .agg(["size", "median"]).sort_values("size", ascending=False).head(8).reset_index())
    stats = {
        "rows_total": int(len(out)), "rows_with_cycle": int(len(known)), "rows_open": int(len(out) - len(known)),
        "wagons": int(out.wagnum.nunique()), "dropped_duplicate_load_events": int(dropped),
        "cycle_days_median": float(cd.median()), "cycle_days_p10": float(cd.quantile(0.1)),
        "cycle_days_p90": float(cd.quantile(0.9)), "cycle_days_mean": round(float(m), 1),
        "share_next_load_at_destination": round(at_dest, 3),
        "share_next_load_same_rod": round(float((known.next_rod_name == known.rod_name).mean()), 3),
        "r2_by_station_pair": r2(["st_name", "rsv_st_name"]), "r2_by_road_pair": r2(["rw_name", "rsv_rw_name"]),
        "r2_by_road_pair_and_rod": r2(["rw_name", "rsv_rw_name", "rod_name"]),
        "same_rod_subset": {
            "rows": int((known.next_rod_name == known.rod_name).sum()),
            "cycle_days_median": float(known[known.next_rod_name == known.rod_name].cycle_days.median()),
            "r2_by_road_pair": r2_subset(known[known.next_rod_name == known.rod_name], ["rw_name", "rsv_rw_name"]),
            "r2_by_station_pair": r2_subset(known[known.next_rod_name == known.rod_name], ["st_name", "rsv_st_name"]),
        },
        "top_road_pairs": [{"from": a, "to": b, "n": int(n), "median_days": float(md)} for a, b, n, md in top.values],
        "period": [str(out.load_dt.min().date()), str(out.load_dt.max().date())],
    }
    if WITH_KEYS:
        KEYS.mkdir(parents=True, exist_ok=True)
        (KEYS / "p4-stats.json").write_text(json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8")
    return stats


# ---------------------------------------------------------------- unload -> daily series (вариант Б для П4)
def build_unload_daily() -> dict:
    cols = ["REPDATE", "ST_CODE", "ST_NAME", "RW_NAME", "ROD_NAME", "GG_NAME", "UNLOADED_WAGONS"]
    df = pd.read_csv(SRC / "V_DM_UNLOAD.csv", usecols=cols, dtype=str)
    df["repdate"] = pd.to_datetime(df["REPDATE"]).dt.date
    df["UNLOADED_WAGONS"] = df["UNLOADED_WAGONS"].astype(int)
    daily = (df.groupby(["repdate", "ST_CODE", "ST_NAME", "RW_NAME", "ROD_NAME", "GG_NAME"], as_index=False)["UNLOADED_WAGONS"]
             .sum().rename(columns={"ST_CODE": "st_code", "ST_NAME": "st_name", "RW_NAME": "rw_name",
                                    "ROD_NAME": "rod_name", "GG_NAME": "gg_name", "UNLOADED_WAGONS": "unloaded_wagons"}))
    daily = daily.sort_values(["repdate", "st_code", "rod_name", "gg_name"])
    daily.to_csv(OUT / "p4-wagon-cycle" / "data" / "unload_daily.csv.gz", index=False, compression="gzip")
    per_station_day = daily.groupby(["repdate", "st_code"])["unloaded_wagons"].sum()
    stats = {"rows": int(len(daily)), "days": int(daily.repdate.nunique()), "stations": int(daily.st_code.nunique()),
             "roads": int(daily.rw_name.nunique()), "wagons_total": int(daily.unloaded_wagons.sum()),
             "station_day_median": float(per_station_day.median()), "station_day_p90": float(per_station_day.quantile(0.9)),
             "period": [str(daily.repdate.min()), str(daily.repdate.max())]}
    if WITH_KEYS and (KEYS / "p4-stats.json").exists():
        st = json.loads((KEYS / "p4-stats.json").read_text(encoding="utf-8"))
        st["unload_daily"] = stats
        (KEYS / "p4-stats.json").write_text(json.dumps(st, ensure_ascii=False, indent=2), encoding="utf-8")
    return stats


# ---------------------------------------------------------------- П6: копия нужных витрин
def copy_p6_data() -> None:
    d = OUT / "p6-forecast-subscription" / "data"
    d.mkdir(parents=True, exist_ok=True)
    for f in ["V_DM_UNLOAD.csv", "V_DM_MPP.csv", "V_DM_UNLOAD_FORECAST.csv", "datamarts.student.json"]:
        shutil.copy2(SRC / f, d / f)


# ---------------------------------------------------------------- main
def main() -> None:
    st = build_stations()
    station_names = set(st.st_name)
    print("stations", len(st), "roads", st.rw_code.nunique())

    for d in [P1, *COPIES] + ([OUT / "p1-mail-orders"] if WITH_LETTERS else []):  # чистим старые реестры, чтобы не осталось лишних файлов
        for f in (d / "data").glob("registry_*.xlsx"):
            f.unlink()
    agg = load_gu12()
    N_REG = 20
    slices = pick_registry_slices(agg, N_REG)
    layouts = [LAYOUT_ORDER[i % len(LAYOUT_ORDER)] for i in range(N_REG)]
    registries = []
    for i, ((client, day), layout) in enumerate(zip(slices, layouts), 1):
        rows = agg[(agg.SND_ORG_NAME == client) & (agg.day == day)]
        registries.append(write_registry(i, layout, client, day, rows, station_names))
        print(registries[-1]["file"], layout, client, day, registries[-1]["clean_rows"], "rows")
    letters = write_letters(agg, set(slices), station_names) if WITH_LETTERS else []
    expected = {
        "source": "data/dataset_student/V_DM_GU12.csv, строки агрегированы по клиенту, дню, станциям, грузу и роду ПС",
        "seed": SEED,
        "date_semantics": "поле date — REPDATE снимка заявок ГУ-12 (день, на который в витрине учтены вагоны по заявкам)",
        "registries": registries, "letters": letters,
    }
    lines = ["| Файл | Раскладка | Клиент | Снимок | Чистых строк | Дополнительные дефекты |", "|---|---|---|---|---|---|"]
    for r in registries:
        lines.append(f"| `{r['file']}` | {r['layout']}: {LAYOUTS[r['layout']]['desc']} | {r['client']} | {r['snapshot_date']} | {r['clean_rows']} | {', '.join(r['extra_defects'])} |")
    # копии для П2 и П3: реестры, справочник, таблица файлов и эталон без писем
    exp_no_letters = dict(expected, letters=[])
    if WITH_KEYS:
        KEYS.mkdir(parents=True, exist_ok=True)
        (KEYS / "p2-p3-expected.json").write_text(json.dumps(expected, ensure_ascii=False, indent=2), encoding="utf-8")
        (KEYS / "p2-p3-FILES.md").write_text("# Реестры\n\n" + "\n".join(lines) + "\n", encoding="utf-8")
    copies = list(COPIES) + ([OUT / "p1-mail-orders"] if WITH_LETTERS else [])
    for d in copies:
        (d / "data").mkdir(parents=True, exist_ok=True)
        for f in list((P1 / "data").glob("registry_*.xlsx")) + [P1 / "data" / "stations.csv"]:
            shutil.copy2(f, d / "data" / f.name)

    stats = build_wagon_cycles()
    print(json.dumps(stats, ensure_ascii=False, indent=1)[:1200])
    print(json.dumps(build_unload_daily(), ensure_ascii=False))
    copy_p6_data()


if __name__ == "__main__":
    main()
