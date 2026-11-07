#!/usr/bin/env python3
"""Пакет синтетических документов демо-объекта для Полиметрики.

Вход — спецификация объекта (JSON, формат в `references/spec.md`). Выход — папка с документами,
которые помощник кабинета Полиметрики разбирает и применяет к проекту:

    upload/        что приложить к чату нового проекта
    telemetry/     почасовые показания приборов учёта и перечень приборов (загружаются в чат помощника)
    ЧТО ЗАЛОЖЕНО.md   ответы: какие потери и аномалии спрятаны в данных
    summary.json   сводка для Claude: итоги, доли, файлы

Все документы считаются из ОДНОЙ почасовой модели объекта: счета = сумма потребления оборудования,
выпуск управляет загрузкой процессного оборудования, показания приборов — те же часы. Поэтому данные
сходятся между собой, а заложенные аномалии находятся анализом. Генерация детерминирована: тот же
`seed` даёт тот же пакет.

Запуск:  python generate.py spec.json --out ./pack
Зависимости: openpyxl, python-docx, reportlab (есть в песочнице Claude); без reportlab счета
выйдут в DOCX.
"""

from __future__ import annotations

import argparse
import calendar
import csv
import json
import math
import random
import shutil
import sys
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

MONTHS = [
    "январь", "февраль", "март", "апрель", "май", "июнь",
    "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь",
]
PROFILES = {"process", "shift_system", "constant", "cooling", "heating", "ventilation", "lighting"}
ANOMALIES = {"off_shift", "unit_overconsumption", "no_night_setback", "drift"}


# --- Спецификация ---------------------------------------------------------------------------------


class SpecError(ValueError):
    pass


@dataclass
class Equipment:
    key: str
    name: str
    model: str
    count: int
    rated_kw: float
    profile: str
    load: float
    idle: float
    location: str
    circuit: str | None
    metered: bool
    in_inventory: bool


@dataclass
class Spec:
    raw: dict[str, Any]
    seed: int
    obj: dict[str, Any]
    supply: dict[str, Any]
    product: dict[str, Any]
    circuits: list[dict[str, Any]]
    equipment: list[Equipment]
    anomalies: list[dict[str, Any]]
    end: date
    bill_months: int
    telemetry_days: int
    climate: dict[str, float]
    call_notes: list[str] = field(default_factory=list)
    links: list[Link] = field(default_factory=list)


@dataclass(frozen=True)
class Link:
    """Связь энергомодели, которую технолог проговаривает в разговоре, а агент собирает при разборе.

    `source` — ключ оборудования или `output` (выпуск продукции), `target` — ключ оборудования.
    """

    source: str
    target: str
    kind: str
    method: str | None = None
    say: str | None = None


#: Выпуск продукции как источник связи: узел-драйвер, который агент заводит сам.
OUTPUT = "output"
LINK_KINDS = {"metric", "process", "unmeasured"}
#: Методы связи-измерения, которые агент считает по данным (docs#610): от выпуска, в окне процесса.
LINK_METHODS = {"regression", "window", "window_const"}


def _req(data: dict[str, Any], key: str, where: str) -> Any:
    if key not in data or data[key] in (None, ""):
        raise SpecError(f"{where}: нет поля «{key}»")
    return data[key]


def load_spec(path: Path) -> Spec:
    raw = json.loads(path.read_text(encoding="utf-8"))
    obj = _req(raw, "object", "spec")
    for key in ("name", "company", "address", "area_m2", "timezone", "assortment"):
        _req(obj, key, "object")
    shifts = _req(obj, "shifts", "object")
    for key in ("title", "start_hour", "end_hour"):
        _req(shifts, key, "object.shifts")
    supply = _req(raw, "supply", "spec")
    for key in ("supplier", "contract_number", "max_power_kw", "tariff_rub_per_kwh"):
        _req(supply, key, "supply")
    product = _req(raw, "product", "spec")
    for key in ("title", "unit", "daily_min", "daily_max"):
        _req(product, key, "product")
    if product["unit"] not in ("кг", "т", "шт"):
        raise SpecError("product.unit: только «кг», «т» или «шт»")
    circuits = raw.get("circuits") or []
    circuit_keys = {c["key"] for c in circuits}
    equipment: list[Equipment] = []
    for i, item in enumerate(_req(raw, "equipment", "spec")):
        where = f"equipment[{i}]"
        profile = _req(item, "profile", where)
        if profile not in PROFILES:
            raise SpecError(f"{where}.profile: одно из {sorted(PROFILES)}")
        circuit = item.get("circuit")
        if circuit and circuit not in circuit_keys:
            raise SpecError(f"{where}.circuit: нет контура «{circuit}»")
        equipment.append(
            Equipment(
                key=_req(item, "key", where),
                name=_req(item, "name", where),
                model=item.get("model", ""),
                count=int(item.get("count", 1)),
                rated_kw=float(_req(item, "rated_kw", where)),
                profile=profile,
                load=float(item.get("load", 0.7)),
                idle=float(item.get("idle", 0.0)),
                location=item.get("location", ""),
                circuit=circuit,
                metered=bool(item.get("metered", True)),
                in_inventory=bool(item.get("in_inventory", True)),
            )
        )
    keys = [e.key for e in equipment]
    if len(keys) != len(set(keys)):
        raise SpecError("equipment: ключи должны быть уникальны")
    anomalies = raw.get("anomalies") or []
    for i, anomaly in enumerate(anomalies):
        if anomaly.get("type") not in ANOMALIES:
            raise SpecError(f"anomalies[{i}].type: одно из {sorted(ANOMALIES)}")
        if anomaly.get("equipment") not in keys:
            raise SpecError(f"anomalies[{i}].equipment: нет оборудования «{anomaly.get('equipment')}»")
    links = _links(raw.get("links"), equipment, circuits)
    period = raw.get("period") or {}
    end = date.fromisoformat(period["end"]) if period.get("end") else _last_full_month_end(date.today())
    return Spec(
        raw=raw,
        seed=int(raw.get("seed", 1)),
        obj=obj,
        supply=supply,
        product=product,
        circuits=circuits,
        equipment=equipment,
        anomalies=anomalies,
        end=end,
        bill_months=int(period.get("bill_months", 12)),
        telemetry_days=int(period.get("telemetry_days", 31)),
        climate={"t_mean": 5.0, "t_amp": 15.0, "t_day_amp": 4.0, **(raw.get("climate") or {})},
        call_notes=list(raw.get("call_notes") or []),
        links=links,
    )


def _link_from(item: Any, where: str, keys: set[str]) -> Link:
    if not isinstance(item, dict):
        raise SpecError(f"{where}: ожидается объект")
    source, target, kind = _req(item, "from", where), _req(item, "to", where), _req(item, "kind", where)
    if source != OUTPUT and source not in keys:
        raise SpecError(f"{where}.from: нет оборудования «{source}» (или «{OUTPUT}» — выпуск)")
    if target not in keys or target == source:
        raise SpecError(f"{where}.to: нет оборудования «{target}»")
    if kind not in LINK_KINDS:
        raise SpecError(f"{where}.kind: одно из {sorted(LINK_KINDS)}")
    method = item.get("method")
    if kind == "metric" and method not in LINK_METHODS:
        raise SpecError(f"{where}.method: у связи-измерения одно из {sorted(LINK_METHODS)}")
    if kind != "metric" and method is not None:
        raise SpecError(f"{where}.method: метод только у связи-измерения (kind: metric)")
    if kind == "process" and source == OUTPUT:
        raise SpecError(f"{where}: технологическая связь — между оборудованием, не от выпуска")
    return Link(source, target, kind, method, item.get("say"))


def _links(raw: Any, equipment: list[Equipment], circuits: list[dict[str, Any]]) -> list[Link]:
    """Связи из спецификации; без поля `links` — по профилям и участкам (`default_links`)."""
    if raw is None:
        return default_links(equipment, circuits)
    if not isinstance(raw, list):
        raise SpecError("links: ожидается список")
    keys = {e.key for e in equipment}
    found = [_link_from(item, f"links[{i}]", keys) for i, item in enumerate(raw)]
    pairs = [(link.source, link.target) for link in found]
    if len(pairs) != len(set(pairs)):
        raise SpecError("links: связь между одной парой — один раз")
    return found


def default_links(equipment: list[Equipment], circuits: list[dict[str, Any]]) -> list[Link]:
    """Связи по участкам: технологическое оборудование — от выпуска, остальное — в окне процесса участка.

    Участки идут по порядку `circuits`: продукт с участка переходит на следующий. Вспомогательное
    оборудование вне участка (компрессоры, ККБ, освещение) связей не получает: их агент предложит сам по
    данным датчиков. Нужны другие связи — их задают полем `links`.
    """
    found: list[Link] = []
    previous: Equipment | None = None
    for circuit in circuits:
        members = [e for e in equipment if e.circuit == circuit["key"] and e.in_inventory]
        lead = next((e for e in members if e.profile == "process" and e.metered), None)
        if lead is None:
            continue
        found.append(Link(OUTPUT, lead.key, "metric", "regression"))
        if previous is not None:
            found.append(Link(previous.key, lead.key, "process"))
        for e in members:
            if e is lead:
                continue
            if not e.metered:
                found.append(Link(lead.key, e.key, "unmeasured"))
            elif e.profile == "process":
                found.append(Link(OUTPUT, e.key, "metric", "regression"))
            else:
                found.append(Link(lead.key, e.key, "metric", "window"))
        previous = lead
    return found


def _last_full_month_end(today: date) -> date:
    return today.replace(day=1) - timedelta(days=1)


# --- Почасовая модель -----------------------------------------------------------------------------


@dataclass
class Hour:
    ts: datetime
    is_work: bool
    t_out: float
    day_factor: float


@dataclass
class Simulation:
    hours: list[Hour]
    #: Мощность по единицам оборудования, кВт: ключ «<key>#<n>» → значение на каждый час.
    units: dict[str, list[float]]
    #: Неучтённая база (дежурное освещение, IT, потери), кВт на каждый час.
    other: list[float]
    production: dict[date, float]
    #: Та же мощность без заложенных аномалий — для оценки потерь в «ЧТО ЗАЛОЖЕНО.md».
    baseline: dict[str, list[float]] = field(default_factory=dict)


def _start(spec: Spec) -> date:
    month, year = spec.end.month, spec.end.year
    for _ in range(spec.bill_months - 1):
        month, year = (12, year - 1) if month == 1 else (month - 1, year)
    return date(year, month, 1)


def _work_days(spec: Spec) -> set[int]:
    return set(spec.obj["shifts"].get("weekdays", [0, 1, 2, 3, 4]))


def _is_work_hour(spec: Spec, ts: datetime, workdays: set[int]) -> bool:
    shifts = spec.obj["shifts"]
    start, end = int(shifts["start_hour"]), int(shifts["end_hour"])
    if ts.weekday() not in workdays:
        return False
    if start < end:
        return start <= ts.hour < end
    return ts.hour >= start or ts.hour < end  # смена через полночь


def _outdoor(spec: Spec, ts: datetime, rng: random.Random) -> float:
    c = spec.climate
    # Минимум года — середина января, минимум суток — 5 утра.
    season = -math.cos(2 * math.pi * (ts.timetuple().tm_yday - 15) / 365.25)
    daily = -math.cos(2 * math.pi * (ts.hour - 5) / 24)
    return c["t_mean"] + c["t_amp"] * season + c["t_day_amp"] * daily + rng.gauss(0, 1.5)


def _anomaly(spec: Spec, key: str, kind: str) -> dict[str, Any] | None:
    return next((a for a in spec.anomalies if a["equipment"] == key and a["type"] == kind), None)


def _unit_fraction(e: Equipment, hour: Hour, rng: random.Random) -> float:
    """Доля номинала единицы оборудования в этот час по профилю, без аномалий."""
    t = hour.t_out
    if e.profile == "process":
        value = e.load * (0.55 + 0.45 * hour.day_factor) if hour.is_work and hour.day_factor > 0 else e.idle
    elif e.profile == "shift_system":
        value = e.load * (0.7 + 0.3 * hour.day_factor) if hour.is_work else e.idle
    elif e.profile == "constant":
        value = e.load
    elif e.profile == "cooling":
        demand = min(1.0, max(0.0, (t - 10) / 18))
        value = e.load * (0.25 + 0.75 * demand) if hour.is_work else e.load * 0.35 * demand
    elif e.profile == "heating":
        demand = min(1.0, max(0.0, (16 - t) / 32))
        value = e.load * demand * (1.0 if hour.is_work else 0.7)
    elif e.profile == "ventilation":
        value = e.load if hour.is_work else e.idle
    else:  # lighting
        value = e.load if hour.is_work else max(e.idle, 0.05)
    return max(0.0, value * (1 + rng.gauss(0, 0.05)))


def simulate(spec: Spec) -> Simulation:
    rng = random.Random(spec.seed)
    workdays = _work_days(spec)
    start = _start(spec)
    production: dict[date, float] = {}
    lo, hi = float(spec.product["daily_min"]), float(spec.product["daily_max"])
    day = start
    while day <= spec.end:
        production[day] = round(rng.uniform(lo, hi)) if day.weekday() in workdays else 0.0
        day += timedelta(days=1)

    hours: list[Hour] = []
    ts = datetime(start.year, start.month, start.day)
    stop = datetime(spec.end.year, spec.end.month, spec.end.day) + timedelta(days=1)
    while ts < stop:
        output = production[ts.date()]
        factor = (output - lo) / (hi - lo) if output and hi > lo else (1.0 if output else 0.0)
        hours.append(Hour(ts, _is_work_hour(spec, ts, workdays), _outdoor(spec, ts, rng), factor))
        ts += timedelta(hours=1)

    drift_from = len(hours) - 24 * 7 * 8  # дрейф — последние 8 недель
    units: dict[str, list[float]] = {}
    baseline: dict[str, list[float]] = {}
    anomaly_rng = random.Random(spec.seed + 2)
    for e in spec.equipment:
        off_shift = _anomaly(spec, e.key, "off_shift")
        no_setback = _anomaly(spec, e.key, "no_night_setback")
        over = _anomaly(spec, e.key, "unit_overconsumption")
        drift = _anomaly(spec, e.key, "drift")
        for n in range(1, e.count + 1):
            series: list[float] = []
            clean: list[float] = []
            for i, hour in enumerate(hours):
                fraction = _unit_fraction(e, hour, rng)
                clean.append(round(e.rated_kw * fraction, 3))
                noise = 1 + anomaly_rng.gauss(0, 0.05)
                if not hour.is_work and off_shift:
                    fraction = max(fraction, float(off_shift.get("share", 0.4)) * noise)
                if not hour.is_work and no_setback:
                    fraction = e.load * noise
                kw = e.rated_kw * fraction
                if over and int(over.get("unit", 1)) == n:
                    kw *= 1 + float(over.get("extra", 0.3))
                if drift and int(drift.get("unit", 1)) == n and i >= drift_from:
                    kw *= 1 + float(drift.get("extra", 0.2)) * (i - drift_from) / (len(hours) - drift_from)
                series.append(round(kw, 3))
            units[f"{e.key}#{n}"] = series
            baseline[f"{e.key}#{n}"] = clean

    rated_total = sum(e.rated_kw * e.count for e in spec.equipment)
    base = float(spec.raw.get("other_load_kw", max(2.0, rated_total * 0.03)))
    other = []
    for i, hour in enumerate(hours):
        total = sum(series[i] for series in units.values())
        other.append(round(base * (1 + rng.gauss(0, 0.1)) + total * 0.02, 3))  # база + потери в сети
    return Simulation(hours, units, other, production, baseline)


# --- Итоги по месяцам -----------------------------------------------------------------------------


@dataclass
class Bill:
    year: int
    month: int
    kwh: float
    tariff: float
    vat: float

    @property
    def net(self) -> float:
        return round(self.kwh * self.tariff, 2)

    @property
    def gross(self) -> float:
        return round(self.net * (1 + self.vat), 2)


def monthly(sim: Simulation) -> dict[tuple[int, int], float]:
    totals: dict[tuple[int, int], float] = {}
    for i, hour in enumerate(sim.hours):
        key = (hour.ts.year, hour.ts.month)
        totals[key] = totals.get(key, 0.0) + sum(s[i] for s in sim.units.values()) + sim.other[i]
    return totals


def bills(spec: Spec, sim: Simulation) -> list[Bill]:
    current = float(spec.supply["tariff_rub_per_kwh"])
    result = []
    end_year = spec.end.year
    for (year, month), kwh in sorted(monthly(sim).items()):
        # Индексация 1 июля: до неё тариф ниже на ~10 %; плюс сезонный шум свободной цены.
        indexed = (year, month) >= ((end_year, 7) if spec.end.month >= 7 else (end_year - 1, 7))
        tariff = current * (1.0 if indexed else 0.91) * (1 + 0.03 * math.sin(month))
        vat = 0.22 if year >= 2026 else 0.20
        result.append(Bill(year, month, round(kwh), round(tariff, 4), vat))
    return result


# --- Документы ------------------------------------------------------------------------------------


def _fmt(value: float, digits: int = 0) -> str:
    text = f"{value:,.{digits}f}".replace(",", " ").replace(".", ",")
    return text


def _docx():
    import docx  # noqa: PLC0415
    from docx.shared import Pt  # noqa: PLC0415

    document = docx.Document()
    style = document.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    return document


def _docx_table(document, header: list[str], rows: list[list[str]]) -> None:
    table = document.add_table(rows=1, cols=len(header))
    table.style = "Table Grid"
    for cell, text in zip(table.rows[0].cells, header, strict=True):
        cell.text = text
    for row in rows:
        cells = table.add_row().cells
        for cell, text in zip(cells, row, strict=True):
            cell.text = text


def write_object_description(spec: Spec, path: Path) -> None:
    o, p = spec.obj, spec.product
    d = _docx()
    d.add_heading("Описание объекта", level=1)
    d.add_paragraph(f"{o['name']} ({o['company']})")
    d.add_heading("1. Общие данные", level=2)
    rows = [
        ["Наименование объекта", o["name"]],
        ["Адрес объекта", o["address"]],
        ["Площадь цеха, м²", _fmt(float(o["area_m2"]))],
        ["Часовой пояс", o["timezone"]],
        ["Климатическая зона (СП 131.13330)", o.get("climate_zone", "—")],
        ["Режим работы", o["shifts"]["title"]],
        ["Ассортимент", o["assortment"]],
        ["Единица выпуска", f"{p['title']}, {p['unit']}"],
        ["Средний выпуск в рабочие сутки", f"{_fmt((p['daily_min'] + p['daily_max']) / 2)} {p['unit']}"],
        ["Тариф на электроэнергию без НДС, руб./кВт·ч", _fmt(float(spec.supply["tariff_rub_per_kwh"]), 2)],
    ]
    _docx_table(d, ["Показатель", "Значение"], rows)
    if spec.circuits:
        d.add_heading("2. Контуры энергомодели", level=2)
        titles = ", ".join(c["title"] for c in spec.circuits)
        d.add_paragraph(f"Контуры энергомодели объекта: {titles}. Других контуров нет.")
        circuit_rows = []
        for c in spec.circuits:
            units = ", ".join(e.name for e in spec.equipment if e.circuit == c["key"])
            circuit_rows.append([c["title"], c["key"], c.get("driver_unit", p["unit"]), units or "—"])
        _docx_table(d, ["Контур", "Ключ", "Единица выпуска", "Основное оборудование"], circuit_rows)
    d.add_heading("3. Инженерные системы", level=2)
    systems = [e for e in spec.equipment if e.profile not in ("process",)]
    for e in systems:
        d.add_paragraph(f"{e.name}: {e.count} шт., {e.location or 'по месту'}.", style="List Bullet")
    if o.get("description"):
        d.add_heading("4. Особенности", level=2)
        d.add_paragraph(o["description"])
    d.save(path)


def write_contract(spec: Spec, path: Path) -> None:
    s, o = spec.supply, spec.obj
    points = s.get("supply_points") or ["ВРУ-0,4 кВ цеха"]
    d = _docx()
    d.add_heading(f"Договор энергоснабжения № {s['contract_number']}", level=1)
    d.add_paragraph(
        f"{s['supplier']}, именуемое в дальнейшем «Поставщик», и {o['company']}, "
        "именуемое в дальнейшем «Потребитель», заключили настоящий договор о нижеследующем."
    )
    d.add_heading("1. Предмет договора", level=2)
    d.add_paragraph(
        "Поставщик осуществляет продажу электрической энергии (мощности), а Потребитель "
        f"оплачивает её на условиях договора. Объект: {o['name']}, {o['address']}."
    )
    d.add_heading("2. Существенные условия", level=2)
    rows = [
        ["Максимальная мощность", f"{_fmt(float(s['max_power_kw']))} кВт"],
        ["Ценовая категория", str(s.get("tariff_category", 1))],
        ["Число точек поставки (вводов)", str(len(points))],
        ["Точки поставки", "; ".join(points)],
        # «кВт·ч» рядом с «…энергосбыт» в названии поставщика агент принимает за счёт — пишем словами.
        ["Средний тариф без НДС", f"{_fmt(float(s['tariff_rub_per_kwh']), 2)} руб. за 1 киловатт-час"],
    ]
    _docx_table(d, ["Условие", "Значение"], rows)
    d.add_heading("3. Срок действия договора", level=2)
    d.add_paragraph("Стороны договорились, что договор вступает в силу с момента подписания и действует "
                    "до 31 декабря текущего года с ежегодной пролонгацией.")
    d.add_heading("4. Реквизиты сторон", level=2)
    d.add_paragraph(f"Поставщик: {s['supplier']}. Потребитель: {o['company']}, {o['address']}.")
    d.save(path)


def write_inventory(spec: Spec, path: Path) -> None:
    from openpyxl import Workbook  # noqa: PLC0415
    from openpyxl.styles import Font  # noqa: PLC0415

    wb = Workbook()
    ws = wb.active
    ws.title = "Оборудование"
    ws.append(["Ведомость оборудования (перечень электроприёмников)"])
    ws.append([f"{spec.obj['name']}, {spec.obj['address']}"])
    ws.append([])
    # Без колонки «№ п/п»: в блоках document-service номер склеивается с наименованием («1 Ростер …»).
    ws.append(["Наименование", "Марка, модель", "Кол-во, шт", "Мощность, кВт", "Место установки"])
    for cell in ws[4]:
        cell.font = Font(bold=True)
    ws["A1"].font = Font(bold=True, size=13)
    for e in spec.equipment:
        if not e.in_inventory:
            continue
        ws.append([e.name, e.model, e.count, e.rated_kw, e.location])
    _fit_columns(ws, (38, 26, 14, 18, 28), numeric="CD")
    wb.save(path)


def _fit_columns(ws: Any, widths: tuple[int, ...], numeric: str = "") -> None:
    """Ширины колонок с запасом и числа по центру ячейки.

    Document-service печатает лист через LibreOffice и собирает таблицу из текстового слоя PDF. Узкая
    колонка обрезает текст («Компрессо»), а число, прижатое к правой границе, сливается с текстом
    соседней ячейки в одну строку слоя («18 Участок обжарки»): колонки мощности и места установки
    склеиваются, и количество с мощностью не доходят до узлов модели.
    """
    from openpyxl.styles import Alignment  # noqa: PLC0415
    from openpyxl.utils import get_column_letter  # noqa: PLC0415

    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for column in numeric:
        for cell in ws[column]:
            cell.alignment = Alignment(horizontal="center")


def _pdf_font() -> str | None:
    """Шрифт с кириллицей для reportlab: системный DejaVu или копия из matplotlib."""
    try:
        from reportlab.pdfbase import pdfmetrics  # noqa: PLC0415
        from reportlab.pdfbase.ttfonts import TTFont  # noqa: PLC0415
    except ImportError:
        return None
    candidates = [
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        Path("/usr/share/fonts/dejavu/DejaVuSans.ttf"),
        Path("/usr/share/fonts/TTF/DejaVuSans.ttf"),
        Path("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"),
        Path("/Library/Fonts/Arial Unicode.ttf"),
        Path("C:/Windows/Fonts/arial.ttf"),
    ]
    try:
        import matplotlib  # noqa: PLC0415

        candidates.append(Path(matplotlib.get_data_path()) / "fonts/ttf/DejaVuSans.ttf")
    except ImportError:
        pass
    for candidate in candidates:
        if candidate.exists():
            pdfmetrics.registerFont(TTFont("Cyr", str(candidate)))
            return "Cyr"
    return None


def _bill_lines(spec: Spec, bill: Bill, number: int) -> tuple[list[str], list[str], list[list[str]], str]:
    s, o = spec.supply, spec.obj
    last_day = calendar.monthrange(bill.year, bill.month)[1]
    issued = date(bill.year, bill.month, last_day) + timedelta(days=5)
    period = f"{MONTHS[bill.month - 1]} {bill.year}"
    head = [
        s["supplier"],
        "Гарантирующий поставщик электрической энергии",
        f"Счёт-фактура № {number} от {issued:%d.%m.%Y}",
        f"Расчётный период: {period}",
        f"Договор энергоснабжения № {s['contract_number']}",
        f"Покупатель: {o['company']}, {o['address']}",
    ]
    point = (s.get("supply_points") or ["ВРУ-0,4 кВ цеха"])[0]
    header = [
        "Точка поставки", "Расчётный период", "Объём, кВт·ч",
        "Тариф без НДС, руб./кВт·ч", "Сумма без НДС, руб.", "Сумма с НДС, руб.",
    ]
    rows = [[point, period, _fmt(bill.kwh), _fmt(bill.tariff, 4), _fmt(bill.net, 2), _fmt(bill.gross, 2)]]
    total = f"Всего к оплате: {_fmt(bill.gross, 2)} руб., в т. ч. НДС {int(bill.vat * 100)} %"
    return head, header, rows, total


def write_bill(spec: Spec, bill: Bill, number: int, path_pdf: Path, font: str | None) -> Path:
    head, header, rows, total = _bill_lines(spec, bill, number)
    if font is None:
        path = path_pdf.with_suffix(".docx")
        d = _docx()
        for line in head:
            d.add_paragraph(line)
        _docx_table(d, header, rows)
        d.add_paragraph(total)
        d.save(path)
        return path
    from reportlab.lib import colors  # noqa: PLC0415
    from reportlab.lib.pagesizes import A4, landscape  # noqa: PLC0415
    from reportlab.lib.styles import ParagraphStyle  # noqa: PLC0415
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle  # noqa: PLC0415

    style = ParagraphStyle("p", fontName=font, fontSize=10, leading=13)
    bold = ParagraphStyle("b", parent=style, fontSize=13, leading=17)
    cell = ParagraphStyle("c", parent=style, fontSize=9, leading=11)
    story: list[Any] = [Paragraph(head[0], bold), Paragraph(head[1], style), Spacer(1, 8)]
    story += [Paragraph(line, bold if i == 2 else style) for i, line in enumerate(head[2:], start=2)]
    story.append(Spacer(1, 10))
    table = Table([[Paragraph(h, cell) for h in header], *[[Paragraph(v, cell) for v in r] for r in rows]])
    table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story += [table, Spacer(1, 10), Paragraph(total, style)]
    SimpleDocTemplate(str(path_pdf), pagesize=landscape(A4)).build(story)
    return path_pdf


def write_production_year(spec: Spec, sim: Simulation, path: Path) -> None:
    from openpyxl import Workbook  # noqa: PLC0415
    from openpyxl.styles import Font  # noqa: PLC0415

    unit = spec.product["unit"]
    totals: dict[tuple[int, int], float] = {}
    for day, value in sim.production.items():
        totals[(day.year, day.month)] = totals.get((day.year, day.month), 0.0) + value
    wb = Workbook()
    ws = wb.active
    ws.title = "Выпуск"
    last = (spec.end.year, spec.end.month)
    totals.pop(last, None)
    ws.append(["Отчёт о выпуске продукции по месяцам"])
    ws.append([f"{spec.obj['name']}: {spec.product['title']}"])
    ws.append([])
    ws.append(["Месяц", f"Выпуск продукции, {unit}"])
    ws["A1"].font = Font(bold=True, size=13)
    for (year, month), value in sorted(totals.items()):
        ws.append([f"{MONTHS[month - 1]} {year}", round(value)])
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 24
    wb.save(path)


def write_production_daily(spec: Spec, sim: Simulation, path: Path) -> tuple[date, date]:
    from openpyxl import Workbook  # noqa: PLC0415
    from openpyxl.styles import Font  # noqa: PLC0415

    unit = spec.product["unit"]
    first = spec.end.replace(day=1)
    wb = Workbook()
    ws = wb.active
    ws.title = "Сутки"
    ws.append([f"Суточный отчёт о выпуске продукции за {MONTHS[first.month - 1]} {first.year}"])
    ws.append([f"{spec.obj['name']}: {spec.product['title']}"])
    ws.append([])
    ws.append(["Дата", f"Выпуск, {unit}"])
    ws["A1"].font = Font(bold=True, size=13)
    day = first
    while day <= spec.end:
        ws.append([day.strftime("%d.%m.%Y"), round(sim.production[day])])
        day += timedelta(days=1)
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 16
    wb.save(path)
    return first, spec.end


_ANOMALY_LINES = {
    "off_shift": "{name}: по-моему, их никто не выключает после смены — утром приходим, они уже работают.",
    "no_night_setback": "{name} гоняем круглосуточно, ночного режима у неё нет, так и не настроили.",
    "unit_overconsumption": "{name} №{unit} шумит сильнее остальных, на ТО говорили, что пора чистить.",
    "drift": "{name} №{unit} последние пару месяцев греется больше обычного.",
}


def _lower(name: str) -> str:
    return name[:1].lower() + name[1:]


def link_line(link: Link, by_key: dict[str, Equipment], product: dict[str, Any]) -> str:
    """Как технолог говорит о связи. Название из ведомости стоит в именительном падеже: склонять его
    генератор не умеет, а агенту нужно узнать узел."""
    if link.say:
        return link.say
    target = by_key[link.target].name
    if link.source == OUTPUT:
        return (f"{target}: потребление идёт за выпуском — больше {product['unit']} за день, больше и берёт, "
                "без выпуска простаивает.")
    source = _lower(by_key[link.source].name)
    if link.kind == "process":
        return f"Продукт идёт по цепочке: сначала {source}, потом {_lower(target)}."
    if link.kind == "unmeasured":
        return f"{target} включается, когда работает {source}; своего счётчика нет, сидит на общем."
    if link.method == "window_const":
        return (f"{target} включаем, только когда работает {source}; нагрузка ровная, "
                "от объёма почти не зависит.")
    return f"{target} включаем, только когда работает {source}; остановились — можно выключать."


def write_call_transcript(spec: Spec, path: Path) -> None:
    by_key = {e.key: e for e in spec.equipment}
    shifts = spec.obj["shifts"]["title"]
    lines: list[tuple[str, str]] = [
        ("Энергетик", "Давай пройдёмся по цеху: что работает и когда."),
        ("Технолог", f"Работаем так: {shifts}. Выпуск — {spec.product['title'].lower()}, "
                     f"в среднем {_fmt((spec.product['daily_min'] + spec.product['daily_max']) / 2)} "
                     f"{spec.product['unit']} за смену."),
    ]
    if spec.links:
        lines.append(("Энергетик", "Расскажи, что от чего зависит: что идёт за выпуском, что работает под какой участок."))
        lines += [("Технолог", link_line(link, by_key, spec.product)) for link in spec.links]
        lines.append(("Энергетик", "Понял, так и свяжем в модели."))
    for anomaly in spec.anomalies:
        e = by_key[anomaly["equipment"]]
        template = anomaly.get("hint") or _ANOMALY_LINES[anomaly["type"]]
        lines.append(("Технолог", template.format(name=e.name, unit=anomaly.get("unit", 1))))
        lines.append(("Энергетик", "Понял, отмечу, проверим по приборам."))
    for e in spec.equipment:
        if not e.in_inventory:
            lines.append(("Технолог", f"Ещё есть {e.name.lower()}, {_fmt(e.rated_kw)} кВт, {e.count} шт. — "
                                      "в ведомость её не вносили."))
    for note in spec.call_notes:
        lines.append(("Технолог", note))
    lines.append(("Энергетик", "Спасибо, внесу в кабинет."))
    d = _docx()
    d.add_heading("Расшифровка разговора с технологом", level=1)
    d.add_paragraph(f"Объект: {spec.obj['name']}. Обход цеха, запись разговора.")
    for who, text in lines:
        d.add_paragraph(f"{who}: {text}")
    d.save(path)


def write_telemetry(spec: Spec, sim: Simulation, folder: Path) -> list[dict[str, Any]]:
    """Почасовые показания приборов учёта за последние `telemetry_days` суток + перечень приборов."""
    folder.mkdir(parents=True, exist_ok=True)
    first = len(sim.hours) - 24 * spec.telemetry_days
    devices: list[dict[str, Any]] = []
    for e in spec.equipment:
        if not e.metered:
            continue
        for n in range(1, e.count + 1):
            devices.append({
                "device": f"m230-{e.key.replace('_', '-')}-{n}",
                "equipment": e.key,
                "unit": n,
                "title": f"{e.name} №{n}" if e.count > 1 else e.name,
                "series": sim.units[f"{e.key}#{n}"],
            })
    rng = random.Random(spec.seed + 1)
    with (folder / "показания_приборов_учёта.csv").open("w", encoding="utf-8", newline="") as fh:
        writer = csv.writer(fh, delimiter=";")
        writer.writerow(["Прибор учёта", "Дата и время", "Активная мощность, кВт", "Показание счётчика, кВт·ч"])
        for device in devices:
            meter = rng.uniform(10_000, 90_000) + sum(device["series"][:first])
            for i in range(first, len(sim.hours)):
                kw = device["series"][i]
                meter += kw
                writer.writerow([device["device"], sim.hours[i].ts.strftime("%d.%m.%Y %H:%M"),
                                 _fmt(kw, 2), _fmt(meter, 1)])
    with (folder / "температура_наружного_воздуха.csv").open("w", encoding="utf-8", newline="") as fh:
        writer = csv.writer(fh, delimiter=";")
        writer.writerow(["Датчик", "Дата и время", "Температура, °C"])
        for i in range(first, len(sim.hours)):
            writer.writerow(["street-t1", sim.hours[i].ts.strftime("%d.%m.%Y %H:%M"), _fmt(sim.hours[i].t_out, 1)])

    from openpyxl import Workbook  # noqa: PLC0415

    wb = Workbook()
    ws = wb.active
    ws.title = "Приборы"
    ws.append(["Перечень приборов учёта"])
    ws.append(["Прибор учёта", "Тип", "Оборудование", "Место установки"])
    by_key = {e.key: e for e in spec.equipment}
    for device in devices:
        ws.append([device["device"], "Меркурий 230", device["title"], by_key[device["equipment"]].location])
    ws.append(["street-t1", "Датчик температуры", "Наружный воздух", "Северная стена"])
    _fit_columns(ws, (24, 22, max(24, *(len(d["title"]) + 4 for d in devices)), 24))
    wb.save(folder / "перечень_приборов_учёта.xlsx")
    return [{k: v for k, v in d.items() if k != "series"} for d in devices]


LINK_KIND_TITLES = {"metric": "измерение", "process": "технологическая", "unmeasured": "без измерения"}
LINK_METHOD_TITLES = {"regression": "от выпуска", "window": "в окне процесса",
                      "window_const": "в окне, постоянная нагрузка"}


def _link_end(key: str, by_key: dict[str, Equipment]) -> str:
    return "Выпуск продукции" if key == OUTPUT else by_key[key].name


def write_answers(spec: Spec, sim: Simulation, bills_: list[Bill], path: Path) -> list[dict[str, Any]]:
    tariff = float(spec.supply["tariff_rub_per_kwh"])
    by_key = {e.key: e for e in spec.equipment}
    hours = len(sim.hours)
    findings = []
    for anomaly in spec.anomalies:
        e = by_key[anomaly["equipment"]]
        title = anomaly.get("title") or {
            "off_shift": f"{e.name} работает вне смены",
            "no_night_setback": f"{e.name} без ночного режима",
            "unit_overconsumption": f"{e.name} №{anomaly.get('unit', 1)} потребляет больше однотипных",
            "drift": f"{e.name} №{anomaly.get('unit', 1)}: растёт потребление",
        }[anomaly["type"]]
        # Вклад этой аномалии: доли однотипных аномалий в одном оборудовании не разделяются.
        delta = sum(
            sum(sim.units[k]) - sum(sim.baseline[k]) for k in sim.units if k.startswith(f"{e.key}#")
        )
        yearly = delta * 8760 / hours
        findings.append({"title": title, "type": anomaly["type"], "equipment": e.key,
                         "kwh_per_year": round(yearly), "rub_per_year": round(yearly * tariff)})
    total_kwh = sum(b.kwh for b in bills_)
    total_out = sum(sim.production.values())
    lines = [
        "# Что заложено в данные",
        "",
        "Это ответы к демо-пакету. Загрузите пакет в Полиметрику и сравните, что найдёт анализ.",
        "",
        f"- Объект: {spec.obj['name']}; период: {_start(spec):%d.%m.%Y} — {spec.end:%d.%m.%Y}.",
        f"- Потребление по счетам: {_fmt(total_kwh)} кВт·ч; выпуск: {_fmt(total_out)} {spec.product['unit']}.",
        f"- Удельный расход: {_fmt(total_kwh / total_out, 3) if total_out else '—'} кВт·ч/{spec.product['unit']}.",
        "",
        "## Заложенные потери",
        "",
        "| Что | кВт·ч в год | ₽ в год (без НДС) |",
        "|---|---:|---:|",
        *[f"| {f['title']} | {_fmt(f['kwh_per_year'])} | {_fmt(f['rub_per_year'])} |" for f in findings],
    ]
    if spec.links:
        lines += ["", "## Связи, которые проговорил технолог", "",
                  "Агент должен собрать их при разборе расшифровки в карточку «Оборудование и учёт».", "",
                  "| Откуда | Куда | Вид | Метод |", "|---|---|---|---|",
                  *[f"| {_link_end(link.source, by_key)} | {by_key[link.target].name} | {LINK_KIND_TITLES[link.kind]} "
                    f"| {LINK_METHOD_TITLES.get(link.method or '', '—')} |" for link in spec.links]]
    hidden = [e for e in spec.equipment if not e.in_inventory]
    if hidden:
        lines += ["", "## Не внесено в ведомость", "",
                  *[f"- {e.name}, {_fmt(e.rated_kw)} кВт × {e.count} — упоминается только в разговоре с технологом."
                    for e in hidden]]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return findings


HOW_TO = """# Как загрузить пакет в Полиметрику

1. Откройте app.polimetrika.ru → «Новый объект».
2. В чат помощника приложите ВСЕ файлы из папки `upload/` (можно одним сообщением) и напишите:
   «Разбери документы и заполни проект».
3. Дождитесь распознавания (1–5 минут) и проверьте карточки предложений: параметры объекта,
   энергомодель, оборудование, ряды по счетам, выпуск. Связи, которые технолог проговорил в разговоре,
   лежат в той же карточке «Оборудование и учёт», что и узлы, с цитатой из разговора. Нажмите «Подтвердить».
4. Спросите помощника о разговоре с технологом: «Что из разговора с технологом стоит внести в модель?»

5. Показания приборов: приложите в тот же чат `telemetry/показания_приборов_учёта.csv` вместе
   с `telemetry/перечень_приборов_учёта.xlsx` и напишите «Загрузи показания приборов учёта из файла».
   Подтвердите карточку «Загрузить показания» — почасовые данные уйдут в хранилище показаний за минуту.
   Проверить: «Сколько кВт·ч потребило <оборудование> за <месяц> по суткам?»

`telemetry/температура_наружного_воздуха.csv` пока не загружайте: он понадобится, когда расчёты начнут
учитывать погоду.

`ЧТО ЗАЛОЖЕНО.md` — ответы: какие потери спрятаны в данных. Откройте после того, как посмотрите анализ.
"""


def build(spec: Spec, out: Path) -> dict[str, Any]:
    if out.exists():
        shutil.rmtree(out)
    upload = out / "upload"
    bills_dir = upload / "счета"
    bills_dir.mkdir(parents=True)
    sim = simulate(spec)
    bills_ = bills(spec, sim)
    files: list[Path] = []

    files.append(upload / "01_Описание_объекта.docx")
    write_object_description(spec, files[-1])
    files.append(upload / "02_Договор_энергоснабжения.docx")
    write_contract(spec, files[-1])
    files.append(upload / "03_Ведомость_оборудования.xlsx")
    write_inventory(spec, files[-1])
    font = _pdf_font()
    for i, bill in enumerate(bills_, start=1):
        name = f"Счёт_{bill.year}-{bill.month:02d}.pdf"
        files.append(write_bill(spec, bill, 4000 + i * 17, bills_dir / name, font))
    files.append(upload / "05_Отчёт_о_выпуске_по_месяцам.xlsx")
    write_production_year(spec, sim, files[-1])
    files.append(upload / "06_Суточный_выпуск.xlsx")
    write_production_daily(spec, sim, files[-1])
    files.append(upload / "07_Разговор_с_технологом.docx")
    write_call_transcript(spec, files[-1])

    devices = write_telemetry(spec, sim, out / "telemetry")
    findings = write_answers(spec, sim, bills_, out / "ЧТО ЗАЛОЖЕНО.md")
    (out / "КАК ЗАГРУЗИТЬ.md").write_text(HOW_TO, encoding="utf-8")

    by_equipment: dict[str, float] = {}
    for key, series in sim.units.items():
        by_equipment[key.split("#")[0]] = by_equipment.get(key.split("#")[0], 0.0) + sum(series)
    total = sum(by_equipment.values()) + sum(sim.other)
    summary = {
        "object": spec.obj["name"],
        "period": [_start(spec).isoformat(), spec.end.isoformat()],
        "bills": [{"month": f"{b.year}-{b.month:02d}", "kwh": b.kwh, "gross_rub": b.gross} for b in bills_],
        "total_kwh": round(total),
        "share_by_equipment": {k: round(v / total, 3) for k, v in sorted(by_equipment.items(), key=lambda x: -x[1])},
        "production_total": round(sum(sim.production.values())),
        "findings": findings,
        "devices": devices,
        "pdf_font": font is not None,
        "upload_files": [str(p.relative_to(out)) for p in files],
    }
    (out / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("spec", type=Path, help="спецификация объекта, JSON")
    parser.add_argument("--out", type=Path, default=Path("polimetrika-pack"), help="папка пакета")
    args = parser.parse_args()
    try:
        spec = load_spec(args.spec)
    except (SpecError, KeyError, ValueError) as exc:
        print(f"Ошибка спецификации: {exc}", file=sys.stderr)
        return 2
    summary = build(spec, args.out)
    print(json.dumps({k: summary[k] for k in ("object", "period", "total_kwh", "production_total",
                                                "findings", "pdf_font", "upload_files")},
                     ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
