#!/usr/bin/env python3
"""Пакет синтетических документов демо-объекта для Полиметрики.

Вход — спецификация объекта (JSON, формат в `references/spec.md`). Выход — папка с документами,
которые помощник кабинета Полиметрики разбирает и применяет к проекту:

    upload/        что приложить к чату нового проекта
    <папка>-upload.zip   та же upload/ одним архивом (рядом с папкой пакета)
    telemetry/     почасовые показания приборов учёта (электроэнергия; газ, тепло, вода на вводах) и перечень
                   приборов — загружаются в чат помощника
    ЧТО ЗАЛОЖЕНО.md   ответы: какие потери и аномалии спрятаны в данных
    summary.json   сводка для Claude: итоги, доли, файлы

Все документы считаются из ОДНОЙ почасовой модели объекта: счета = сумма потребления оборудования,
счета за газ, тепло и воду — из тех же часов, выпуска и улицы (поле ),
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
import re
import shutil
import sys
import zipfile
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
#: Аномалии покупных ресурсов: у них вместо `equipment` поле `resource`.
RESOURCE_ANOMALIES = {"water_leak": {"water"}, "heating_no_setback": {"gas", "heat"}}
RESOURCES = ("gas", "heat", "water", "liquid_fuel")


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
    #: Покупные ресурсы кроме электроэнергии: газ, тепло, вода, жидкое топливо (`references/spec.md`).
    resources: dict[str, dict[str, Any]] = field(default_factory=dict)
    #: Участки со своим выпуском и графиком — только при двух контурах и больше (`areas_of`).
    areas: list[Area] = field(default_factory=list)


@dataclass(frozen=True)
class Area:
    """Участок — контур энергомодели со своим выпуском и графиком работы.

    Заводится только при двух контурах и больше: объект без контуров или с одним работает и выпускает
    по `object.shifts` и `product`, как раньше. Без своих `daily_min`/`daily_max` выпуск участка — равная
    доля выпуска объекта, без своих `shifts` — график объекта.
    """

    key: str
    title: str
    unit: str
    low: float
    high: float
    #: `start_hour`, `end_hour`, `weekdays` (по умолчанию — рабочие дни объекта), `title`.
    shifts: dict[str, Any]
    #: Выпуск задан у контура; иначе — доля выпуска объекта.
    own: bool


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
    circuits = _circuits(raw.get("circuits") or [])
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
    resources = _resources(raw.get("resources"))
    _resource_circuits(resources, circuit_keys)
    anomalies = raw.get("anomalies") or []
    for i, anomaly in enumerate(anomalies):
        kind = anomaly.get("type")
        if kind in RESOURCE_ANOMALIES:
            if anomaly.get("resource") not in RESOURCE_ANOMALIES[kind] & resources.keys():
                raise SpecError(f"anomalies[{i}].resource: у «{kind}» ресурс из {sorted(RESOURCE_ANOMALIES[kind])}, "
                                "описанный в `resources`")
            continue
        if kind not in ANOMALIES:
            raise SpecError(f"anomalies[{i}].type: одно из {sorted(ANOMALIES | RESOURCE_ANOMALIES.keys())}")
        if anomaly.get("equipment") not in keys:
            raise SpecError(f"anomalies[{i}].equipment: нет оборудования «{anomaly.get('equipment')}»")
    links = _links(raw.get("links"), equipment, circuits)
    period = raw.get("period") or {}
    end = date.fromisoformat(period["end"]) if period.get("end") else _last_full_month_end(date.today())
    spec = Spec(
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
        resources=resources,
    )
    spec.areas = areas_of(spec)
    return spec


def _circuits(raw: Any) -> list[dict[str, Any]]:
    """Контуры: ключ и название обязательны, свой выпуск и график — только при двух контурах и больше."""
    if not isinstance(raw, list):
        raise SpecError("circuits: ожидается список")
    for i, circuit in enumerate(raw):
        where = f"circuits[{i}]"
        _req(circuit, "key", where)
        _req(circuit, "title", where)
        own = [name for name in ("daily_min", "daily_max", "shifts") if name in circuit]
        if own and len(raw) < 2:
            raise SpecError(f"{where}: свой выпуск и график — у двух контуров и больше; у единственного "
                            "контура они — объекта (`product`, `object.shifts`)")
        if ("daily_min" in circuit) != ("daily_max" in circuit):
            raise SpecError(f"{where}: выпуск участка — оба поля, daily_min и daily_max")
        if "daily_min" in circuit and not 0 < float(circuit["daily_min"]) <= float(circuit["daily_max"]):
            raise SpecError(f"{where}: нужно 0 < daily_min ≤ daily_max")
        if "shifts" in circuit:
            for key in ("title", "start_hour", "end_hour"):
                _req(circuit["shifts"], key, f"{where}.shifts")
    names = [_area_name(c["title"]) for c in raw]
    if len(names) != len(set(names)):
        raise SpecError("circuits: названия контуров должны различаться — по ним Полиметрика разносит выпуск")
    return raw


def _area_name(text: str) -> str:
    """Название так, как его сравнивает Полиметрика: только буквы и цифры, «ё» — «е»."""
    return re.sub(r"[^a-zа-я0-9]+", "", str(text).lower().replace("ё", "е"))


def _resource_circuits(resources: dict[str, dict[str, Any]], circuit_keys: set[str]) -> None:
    for key, item in resources.items():
        if "circuit" not in item:
            continue
        if key not in ("gas", "water"):
            raise SpecError(f"resources.{key}.circuit: контур — только у газа и воды")
        if item["circuit"] not in circuit_keys:
            raise SpecError(f"resources.{key}.circuit: нет контура «{item['circuit']}»")


def areas_of(spec: Spec) -> list[Area]:
    """Участки объекта с двумя контурами и больше; у объекта без контуров или с одним — пусто."""
    if len(spec.circuits) < 2:
        return []
    share = len(spec.circuits)
    low, high = float(spec.product["daily_min"]), float(spec.product["daily_max"])
    weekdays = sorted(_work_days(spec))
    found = []
    for c in spec.circuits:
        own = "daily_min" in c
        found.append(Area(
            key=c["key"],
            title=c["title"],
            unit=c.get("driver_unit") or spec.product["unit"],
            low=float(c["daily_min"]) if own else low / share,
            high=float(c["daily_max"]) if own else high / share,
            shifts={"weekdays": weekdays, **(c.get("shifts") or spec.obj["shifts"])},
            own=own,
        ))
    return found


#: Обязательные поля ресурса: поставщик, договор и тариф без НДС в единицах счёта.
_RESOURCE_FIELDS = {
    "gas": ("supplier", "contract_number", "tariff_rub_per_1000m3"),
    "heat": ("supplier", "contract_number", "tariff_rub_per_gcal", "heating_peak_gcal_h"),
    "water": ("supplier", "contract_number", "tariff_rub_per_m3", "wastewater_tariff_rub_per_m3"),
    "liquid_fuel": ("supplier", "price_rub_per_l", "l_per_work_day"),
}


#: Слова, по которым агент узнаёт суточный архив корректора газа (сильный признак типа документа). В марке
#: прибора они уводят ведомость оборудования из инвентаризации: разрыв баллов падает, тип решает модель.
METER_ARCHIVE_WORDS = re.compile(r"ирвис|корректор|архив")


def _resources(raw: Any) -> dict[str, dict[str, Any]]:
    if raw is None:
        return {}
    if not isinstance(raw, dict):
        raise SpecError("resources: ожидается объект {ресурс: описание}")
    for key, item in raw.items():
        if key not in _RESOURCE_FIELDS:
            raise SpecError(f"resources.{key}: ресурс из {list(RESOURCES)}")
        for name in _RESOURCE_FIELDS[key]:
            _req(item, name, f"resources.{key}")
    gas = raw.get("gas")
    if gas and not (gas.get("process_m3_per_unit") or gas.get("heating_peak_m3_h")):
        raise SpecError("resources.gas: нужен расход на выпуск (process_m3_per_unit) или котельная (heating_peak_m3_h)")
    if gas and gas.get("heating_peak_m3_h") and "heat" in raw:
        raise SpecError("resources: отопление либо своей котельной на газе (gas.heating_peak_m3_h), "
                        "либо от сети (heat)")
    for key, item in raw.items():
        if METER_ARCHIVE_WORDS.search(str(item.get("meter", "")).lower()):
            raise SpecError(f"resources.{key}.meter: марка без «Ирвис», «корректор», «архив» — по ним ведомость "
                            "распознаётся как архив корректора газа; например «СГ-16МТ-160»")
    fuel = raw.get("liquid_fuel")
    if fuel and fuel.get("fuel_kind", "diesel") not in FUEL_KINDS:
        raise SpecError(f"resources.liquid_fuel.fuel_kind: одно из {sorted(FUEL_KINDS)}")
    return {key: raw[key] for key in RESOURCES if key in raw}


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
    #: Выпуск участков по суткам (`Spec.areas`): ключ контура → сутки → выпуск в единице контура.
    outputs: dict[str, dict[date, float]] = field(default_factory=dict)
    #: Часы участков: признак работы по графику участка и загрузка по его выпуску; ключ — контур.
    area_hours: dict[str, list[Hour]] = field(default_factory=dict)


def _start(spec: Spec) -> date:
    month, year = spec.end.month, spec.end.year
    for _ in range(spec.bill_months - 1):
        month, year = (12, year - 1) if month == 1 else (month - 1, year)
    return date(year, month, 1)


def _work_days(spec: Spec) -> set[int]:
    return set(spec.obj["shifts"].get("weekdays", [0, 1, 2, 3, 4]))


def _is_work_hour(spec: Spec, ts: datetime, workdays: set[int]) -> bool:
    return _in_shift(spec.obj["shifts"], ts, workdays)


def _in_shift(shifts: dict[str, Any], ts: datetime, workdays: set[int]) -> bool:
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
    return next((a for a in spec.anomalies if a.get("equipment") == key and a["type"] == kind), None)


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
        factor = _day_factor(production[ts.date()], lo, hi)
        hours.append(Hour(ts, _is_work_hour(spec, ts, workdays), _outdoor(spec, ts, rng), factor))
        ts += timedelta(hours=1)
    # Участки — своими генераторами случайных чисел: общий ряд чисел, а с ним улица и общие системы, не сдвигается.
    outputs = area_outputs(spec, production)
    by_area = {area.key: area_hours(area, hours, outputs[area.key]) for area in spec.areas}

    drift_from = len(hours) - 24 * 7 * 8  # дрейф — последние 8 недель
    units: dict[str, list[float]] = {}
    baseline: dict[str, list[float]] = {}
    anomaly_rng = random.Random(spec.seed + 2)
    for e in spec.equipment:
        off_shift = _anomaly(spec, e.key, "off_shift")
        no_setback = _anomaly(spec, e.key, "no_night_setback")
        over = _anomaly(spec, e.key, "unit_overconsumption")
        drift = _anomaly(spec, e.key, "drift")
        # Оборудование участка работает по часам и выпуску участка, «вне смены» — вне окна участка.
        own_hours = by_area.get(e.circuit or "", hours)
        for n in range(1, e.count + 1):
            series: list[float] = []
            clean: list[float] = []
            for i, hour in enumerate(own_hours):
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
    return Simulation(hours, units, other, production, baseline, outputs=outputs, area_hours=by_area)


def _day_factor(output: float, low: float, high: float) -> float:
    """Загрузка суток по выпуску: 0 — минимум диапазона, 1 — максимум; без выпуска — 0."""
    return (output - low) / (high - low) if output and high > low else (1.0 if output else 0.0)


def area_outputs(spec: Spec, production: dict[date, float]) -> dict[str, dict[date, float]]:
    """Суточный выпуск участков: свой диапазон в рабочие дни участка или равная доля выпуска объекта."""
    found: dict[str, dict[date, float]] = {}
    for index, area in enumerate(spec.areas):
        rng = random.Random(spec.seed + 40 + index)
        workdays = set(area.shifts["weekdays"])
        days: dict[date, float] = {}
        for day, total in production.items():
            if day.weekday() not in workdays:
                days[day] = 0.0
            elif area.own:
                days[day] = float(round(rng.uniform(area.low, area.high)))
            else:
                days[day] = float(round(total / len(spec.areas)))
        found[area.key] = days
    return found


def area_hours(area: Area, hours: list[Hour], output: dict[date, float]) -> list[Hour]:
    """Часы участка: работа — по его графику, загрузка — по его выпуску; улица общая."""
    workdays = set(area.shifts["weekdays"])
    return [Hour(h.ts, _in_shift(area.shifts, h.ts, workdays), h.t_out,
                 _day_factor(output[h.ts.date()], area.low, area.high)) for h in hours]


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


def _indexed(spec: Spec, year: int, month: int) -> bool:
    """Регулируемые тарифы индексируют 1 июля: до этой даты в году счёта тариф ниже."""
    return (year, month) >= ((spec.end.year, 7) if spec.end.month >= 7 else (spec.end.year - 1, 7))


def _vat(year: int) -> float:
    return 0.22 if year >= 2026 else 0.20


# --- Покупные ресурсы -----------------------------------------------------------------------------
#
# Газ, тепло и вода считаются по тем же часам, выпуску и улице, что и электроэнергия, но своим генератором
# случайных чисел и после `simulate`: электрическая часть пакета от них не меняется ни в одной цифре.

#: Подписи ресурсов в документах и ответах; единица — та, в которой ведётся учёт.
RESOURCE_TITLES = {"gas": ("Природный газ", "м³"), "heat": ("Тепловая энергия", "Гкал"),
                   "water": ("Холодная вода", "м³"), "liquid_fuel": ("Жидкое топливо", "л")}
#: Топливо: название в накладной, плотность кг/л, т CO₂ на тонну (приказ Минприроды № 371), слова для шапки.
FUEL_KINDS = {"diesel": ("Топливо дизельное ЕВРО, межсезонное, экологического класса К5", 0.84, 3.15,
                         "дизельное топливо (ГСМ)"),
              "fuel_oil": ("Мазут топочный М-100", 0.97, 3.11, "мазут топочный")}
#: т CO₂ на тысячу м³ природного газа (решение владельца 11.11).
GAS_CO2_T_PER_1000M3 = 1.80
#: Ночное снижение отопления вне смены: доля дневного расхода.
HEATING_SETBACK = 0.6
#: Отопительный сезон: отопление включено в сутки со средней температурой улицы ниже этой, °C. Летом
#: счёт за отопление нулевой, как у ИТП и котельной по графику сезона.
HEATING_SEASON_C = 10.0
#: Ночная база водомера без утечки, м³/ч: капельный расход, ниже порога детектора утечки.
WATER_NIGHT_BASE_M3_H = 0.004


@dataclass
class ResourceSeries:
    """Часовой расход ресурса (м³ или Гкал) с заложенными аномалиями и без них."""

    hourly: list[float]
    baseline: list[float]


@dataclass
class ResourceBill:
    resource: str
    year: int
    month: int
    #: Объём в единице учёта: м³, Гкал, л. У воды — холодная вода, водоотведение отдельно.
    volume: float
    tariff: float
    vat: float
    wastewater: float = 0.0
    wastewater_tariff: float = 0.0

    @property
    def net(self) -> float:
        per_unit = self.volume / 1000 if self.resource == "gas" else self.volume  # газ — за тысячу м³
        return round(per_unit * self.tariff, 2) + round(self.wastewater * self.wastewater_tariff, 2)

    @property
    def gross(self) -> float:
        return round(self.net * (1 + self.vat), 2)


def _heating(t_out: float) -> float:
    """Потребность в отоплении, доля пиковой: та же кривая, что у профиля `heating` электроэнергии."""
    return min(1.0, max(0.0, (16 - t_out) / 32))


def _heating_days(sim: Simulation) -> set[date]:
    temps: dict[date, list[float]] = {}
    for hour in sim.hours:
        temps.setdefault(hour.ts.date(), []).append(hour.t_out)
    return {day for day, values in temps.items() if sum(values) / len(values) < HEATING_SEASON_C}


def _work_hours(hours: list[Hour]) -> dict[date, int]:
    counts: dict[date, int] = {}
    for hour in hours:
        counts[hour.ts.date()] = counts.get(hour.ts.date(), 0) + (1 if hour.is_work else 0)
    return counts


#: Выпуск, от которого идёт расход ресурса: суточный выпуск участка, его часы и число рабочих часов в сутках.
Driver = tuple[dict[date, float], list[Hour], dict[date, int]]


def _drivers(spec: Spec, sim: Simulation, resource: dict[str, Any]) -> list[Driver]:
    """Участки, от выпуска которых идёт расход газа или воды на технологию.

    Ресурс с `circuit` — от выпуска своего участка; без него — от суммы участков в единице продукции объекта.
    Участков нет (или ни одного в этой единице) — пусто: расход идёт от выпуска объекта, как раньше.
    """
    if not spec.areas:
        return []
    keys = [resource["circuit"]] if resource.get("circuit") else [
        area.key for area in spec.areas if area.unit == spec.product["unit"]]
    return [(sim.outputs[key], sim.area_hours[key], _work_hours(sim.area_hours[key])) for key in keys]


def _driven(drivers: list[Driver], i: int, day: date) -> float:
    """Выпуск участков, приходящийся на час `i`: суточный выпуск делится поровну на рабочие часы участка."""
    total = 0.0
    for output, hours, counts in drivers:
        if hours[i].is_work and counts[day]:
            total += output[day] / counts[day]
    return total


def _resource_anomaly(spec: Spec, resource: str, kind: str) -> dict[str, Any] | None:
    return next((a for a in spec.anomalies if a.get("resource") == resource and a["type"] == kind), None)


def simulate_resources(spec: Spec, sim: Simulation) -> dict[str, ResourceSeries]:
    """Часовой расход газа (м³), тепла (Гкал) и воды (м³). Топливо часового учёта не имеет."""
    work_hours = _work_hours(sim.hours)
    heating_days = _heating_days(sim)
    leak_from = len(sim.hours)
    result: dict[str, ResourceSeries] = {}
    for index, resource in enumerate(("gas", "heat", "water")):
        r = spec.resources.get(resource)
        if r is None:
            continue
        rng = random.Random(spec.seed + 10 + index)
        no_setback = _resource_anomaly(spec, resource, "heating_no_setback")
        leak = _resource_anomaly(spec, resource, "water_leak")
        if leak:
            leak_from = len(sim.hours) - 24 * int(leak.get("days", 21))
        peak = float(r.get("heating_peak_gcal_h" if resource == "heat" else "heating_peak_m3_h", 0) or 0)
        per_unit = float(r.get("process_m3_per_unit", 0) or 0)
        domestic = float(r.get("domestic_m3_per_work_day", 0) or 0)
        drivers = _drivers(spec, sim, r)
        hourly: list[float] = []
        baseline: list[float] = []
        for i, hour in enumerate(sim.hours):
            day = hour.ts.date()
            share = 1 / work_hours[day] if hour.is_work and work_hours[day] else 0.0
            process = per_unit * _driven(drivers, i, day) if drivers else per_unit * sim.production[day] * share
            flow = process + domestic * share
            heat = peak * _heating(hour.t_out) if day in heating_days else 0.0
            if resource == "water":
                flow += WATER_NIGHT_BASE_M3_H
            noise = 1 + rng.gauss(0, 0.04)
            clean = (flow + heat * (1.0 if hour.is_work else HEATING_SETBACK)) * noise
            value = (flow + heat * (1.0 if hour.is_work or no_setback else HEATING_SETBACK)) * noise
            if leak and i >= leak_from:
                value += float(leak.get("lph", 100)) / 1000
            digits = 4 if resource == "heat" else 3
            hourly.append(round(max(0.0, value), digits))
            baseline.append(round(max(0.0, clean), digits))
        result[resource] = ResourceSeries(hourly, baseline)
    return result


def resource_bills(spec: Spec, sim: Simulation, series: dict[str, ResourceSeries]) -> dict[str, list[ResourceBill]]:
    """Месячные счета по ресурсам: объём — сумма часов месяца, тариф индексируется 1 июля."""
    result: dict[str, list[ResourceBill]] = {}
    months = sorted({(h.ts.year, h.ts.month) for h in sim.hours})
    for resource, r in spec.resources.items():
        rows: list[ResourceBill] = []
        for year, month in months:
            factor = 1.0 if _indexed(spec, year, month) else 0.91
            if resource == "liquid_fuel":
                volume = _fuel_month(spec, sim, year, month)
                tariff = float(r["price_rub_per_l"]) * (1 + 0.02 * math.sin(month * 1.7))  # цена свободная
                rows.append(ResourceBill(resource, year, month, volume, round(tariff, 2), _vat(year)))
                continue
            volume = sum(v for h, v in zip(sim.hours, series[resource].hourly, strict=True)
                         if (h.ts.year, h.ts.month) == (year, month))
            key = {"gas": "tariff_rub_per_1000m3", "heat": "tariff_rub_per_gcal", "water": "tariff_rub_per_m3"}
            bill = ResourceBill(resource, year, month, round(volume, 3 if resource == "heat" else 0),
                                round(float(r[key[resource]]) * factor, 2), _vat(year))
            if resource == "water":
                bill.wastewater = round(volume * float(r.get("wastewater_share", 0.95)))
                bill.wastewater_tariff = round(float(r["wastewater_tariff_rub_per_m3"]) * factor, 2)
            rows.append(bill)
        result[resource] = rows
    return result


def _fuel_month(spec: Spec, sim: Simulation, year: int, month: int) -> float:
    """Литры топлива за месяц: норма на рабочие сутки × рабочие сутки, с разбросом заправок."""
    rng = random.Random(spec.seed + 20 + year * 12 + month)
    days = sum(1 for day, out in sim.production.items() if (day.year, day.month) == (year, month) and out)
    return float(round(float(spec.resources["liquid_fuel"]["l_per_work_day"]) * days * (1 + rng.gauss(0, 0.06)), -1))


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
        _describe_areas(d, spec)
    d.add_heading("3. Инженерные системы", level=2)
    systems = [e for e in spec.equipment if e.profile not in ("process",)]
    for e in systems:
        d.add_paragraph(f"{e.name}: {e.count} шт., {e.location or 'по месту'}.", style="List Bullet")
    if o.get("description"):
        d.add_heading("4. Особенности", level=2)
        d.add_paragraph(o["description"])
    d.save(path)


_WEEKDAYS = ("пн", "вт", "ср", "чт", "пт", "сб", "вс")


def _days_text(weekdays: list[int]) -> str:
    days = sorted(set(weekdays))
    if len(days) == 7:
        return "ежедневно"
    if len(days) > 2 and days == list(range(days[0], days[-1] + 1)):
        return f"{_WEEKDAYS[days[0]]}–{_WEEKDAYS[days[-1]]}"
    return ", ".join(_WEEKDAYS[day] for day in days)


def area_schedule(area: Area) -> str:
    """Окно участка словами: «с 06:00 до 18:00, пн–пт» — так его ставят окном контура."""
    s = area.shifts
    return f"с {int(s['start_hour']):02d}:00 до {int(s['end_hour']):02d}:00, {_days_text(s['weekdays'])}"


def _describe_areas(d: Any, spec: Spec) -> None:
    """Графики и выпуск участков — абзацами под таблицей контуров, саму таблицу не трогаем."""
    if not spec.areas:
        return
    d.add_paragraph("Графики работы и выпуск участков (выпуск учитывается по каждому участку отдельно):")
    for area in spec.areas:
        d.add_paragraph(f"{area.title}: работает {area_schedule(area)}; выпуск в рабочие сутки "
                        f"{_fmt(area.low)}–{_fmt(area.high)} {area.unit}.", style="List Bullet")


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
    for meter in resource_meters(spec):
        ws.append([meter["node"], meter["model"], 1, None, meter["location"]])
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
    return _write_payment_document(*_bill_lines(spec, bill, number), path_pdf, font)


def _write_payment_document(head: list[str], header: list[str], rows: list[list[str]], total: str,
                            path_pdf: Path, font: str | None) -> Path:
    """Счёт или накладная: шапка (первая и третья строки крупно), таблица, итог. Без шрифта — DOCX."""
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


def _period(year: int, month: int) -> str:
    last = calendar.monthrange(year, month)[1]
    return f"{MONTHS[month - 1]} {year} (01.{month:02d}.{year}–{last}.{month:02d}.{year})"


def _issued(year: int, month: int, delay: int) -> date:
    return date(year, month, calendar.monthrange(year, month)[1]) + timedelta(days=delay)


def resource_bill_lines(
    spec: Spec, bill: ResourceBill, number: int
) -> tuple[list[str], list[str], list[list[str]], str]:
    """Шапка, таблица и итог счёта по ресурсу. Слова в шапке — те, по которым агент узнаёт тип документа:
    «договор поставки газа», «водоотведение», «тепловая энергия, Гкал», «УПД» + «дизельное топливо»."""
    r, o = spec.resources[bill.resource], spec.obj
    buyer = f"{o['company']}, {o['address']}"
    vat_line = f"Всего к оплате: {_fmt(bill.gross, 2)} руб., в т. ч. НДС {int(bill.vat * 100)} %"
    if bill.resource == "gas":
        head = [r["supplier"], "Поставщик природного газа",
                f"Счёт № {number}/Г от {_issued(bill.year, bill.month, 5):%d.%m.%Y}",
                f"Расчётный период: {_period(bill.year, bill.month)}",
                f"Договор поставки газа № {r['contract_number']}", f"Покупатель: {buyer}"]
        header = ["Наименование", "Объём газа, тыс. м³", "Тариф без НДС, руб./тыс. м³", "Сумма без НДС, руб.",
                  "Сумма с НДС, руб."]
        rows = [["Газ природный", _fmt(bill.volume / 1000, 3), _fmt(bill.tariff, 2), _fmt(bill.net, 2),
                 _fmt(bill.gross, 2)]]
        return head, header, rows, vat_line
    if bill.resource == "heat":
        head = [r["supplier"], "Теплоснабжающая организация",
                f"Счёт № {number}-Т от {_issued(bill.year, bill.month, 7):%d.%m.%Y}",
                f"Расчётный период: {_period(bill.year, bill.month)}",
                f"Договор теплоснабжения № {r['contract_number']}", f"Потребитель: {buyer}"]
        header = ["Наименование услуги", "Ед. изм.", "Количество", "Тариф без НДС, руб./Гкал",
                  "Сумма без НДС, руб."]
        rows = [["Тепловая энергия на отопление", "Гкал", _fmt(bill.volume, 3), _fmt(bill.tariff, 2),
                 _fmt(bill.net, 2)]]
        return head, header, rows, vat_line
    if bill.resource == "water":
        head = [r["supplier"], "Организация водопроводно-канализационного хозяйства",
                f"Счёт № {number} от {_issued(bill.year, bill.month, 3):%d.%m.%Y}",
                f"Расчётный период: {_period(bill.year, bill.month)}",
                f"Договор холодного водоснабжения и водоотведения № {r['contract_number']}", f"Абонент: {buyer}"]
        header = ["№", "Наименование услуги", "Ед. изм.", "Количество", "Тариф без НДС, руб.", "Сумма без НДС, руб."]
        rows = [["1", "Холодное водоснабжение", "м³", _fmt(bill.volume), _fmt(bill.tariff, 2),
                 _fmt(round(bill.volume * bill.tariff, 2), 2)],
                ["2", "Водоотведение", "м³", _fmt(bill.wastewater), _fmt(bill.wastewater_tariff, 2),
                 _fmt(round(bill.wastewater * bill.wastewater_tariff, 2), 2)]]
        return head, header, rows, vat_line
    title = FUEL_KINDS[r.get("fuel_kind", "diesel")][0]
    net = bill.net
    head = [r["supplier"], "Продавец нефтепродуктов",
            f"Универсальный передаточный документ (УПД) № {number} от {_issued(bill.year, bill.month, 0):%d.%m.%Y}",
            # Название топлива — в шапке: в ячейке таблицы оно переносится по строкам, и в текстовом слое PDF
            # «топливо» и «дизельное» разъезжаются — признак накладной на топливо не срабатывает.
            f"Отгрузка: {FUEL_KINDS[r.get('fuel_kind', 'diesel')][3]} за {MONTHS[bill.month - 1]} {bill.year}, "
            "статус 1 (счёт-фактура и передаточный документ)",
            f"Продавец: {r['supplier']}", f"Покупатель: {buyer}"]
    header = ["Наименование товара", "Ед. изм.", "Количество", "Цена без НДС, руб./л", "Стоимость без НДС, руб.",
              f"НДС {int(bill.vat * 100)} %, руб.", "Стоимость с НДС, руб."]
    rows = [[title, "л", _fmt(bill.volume), _fmt(bill.tariff, 2), _fmt(net, 2), _fmt(bill.gross - net, 2),
             _fmt(bill.gross, 2)]]
    return head, header, rows, vat_line


#: Папка и имя файла счетов ресурса в `upload/`.
RESOURCE_FILES = {"gas": ("счета_газ", "Счёт_газ"), "heat": ("счета_тепло", "Счёт_тепло"),
                  "water": ("счета_вода", "Счёт_вода"), "liquid_fuel": ("топливо", "УПД_топливо")}


def write_resource_bills(spec: Spec, bills_: dict[str, list[ResourceBill]], upload: Path,
                         font: str | None) -> list[Path]:
    files: list[Path] = []
    for index, (resource, rows) in enumerate(bills_.items()):
        folder, stem = RESOURCE_FILES[resource]
        (upload / folder).mkdir(parents=True, exist_ok=True)
        for i, bill in enumerate(rows, start=1):
            path = upload / folder / f"{stem}_{bill.year}-{bill.month:02d}.pdf"
            number = 700 + index * 300 + i * 11
            files.append(_write_payment_document(*resource_bill_lines(spec, bill, number), path, font))
    return files


#: Колонка участка в отчётах о выпуске при двух контурах и больше. Полиметрика берёт её как точку учёта, а
#: название участка в ячейке сопоставляет с контуром модели (`app/intake/output.py`, `zone_circuits`): у
#: выпуска точки учёта не бывает, это контур. Широкую таблицу «Дата | Обжарка, кг | Фасовка, кг» агент
#: по контурам не разносит — зоны колонок он узнаёт только у часов загрузки и посетителей.
AREA_COLUMN = "Участок (точка учёта)"


def _area_groups(spec: Spec) -> dict[str, list[Area]]:
    """Участки по единице выпуска: у колонки выпуска одна единица, поэтому на единицу — свой лист."""
    groups: dict[str, list[Area]] = {}
    for area in spec.areas:
        groups.setdefault(area.unit, []).append(area)
    return groups


def write_area_production(spec: Spec, path: Path, head: str, columns: tuple[str, str],
                          rows: list[tuple[str, dict[str, float]]]) -> None:
    """Отчёт о выпуске по участкам: строка — период и участок, «Итого» нет (его агент взял бы за объект).

    `columns` — подписи колонок периода и выпуска («Месяц», «Выпуск продукции»), `rows` — подпись периода
    и выпуск участков за него.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font

    wb = Workbook()
    wb.remove(wb.active)
    groups = _area_groups(spec)
    for unit, members in groups.items():
        ws = wb.create_sheet("Выпуск" if len(groups) == 1 else f"Выпуск, {unit}")
        ws.append([head])
        ws.append([f"{spec.obj['name']}: {spec.product['title']}"])
        ws.append([])
        ws.append([columns[0], AREA_COLUMN, f"{columns[1]}, {unit}"])
        ws["A1"].font = Font(bold=True, size=13)
        for label, values in rows:
            for area in members:
                ws.append([label, area.title, round(values[area.key])])
        _fit_columns(ws, (18, max(len(AREA_COLUMN), *(len(a.title) for a in members)) + 4, 26), numeric="C")
    wb.save(path)


def _area_months(spec: Spec, sim: Simulation) -> list[tuple[str, dict[str, float]]]:
    """Выпуск участков по месяцам, кроме последнего: его сутки — в суточном отчёте."""
    totals: dict[tuple[int, int], dict[str, float]] = {}
    for key, days in sim.outputs.items():
        for day, value in days.items():
            month = totals.setdefault((day.year, day.month), {})
            month[key] = month.get(key, 0.0) + value
    totals.pop((spec.end.year, spec.end.month), None)
    return [(f"{MONTHS[month - 1]} {year}", values) for (year, month), values in sorted(totals.items())]


def _area_days(spec: Spec, sim: Simulation) -> list[tuple[str, dict[str, float]]]:
    first = spec.end.replace(day=1)
    days = sorted(day for day in sim.production if day >= first)
    return [(day.strftime("%d.%m.%Y"), {key: sim.outputs[key][day] for key in sim.outputs}) for day in days]


def write_production_year(spec: Spec, sim: Simulation, path: Path) -> None:
    if spec.areas:
        write_area_production(spec, path, "Отчёт о выпуске продукции по месяцам", ("Месяц", "Выпуск продукции"),
                              _area_months(spec, sim))
        return
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
    first = spec.end.replace(day=1)
    if spec.areas:
        head = f"Суточный отчёт о выпуске продукции за {MONTHS[first.month - 1]} {first.year}"
        write_area_production(spec, path, head, ("Дата", "Выпуск"), _area_days(spec, sim))
        return first, spec.end
    from openpyxl import Workbook  # noqa: PLC0415
    from openpyxl.styles import Font  # noqa: PLC0415

    unit = spec.product["unit"]
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


_RESOURCE_ANOMALY_LINES = {
    "water_leak": "По ночам на вводе слышно, как вода идёт, хотя никто не работает — где-то подтекает.",
    "heating_no_setback": "{place}: автоматики нет, ночью и в выходные топим так же, как в смену.",
}


def resource_call_lines(spec: Spec) -> list[tuple[str, str]]:
    """Реплики о газе, тепле, воде и топливе: кто поставляет, на что идёт, где учёт, что не так."""
    if not spec.resources:
        return []
    lines: list[tuple[str, str]] = [("Энергетик", "А кроме электричества что покупаем и на что оно идёт?")]
    r = spec.resources
    if "gas" in r:
        uses = [u for u, on in (("на горелки, на технологию", r["gas"].get("process_m3_per_unit")),
                                ("на котельную, на отопление", r["gas"].get("heating_peak_m3_h"))) if on]
        lines.append(("Технолог", f"Газ берём у {r['gas']['supplier']}, идёт {' и '.join(uses)}."))
    if "heat" in r:
        lines.append(("Технолог", f"Отопление от городской сети, {r['heat']['supplier']}, тепло приходит в ИТП."))
    if "water" in r:
        lines.append(("Технолог", f"Вода и канализация — {r['water']['supplier']}, вода в основном на мойку "
                                  "и на бытовые нужды."))
    if "liquid_fuel" in r:
        use = r["liquid_fuel"].get("use", "погрузчики")
        fuel = "Мазут" if r["liquid_fuel"].get("fuel_kind") == "fuel_oil" else "Дизель"
        lines.append(("Технолог", f"{fuel} покупаем у {r['liquid_fuel']['supplier']} на {use}, "
                                  "по накладным каждый месяц."))
    meters = resource_meters(spec)
    if meters:
        names = ", ".join(f"{_lower(m['node'])} ({m['model']})" for m in meters)
        lines.append(("Технолог", f"Приборы на вводах есть: {names}; почасовые выгрузки отдадим файлами."))
    for anomaly in spec.anomalies:
        if "resource" not in anomaly:
            continue
        place = "ИТП" if anomaly["resource"] == "heat" else "Котельная"
        template = anomaly.get("hint") or _RESOURCE_ANOMALY_LINES[anomaly["type"]]
        lines.append(("Технолог", template.format(place=place)))
        lines.append(("Энергетик", "Понял, посмотрим по счётчику на вводе."))
    return lines


def _lower(name: str) -> str:
    return name[:1].lower() + name[1:]


def _target_area(link: Link, by_key: dict[str, Equipment], areas: list[Area]) -> Area | None:
    """Участок оборудования, к которому идёт связь: `output` при участках — выпуск этого участка."""
    return next((area for area in areas if area.key == by_key[link.target].circuit), None)


def link_line(link: Link, by_key: dict[str, Equipment], product: dict[str, Any],
              areas: list[Area] | None = None) -> str:
    """Как технолог говорит о связи. Название из ведомости стоит в именительном падеже: склонять его
    генератор не умеет, а агенту нужно узнать узел."""
    if link.say:
        return link.say
    target = by_key[link.target].name
    if link.source == OUTPUT:
        area = _target_area(link, by_key, areas or [])
        whose = f"за выпуском участка «{area.title}»" if area else "за выпуском"
        return (f"{target}: потребление идёт {whose} — больше {area.unit if area else product['unit']} за день, "
                "больше и берёт, без выпуска простаивает.")
    source = _lower(by_key[link.source].name)
    if link.kind == "process":
        return f"Продукт идёт по цепочке: сначала {source}, потом {_lower(target)}."
    if link.kind == "unmeasured":
        return f"{target} включается, когда работает {source}; своего счётчика нет, сидит на общем."
    if link.method == "window_const":
        return (f"{target} включаем, только когда работает {source}; нагрузка ровная, "
                "от объёма почти не зависит.")
    return f"{target} включаем, только когда работает {source}; остановились — можно выключать."


def area_call_lines(spec: Spec) -> list[tuple[str, str]]:
    """Графики и выпуск участков словами технолога: по ним помощник ставит окна контуров."""
    if not spec.areas:
        return []
    schedules = "; ".join(f"{area.title} — {area_schedule(area)}" for area in spec.areas)
    outputs = "; ".join(f"{_lower(area.title)} — в среднем {_fmt((area.low + area.high) / 2)} {area.unit}"
                        for area in spec.areas)
    return [
        ("Энергетик", "Участки работают в одно время? Запишу окно работы каждого участка."),
        ("Технолог", f"У каждого участка свой график. {schedules}. Вне этих часов оборудование участка должно стоять."),
        ("Энергетик", "А выпуск считаете общий?"),
        ("Технолог", (f"Нет, по каждому участку отдельно, за рабочие сутки: {outputs}. В отчётах о выпуске так "
                      "и разбито по участкам.")),
    ]


def write_call_transcript(spec: Spec, path: Path) -> None:
    by_key = {e.key: e for e in spec.equipment}
    shifts = spec.obj["shifts"]["title"]
    lines: list[tuple[str, str]] = [
        ("Энергетик", "Давай пройдёмся по цеху: что работает и когда."),
        ("Технолог", f"Работаем так: {shifts}. Выпуск — {spec.product['title'].lower()}, "
                     f"в среднем {_fmt((spec.product['daily_min'] + spec.product['daily_max']) / 2)} "
                     f"{spec.product['unit']} за смену."),
        *area_call_lines(spec),
    ]
    if spec.links:
        lines.append(("Энергетик", "Расскажи, что от чего зависит: что идёт за выпуском, что работает под какой участок."))
        lines += [("Технолог", link_line(link, by_key, spec.product, spec.areas)) for link in spec.links]
        lines.append(("Энергетик", "Понял, так и свяжем в модели."))
    for anomaly in spec.anomalies:
        if "resource" in anomaly:
            continue
        e = by_key[anomaly["equipment"]]
        template = anomaly.get("hint") or _ANOMALY_LINES[anomaly["type"]]
        lines.append(("Технолог", template.format(name=e.name, unit=anomaly.get("unit", 1))))
        lines.append(("Энергетик", "Понял, отмечу, проверим по приборам."))
    lines += resource_call_lines(spec)
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


#: Узел ввода ресурса в ведомости — название из синонимов класса `input_meter.<ресурс>` каталога, прибор,
#: имя прибора в файле показаний (агент привязывает его к вводу своего ресурса по слову `gas`/`heat`/`water`
#: из ключа узла) и подпись колонки показания, по которой разбор узнаёт ресурс.
METER_NODES = {
    "gas": ("Узел учёта газа", "gas-vvod-1", "Показание счётчика газа, м³", "ГРПШ у корпуса"),
    "heat": ("Узел учёта тепла (ИТП)", "heat-vvod-1", "Показание счётчика тепла, Гкал", "Тепловой пункт"),
    "water": ("Водомер на вводе", "water-vvod-1", "Показание счётчика воды, м³", "Водомерный узел"),
}
METER_FILES = {"gas": "показания_приборов_учёта_газ.csv", "heat": "показания_приборов_учёта_тепло.csv",
               "water": "показания_приборов_учёта_вода.csv"}


def resource_meters(spec: Spec) -> list[dict[str, str]]:
    """Приборы учёта на вводах ресурсов: только у ресурсов с полем `meter` (марка прибора)."""
    meters = []
    for resource, (node, device, column, location) in METER_NODES.items():
        r = spec.resources.get(resource)
        if r and r.get("meter"):
            meters.append({"resource": resource, "node": node, "device": device, "column": column,
                           "model": r["meter"], "location": r.get("meter_location", location)})
    return meters


def write_resource_telemetry(spec: Spec, sim: Simulation, series: dict[str, ResourceSeries],
                             folder: Path) -> list[dict[str, str]]:
    """Почасовые показания счётчиков газа, тепла и воды нарастающим итогом, по файлу на ресурс.

    Отдельный файл — как отдельная выгрузка корректора, тепловычислителя или водомера. Показание — итог
    с начала года счетов: он растёт и на часах до окна выгрузки, поэтому файл начинается не с нуля."""
    folder.mkdir(parents=True, exist_ok=True)
    first = len(sim.hours) - 24 * spec.telemetry_days
    meters = resource_meters(spec)
    for meter in meters:
        hourly = series[meter["resource"]].hourly
        digits = 4 if meter["resource"] == "heat" else 3
        rng = random.Random(spec.seed + 30 + len(meter["device"]))
        total = rng.uniform(1_000, 9_000) + sum(hourly[:first])
        with (folder / METER_FILES[meter["resource"]]).open("w", encoding="utf-8", newline="") as fh:
            writer = csv.writer(fh, delimiter=";")
            writer.writerow(["Прибор учёта", "Дата и время", meter["column"]])
            for i in range(first, len(sim.hours)):
                total += hourly[i]
                writer.writerow([meter["device"], sim.hours[i].ts.strftime("%d.%m.%Y %H:%M"), _fmt(total, digits)])
    return meters


def write_telemetry(spec: Spec, sim: Simulation, folder: Path,
                    meters: list[dict[str, str]] | None = None) -> list[dict[str, Any]]:
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
    for meter in meters or []:
        ws.append([meter["device"], meter["model"], meter["node"], meter["location"]])
    ws.append(["street-t1", "Датчик температуры", "Наружный воздух", "Северная стена"])
    _fit_columns(ws, (24, 22, max(24, *(len(d["title"]) + 4 for d in devices)), 24))
    wb.save(folder / "перечень_приборов_учёта.xlsx")
    return [{k: v for k, v in d.items() if k != "series"} for d in devices]


LINK_KIND_TITLES = {"metric": "измерение", "process": "технологическая", "unmeasured": "без измерения"}
LINK_METHOD_TITLES = {"regression": "от выпуска", "window": "в окне процесса",
                      "window_const": "в окне, постоянная нагрузка"}


def _resource_rub(spec: Spec, resource: str, volume: float) -> float:
    """Рубли без НДС за объём ресурса по текущему тарифу; у воды — вместе с водоотведением."""
    r = spec.resources[resource]
    if resource == "gas":
        return volume / 1000 * float(r["tariff_rub_per_1000m3"])
    if resource == "heat":
        return volume * float(r["tariff_rub_per_gcal"])
    return volume * (float(r["tariff_rub_per_m3"]) + float(r["wastewater_tariff_rub_per_m3"]))


def _co2(spec: Spec, resource: str, volume: float) -> str:
    if resource == "gas":
        return _fmt(volume / 1000 * GAS_CO2_T_PER_1000M3, 1)
    if resource == "liquid_fuel":
        _, density, factor, _ = FUEL_KINDS[spec.resources[resource].get("fuel_kind", "diesel")]
        return _fmt(volume * density / 1000 * factor, 1)
    return "—"


def resource_answers(spec: Spec, sim: Simulation, series: dict[str, ResourceSeries],
                     rbills: dict[str, list[ResourceBill]]) -> tuple[list[str], list[dict[str, Any]]]:
    """Раздел «Покупные ресурсы» ответов: итоги по счетам, приборы на вводах, заложенные потери."""
    if not rbills:
        return [], []
    units = {"gas": "м³", "heat": "Гкал", "water": "м³", "liquid_fuel": "л"}
    tariff_units = {"gas": "₽/тыс. м³", "heat": "₽/Гкал", "water": "₽/м³", "liquid_fuel": "₽/л"}
    lines = ["", "## Покупные ресурсы", "",
             "Счета лежат в `upload/` по папкам ресурсов. Агент раскладывает их по своим рядам; ни одна строка "
             "не должна попасть в электроэнергию.", "",
             "| Ресурс | Объём за период | ₽ без НДС | Последний тариф без НДС | CO₂, т |", "|---|---:|---:|---:|---:|"]
    for resource, rows in rbills.items():
        digits = 3 if resource == "heat" else 0
        volume = sum(b.volume for b in rows)
        # У воды водоотведение — отдельной строкой ниже, здесь только холодная вода.
        rub = sum(round(b.volume * b.tariff, 2) for b in rows) if resource == "water" else sum(b.net for b in rows)
        lines.append(f"| {RESOURCE_TITLES[resource][0]} | {_fmt(volume, digits)} {units[resource]} | {_fmt(rub)} "
                     f"| {_fmt(rows[-1].tariff, 2)} {tariff_units[resource]} | {_co2(spec, resource, volume)} |")
        if resource == "water":
            lines.append(f"| Водоотведение | {_fmt(sum(b.wastewater for b in rows))} м³ "
                         f"| {_fmt(sum(round(b.wastewater * b.wastewater_tariff, 2) for b in rows))} "
                         f"| {_fmt(rows[-1].wastewater_tariff, 2)} ₽/м³ | — |")
    meters = resource_meters(spec)
    if meters:
        lines += ["", "Приборы на вводах (почасовые показания за последние "
                  f"{spec.telemetry_days} суток в `telemetry/`): "
                  + "; ".join(f"{m['device']} → «{m['node']}», `{METER_FILES[m['resource']]}`" for m in meters)
                  + ". Агент привязывает прибор к вводу своего ресурса; электрические расчёты от этого не меняются."]
    findings: list[dict[str, Any]] = []
    for anomaly in spec.anomalies:
        resource = anomaly.get("resource")
        if resource is None:
            continue
        if anomaly["type"] == "water_leak":
            yearly = float(anomaly.get("lph", 100)) / 1000 * 8760
            title = anomaly.get("title") or f"Утечка воды {_fmt(float(anomaly.get('lph', 100)))} л/ч"
            how = (f"последние {anomaly.get('days', 21)} суток ночной минимум расхода на вводе выше порога: "
                   "детектор ночной утечки поднимает инцидент с рублями по воде и водоотведению")
        else:
            delta = sum(series[resource].hourly) - sum(series[resource].baseline)
            yearly = delta * 8760 / len(sim.hours)
            title = anomaly.get("title") or f"{RESOURCE_TITLES[resource][0]}: отопление без ночного снижения"
            how = ("видно на часовом графике ввода и в преаудите по счетам; детекторы режима работы смотрят только "
                   "на электроэнергию, поэтому инцидента нет — спросите помощника")
        findings.append({"title": title, "type": anomaly["type"], "resource": resource, "unit": units[resource],
                         "volume_per_year": round(yearly, 1),
                         "rub_per_year": round(_resource_rub(spec, resource, yearly)),
                         "how_found": how})
    if findings:
        lines += ["", "### Заложенные потери по ресурсам", "", "| Что | Объём в год | ₽ в год (без НДС) | Как найти |",
                  "|---|---:|---:|---|",
                  *[f"| {f['title']} | {_fmt(f['volume_per_year'], 1 if f['unit'] == 'Гкал' else 0)} {f['unit']} "
                    f"| {_fmt(f['rub_per_year'])} | {f['how_found']} |" for f in findings]]
    return lines, findings


def _link_end(link: Link, by_key: dict[str, Equipment], areas: list[Area]) -> str:
    if link.source != OUTPUT:
        return by_key[link.source].name
    area = _target_area(link, by_key, areas)
    return f"Выпуск участка «{area.title}»" if area else "Выпуск продукции"


def _totals_lines(spec: Spec, total_kwh: float, total_out: float) -> list[str]:
    """Итоги периода. При участках выпуск разных участков не складывается, удельный — по участкам ниже."""
    if spec.areas:
        return [(f"- Потребление по счетам: {_fmt(total_kwh)} кВт·ч; выпуск — по участкам, см. «Выпуск "
                 "и удельный расход по участкам».")]
    return [
        f"- Потребление по счетам: {_fmt(total_kwh)} кВт·ч; выпуск: {_fmt(total_out)} {spec.product['unit']}.",
        f"- Удельный расход: {_fmt(total_kwh / total_out, 3) if total_out else '—'} кВт·ч/{spec.product['unit']}.",
    ]


def _area_kwh(spec: Spec, sim: Simulation, area: Area, hours: list[int], *, metered: bool) -> float:
    """кВт·ч оборудования участка за часы `hours`; `metered` — только со своим прибором учёта."""
    keys = [f"{e.key}#{n}" for e in spec.equipment if e.circuit == area.key and (e.metered or not metered)
            for n in range(1, e.count + 1)]
    return sum(sim.units[key][i] for key in keys for i in hours)


def _ratio(kwh: float, output: float) -> str:
    return _fmt(kwh / output, 3) if output else "—"


def area_answers(spec: Spec, sim: Simulation) -> list[str]:
    """Выпуск, окно и удельный расход каждого участка: что должен показать обзор «Выпуск и удельный расход»."""
    if not spec.areas:
        return []
    first = spec.end.replace(day=1)
    lines = ["", "## Выпуск и удельный расход по участкам", "",
             (f"Выпуск в отчётах записан по участкам (колонка «{AREA_COLUMN}»): помощник разносит его по контурам "
              "модели с теми же названиями. Полиметрика считает удельный по контурам по счётчикам узлов контура, "
              "общие системы (оборудование без участка) между участками не делит."), "",
             ("| Участок | Окно работы | Выпуск за период | кВт·ч оборудования участка | кВт·ч на единицу "
              "| Последний месяц по счётчикам узлов, кВт·ч на единицу |"),
             "|---|---|---:|---:|---:|---:|"]
    every = list(range(len(sim.hours)))
    for area in spec.areas:
        output = sim.outputs[area.key]
        total = sum(output.values())
        kwh = _area_kwh(spec, sim, area, every, metered=False)
        days = {day for day, value in output.items() if day >= first and value}
        month = [i for i, hour in enumerate(sim.hours) if hour.ts.date() in days]
        sensors = _area_kwh(spec, sim, area, month, metered=True)
        lines.append(f"| {area.title} | {area_schedule(area)} | {_fmt(total)} {area.unit} | {_fmt(kwh)} "
                     f"| {_ratio(kwh, total)} кВт·ч/{area.unit} "
                     f"| {_ratio(sensors, sum(output[day] for day in days))} кВт·ч/{area.unit} |")
    lines += ["", ("Окна контуров технолог назвал в разговоре (графики участков). Если после разбора у контура "
                   "нет окна работы, напишите помощнику: «Поставь окна работы контуров по графикам участков "
                   "из разговора с технологом».")]
    return lines


def write_answers(spec: Spec, sim: Simulation, bills_: list[Bill], path: Path,
                  series: dict[str, ResourceSeries] | None = None,
                  rbills: dict[str, list[ResourceBill]] | None = None) -> list[dict[str, Any]]:
    tariff = float(spec.supply["tariff_rub_per_kwh"])
    by_key = {e.key: e for e in spec.equipment}
    hours = len(sim.hours)
    findings = []
    for anomaly in spec.anomalies:
        if "resource" in anomaly:
            continue
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
        *_totals_lines(spec, total_kwh, total_out),
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
                  *[f"| {_link_end(link, by_key, spec.areas)} | {by_key[link.target].name} "
                    f"| {LINK_KIND_TITLES[link.kind]} "
                    f"| {LINK_METHOD_TITLES.get(link.method or '', '—')} |" for link in spec.links]]
    lines += area_answers(spec, sim)
    resource_lines, resource_findings = resource_answers(spec, sim, series or {}, rbills or {})
    lines += resource_lines
    findings += resource_findings
    hidden = [e for e in spec.equipment if not e.in_inventory]
    if hidden:
        lines += ["", "## Не внесено в ведомость", "",
                  *[f"- {e.name}, {_fmt(e.rated_kw)} кВт × {e.count} — упоминается только в разговоре с технологом."
                    for e in hidden]]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return findings


_HOW_TO_HEAD = """# Как загрузить пакет в Полиметрику

1. Откройте app.polimetrika.ru → «Новый объект».
"""
_HOW_TO_ALL = """2. В чат помощника приложите ВСЕ файлы из папки `upload/` (можно одним сообщением) и напишите:
   «Разбери документы и заполни проект».
3. Дождитесь распознавания (1–5 минут) и проверьте карточки предложений: параметры объекта,
   энергомодель, оборудование, ряды по счетам, выпуск. Связи, которые технолог проговорил в разговоре,
   лежат в той же карточке «Оборудование и учёт», что и узлы, с цитатой из разговора. Нажмите «Подтвердить».
"""


def _how_to_tail(step: int) -> str:
    ask = "Спросите помощника о разговоре с технологом: «Что из разговора с технологом стоит внести в модель?»"
    return f"""{step}. {ask}

{step + 1}. Показания приборов: приложите в тот же чат `telemetry/показания_приборов_учёта.csv` вместе
   с `telemetry/перечень_приборов_учёта.xlsx` и напишите «Загрузи показания приборов учёта из файла».
   Подтвердите карточку «Загрузить показания» — почасовые данные уйдут в хранилище показаний за минуту.
   Проверить: «Сколько кВт·ч потребило <оборудование> за <месяц> по суткам?»

`telemetry/температура_наружного_воздуха.csv` пока не загружайте: он понадобится, когда расчёты начнут
учитывать погоду.

`ЧТО ЗАЛОЖЕНО.md` — ответы: какие потери спрятаны в данных. Откройте после того, как посмотрите анализ.
"""


HOW_TO = _HOW_TO_HEAD + _HOW_TO_ALL + _how_to_tail(4)

#: Папка отчётов о выпуске при участках: их прикладывают вторым сообщением, когда контуры уже в модели.
#: Выпуск, разобранный раньше контуров, агент складывает в выпуск объекта («по зонам не разнесено»).
PRODUCTION_LATER = "выпуск_вторым_сообщением"


def _area_steps(spec: Spec) -> str:
    """Шаги 2–4 при участках: сначала документы без отчётов о выпуске, потом отчёты отдельным сообщением."""
    titles = ", ".join(f"«{area.title}»" for area in spec.areas)
    return f"""2. В чат помощника приложите все файлы из папки `upload/`, КРОМЕ папки `upload/{PRODUCTION_LATER}/`,
   и напишите «Разбери документы и заполни проект». Отчёты о выпуске пока не прикладывайте: выпуск
   разносится по контурам модели, а контуров в ней ещё нет.
3. Дождитесь распознавания (1–5 минут) и проверьте карточки предложений: параметры объекта,
   энергомодель с контурами {titles}, оборудование, ряды по счетам. Связи, которые технолог проговорил
   в разговоре, лежат в той же карточке «Оборудование и учёт», что и узлы, с цитатой из разговора.
   Нажмите «Подтвердить».
4. Только после подтверждения приложите вторым сообщением оба файла из `upload/{PRODUCTION_LATER}/`
   (`05_Отчёт_о_выпуске_по_месяцам.xlsx` и `06_Суточный_выпуск.xlsx`) и напишите:
   «Вот отчёты о выпуске по участкам, разнеси выпуск по контурам».
   Подтвердите карточку выпуска: в ней выпуск каждого контура (колонка «{AREA_COLUMN}») и окна
   работы контуров из разговора с технологом. Если окон в карточке нет, напишите:
   «Поставь окна работы контуров по графикам участков из разговора с технологом».
"""


def how_to(spec: Spec) -> str:
    """`КАК ЗАГРУЗИТЬ.md`: без участков — все документы одним сообщением, при участках — выпуск вторым."""
    if not spec.areas:
        return HOW_TO + resource_how_to(spec)
    return _HOW_TO_HEAD + _area_steps(spec) + _how_to_tail(5) + resource_how_to(spec, 7)


def resource_how_to(spec: Spec, step: int = 6) -> str:
    """Шаги загрузки счетов и показаний газа, тепла, воды и топлива — только для описанных ресурсов."""
    if not spec.resources:
        return ""
    folders = ", ".join(f"`upload/{RESOURCE_FILES[r][0]}/`" for r in spec.resources)
    text = f"""
## Газ, тепло, вода и топливо

{step}. Счета по другим ресурсам лежат в {folders}.
   Если не приложили их на шаге 2, приложите вторым сообщением и напишите «Вот счета за газ, тепло, воду
   и топливо». Подтвердите карточки «Документы» и «Ряды по счетам»:
   у каждого ресурса свои ряды, тариф и удельный расход, в обзоре появится переключатель ресурса и «Всего, ₽».
"""
    meters = resource_meters(spec)
    if meters:
        files = ", ".join(f"`telemetry/{METER_FILES[m['resource']]}`" for m in meters)
        text += f"""{step + 1}. Показания счётчиков на вводах: {files}.
   Каждый файл — отдельным сообщением с текстом «Загрузи показания приборов учёта из файла».
   Прибор привяжется к узлу ввода своего ресурса из ведомости («{meters[0]['node']}» и др.).
   Проверить: «Сколько м³ газа (Гкал, м³ воды) ушло на вводе за последнюю неделю по суткам?»
"""
    return text


def pack_upload(upload: Path, target: Path) -> Path:
    """Складывает `upload/` в ZIP с кириллическими именами, которые Windows читает без кракозябр.

    `zipfile` сам ставит флаг UTF-8 (бит 11) для имён вне ASCII. Архив, собранный утилитой `zip` или
    моделью «на глаз», этого флага часто не имеет — тогда Проводник Windows читает имена в cp866.
    """
    with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(upload.rglob("*")):
            if path.is_file():
                zf.write(path, path.relative_to(upload).as_posix())
    return target


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
    production = upload / PRODUCTION_LATER if spec.areas else upload
    production.mkdir(exist_ok=True)
    files.append(production / "05_Отчёт_о_выпуске_по_месяцам.xlsx")
    write_production_year(spec, sim, files[-1])
    files.append(production / "06_Суточный_выпуск.xlsx")
    write_production_daily(spec, sim, files[-1])
    files.append(upload / "07_Разговор_с_технологом.docx")
    write_call_transcript(spec, files[-1])

    flows = simulate_resources(spec, sim)
    rbills = resource_bills(spec, sim, flows)
    files += write_resource_bills(spec, rbills, upload, font)

    meters = write_resource_telemetry(spec, sim, flows, out / "telemetry")
    devices = write_telemetry(spec, sim, out / "telemetry", meters)
    findings = write_answers(spec, sim, bills_, out / "ЧТО ЗАЛОЖЕНО.md", flows, rbills)
    (out / "КАК ЗАГРУЗИТЬ.md").write_text(how_to(spec), encoding="utf-8")

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
        **({"production_by_circuit": {
            area.key: {"title": area.title, "unit": area.unit, "total": round(sum(sim.outputs[area.key].values()))}
            for area in spec.areas}} if spec.areas else {}),
        "findings": findings,
        "devices": devices,
        "resources": {
            resource: {"unit": RESOURCE_TITLES[resource][1],
                       "bills": [{"month": f"{b.year}-{b.month:02d}", "volume": b.volume, "gross_rub": b.gross,
                                  **({"wastewater_m3": b.wastewater} if resource == "water" else {})}
                                 for b in rows]}
            for resource, rows in rbills.items()
        },
        "resource_meters": meters,
        "pdf_font": font is not None,
        "upload_files": [str(p.relative_to(out)) for p in files],
        "upload_zip": str(pack_upload(upload, out.with_name(f"{out.name}-upload.zip"))),
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
    print(json.dumps({**{k: summary[k] for k in ("object", "period", "total_kwh", "production_total",
                                                "findings", "pdf_font", "upload_files", "upload_zip")},
                     "resources": {k: len(v["bills"]) for k, v in summary["resources"].items()}},
                     ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
