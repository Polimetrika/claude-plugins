"""Тесты генератора демо-пакета: совместимость без участков, выпуск, график и ресурсы по участкам, связи
по правилам модели, счётчик электроэнергии на вводе.

Запуск из корня репозитория: `pytest plugins/polimetrika-demo/skills/generate-project-pack/tests`.
Нужны openpyxl, python-docx; reportlab — по желанию (без него счета выходят в DOCX, на тесты это не влияет).
"""

from __future__ import annotations

import copy
import hashlib
import importlib.util
import json
import os
import re
import sys
from collections import defaultdict
from datetime import datetime
from itertools import pairwise
from pathlib import Path

import docx
import openpyxl
import pytest

SKILL = Path(__file__).resolve().parents[1]
FIXTURES = Path(__file__).resolve().parent / "fixtures"


def _module(path: Path):
    spec = importlib.util.spec_from_file_location("generate_under_test", path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module  # dataclasses ищут модуль класса в sys.modules
    spec.loader.exec_module(module)
    return module


#: `GENERATE_PY` — другой файл генератора: так считались эталонные отпечатки ниже.
gen = _module(Path(os.environ.get("GENERATE_PY", SKILL / "scripts" / "generate.py")))

#: Отпечатки пакета 0.4.0 на `fixtures/no-circuits.json`: без контуров и с одним контуром. 0.5.0 без `supply.meter`
#: пакет не меняет — счётчик на вводе только по полю.
GOLDEN = {
    "no_circuits": "9a17ac960765b440df3613c9ea16704eeaf981b37c9be2fa8b0e0cf1905348a9",
    "one_circuit": "ab45be073d0ff3a124c0cb5362910816ee26860e3e5aa8e1b65a271604249f5a",
}
#: Отпечаток только чисел (счета, показания, выпуск, ведомость — `digest(numbers=True)`) генератора 0.2.1, до
#: участков, — один на оба случая: в 0.3.0 и 0.4.0 пакет без участков в цифрах не изменился. 0.4.0 поменял
#: лишь связи — разговор с технологом и таблицу связей в «ЧТО ЗАЛОЖЕНО.md».
NUMBERS = "b13a5285808fd2b01ee8f59dfe044706f12a7241eae91576359f8bf7ff1a3b1b"


def _raw(name: str) -> dict:
    return json.loads((FIXTURES / name).read_text(encoding="utf-8"))


#: Узлы фикстуры, которые связаны в разговоре: при контурах в модели они стоят в контуре — связь от выпуска
#: идёт только к узлу с контуром, а процесс и система связаны только внутри контура.
LINKED = ("roaster", "packing_line", "compressor", "nitrogen", "exhaust")


def _one_circuit(raw: dict) -> dict:
    raw = copy.deepcopy(raw)
    raw["circuits"] = [{"key": "production", "title": "Производство", "driver_unit": "кг"}]
    for item in raw["equipment"]:
        if item["key"] in LINKED:
            item["circuit"] = "production"
    return raw


def _load(raw: dict, tmp_path: Path, name: str = "spec.json"):
    path = tmp_path / name
    path.write_text(json.dumps(raw, ensure_ascii=False), encoding="utf-8")
    return gen.load_spec(path)


def _xlsx(path: Path) -> list:
    return [(ws.title, [list(row) for row in ws.iter_rows(values_only=True)]) for ws in openpyxl.load_workbook(path)]


def _docx(path: Path) -> list:
    document = docx.Document(path)
    return [p.text for p in document.paragraphs] + [
        [cell.text for cell in row.cells] for table in document.tables for row in table.rows
    ]


def digest(out: Path, *, numbers: bool = False) -> str:
    """Отпечаток содержимого пакета без того, что зависит от машины: пути архива, шрифта PDF, метаданных
    DOCX/XLSX. Счета — по `summary.json` (кВт·ч, суммы), документы — по тексту и ячейкам. `numbers` — только
    числа: `summary.json`, показания (CSV) и таблицы (XLSX), без текстов (MD, DOCX)."""
    summary = json.loads((out / "summary.json").read_text(encoding="utf-8"))
    for key in ("upload_zip", "pdf_font", "upload_files"):
        summary.pop(key)
    parts: list = [summary]
    for path in sorted(out.rglob("*")):
        if path.suffix == ".csv" or (path.suffix == ".md" and not numbers):
            parts.append((path.name, path.read_text(encoding="utf-8")))
        elif path.suffix == ".xlsx":
            parts.append((path.name, _xlsx(path)))
        elif path.suffix == ".docx" and path.parent.name == "upload" and not numbers:
            parts.append((path.name, _docx(path)))
    text = json.dumps(parts, ensure_ascii=False, default=str, sort_keys=True)
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


# --- Совместимость ---------------------------------------------------------------------------------


@pytest.mark.parametrize("case", ["no_circuits", "one_circuit"])
def test_pack_without_areas_is_unchanged(case: str, tmp_path: Path) -> None:
    raw = _raw("no-circuits.json")
    spec = _load(raw if case == "no_circuits" else _one_circuit(raw), tmp_path)
    gen.build(spec, tmp_path / "pack")
    assert digest(tmp_path / "pack", numbers=True) == NUMBERS
    assert digest(tmp_path / "pack") == GOLDEN[case]


def test_same_seed_same_pack(tmp_path: Path) -> None:
    spec = gen.load_spec(SKILL / "examples" / "coffee-roastery.json")
    gen.build(spec, tmp_path / "a")
    gen.build(spec, tmp_path / "b")
    assert digest(tmp_path / "a") == digest(tmp_path / "b")


# --- Участки -----------------------------------------------------------------------------------------


@pytest.fixture(scope="module")
def coffee(tmp_path_factory: pytest.TempPathFactory):
    spec = gen.load_spec(SKILL / "examples" / "coffee-roastery.json")
    out = tmp_path_factory.mktemp("coffee") / "pack"
    gen.build(spec, out)
    return spec, gen.simulate(spec), out


def _rows(path: Path) -> list[list]:
    """Строки данных отчёта о выпуске: после двух строк названия, пустой строки и шапки."""
    _, rows = _xlsx(path)[0]
    return rows[4:]


def test_production_reports_are_split_by_area_titles(coffee) -> None:
    spec, _, out = coffee
    titles = {c["title"] for c in spec.raw["circuits"]}
    for name, first in (("05_Отчёт_о_выпуске_по_месяцам.xlsx", "Месяц"), ("06_Суточный_выпуск.xlsx", "Дата")):
        (_, rows), = _xlsx(out / "upload" / gen.PRODUCTION_LATER / name)
        assert rows[3] == [first, gen.AREA_COLUMN, f"{'Выпуск продукции' if first == 'Месяц' else 'Выпуск'}, кг"]
        data = rows[4:]
        assert {row[1] for row in data} == titles  # подписи — названия контуров дословно
        assert not any("итог" in str(cell).lower() for row in data for cell in row)


def test_daily_sums_match_monthly_report(coffee) -> None:
    spec, sim, out = coffee
    later = out / "upload" / gen.PRODUCTION_LATER
    monthly = {(row[0], row[1]): row[2] for row in _rows(later / "05_Отчёт_о_выпуске_по_месяцам.xlsx")}
    sums: dict[tuple[str, str], float] = defaultdict(float)
    for area in spec.areas:
        for day, value in sim.outputs[area.key].items():
            sums[(f"{gen.MONTHS[day.month - 1]} {day.year}", area.title)] += value
    assert monthly
    assert all(sums[key] == value for key, value in monthly.items())
    daily = _rows(later / "06_Суточный_выпуск.xlsx")
    by_title = {area.title: area.key for area in spec.areas}
    last_month = sum(row[2] for row in daily)
    first = spec.end.replace(day=1)
    assert last_month == sum(v for area in spec.areas for d, v in sim.outputs[area.key].items() if d >= first)
    for row in daily[:6]:
        day = next(d for d in sim.production if d.strftime("%d.%m.%Y") == row[0])
        assert sim.outputs[by_title[row[1]]][day] == row[2]


def test_area_equipment_idle_outside_its_window(coffee) -> None:
    spec, sim, _ = coffee
    roasting = next(a for a in spec.areas if a.key == "roasting")
    workdays = set(roasting.shifts["weekdays"])
    series = sim.units["roaster#1"]
    outside = [kw for hour, kw in zip(sim.hours, series) if not gen._in_shift(roasting.shifts, hour.ts, workdays)]
    assert outside and max(outside) == 0.0
    # Обжарка начинает в 6, раньше смены объекта (7): участок живёт по своему окну, а не по графику объекта.
    early = [kw for hour, kw in zip(sim.hours, series) if hour.ts.hour == 6 and hour.ts.weekday() in workdays]
    assert max(early) > 0
    packing = [kw for hour, kw in zip(sim.hours, sim.units["packing_line#1"]) if hour.ts.hour in (6, 7, 20)]
    assert max(packing) == 0.0  # фасовка — с 8 до 20


def test_off_shift_anomaly_follows_area_window(tmp_path: Path) -> None:
    raw = json.loads((SKILL / "examples" / "coffee-roastery.json").read_text(encoding="utf-8"))
    raw["anomalies"] = [{"type": "off_shift", "equipment": "packing_line", "share": 0.3}]
    spec = _load(raw, tmp_path)
    sim = gen.simulate(spec)
    packing = next(a for a in spec.areas if a.key == "packing")
    workdays = set(packing.shifts["weekdays"])
    # 19:00 — смена объекта уже кончилась, а окно фасовки ещё идёт: это работа, а не потери вне смены.
    at_19 = [kw - base for hour, kw, base in zip(sim.hours, sim.units["packing_line#1"], sim.baseline["packing_line#1"])
             if hour.ts.hour == 19 and hour.ts.weekday() in workdays]
    assert max(at_19) == pytest.approx(0.0, abs=1e-9)
    night = [kw for hour, kw in zip(sim.hours, sim.units["packing_line#1"]) if hour.ts.hour == 23]
    assert min(night) > 0


def test_resource_with_circuit_follows_its_area(coffee) -> None:
    spec, sim, _ = coffee
    gas = gen.simulate_resources(spec, sim)["gas"].hourly
    roasting = next(a for a in spec.areas if a.key == "roasting")
    workdays = set(roasting.shifts["weekdays"])
    outside = [m3 for hour, m3 in zip(sim.hours, gas) if not gen._in_shift(roasting.shifts, hour.ts, workdays)]
    assert max(outside) == 0.0  # газ без отопления — только на обжарку и только в её окно
    per_unit = spec.resources["gas"]["process_m3_per_unit"]
    by_day: dict = defaultdict(float)
    for hour, m3 in zip(sim.hours, gas):
        by_day[hour.ts.date()] += m3
    expected = sum(per_unit * value for value in sim.outputs["roasting"].values())
    assert sum(by_day.values()) == pytest.approx(expected, rel=0.01)


def test_summary_has_output_by_circuit(coffee) -> None:
    spec, sim, out = coffee
    summary = json.loads((out / "summary.json").read_text(encoding="utf-8"))
    assert summary["production_by_circuit"] == {
        area.key: {"title": area.title, "unit": area.unit, "total": round(sum(sim.outputs[area.key].values()))}
        for area in spec.areas
    }
    answers = (out / "ЧТО ЗАЛОЖЕНО.md").read_text(encoding="utf-8")
    assert "## Выпуск и удельный расход по участкам" in answers
    assert all(f"| {area.title} | {gen.area_schedule(area)} |" in answers for area in spec.areas)


def test_transcript_names_area_windows(coffee) -> None:
    _, _, out = coffee
    text = "\n".join(_docx(out / "upload" / "07_Разговор_с_технологом.docx"))
    assert "Обжарка — с 06:00 до 18:00, пн–пт; Фасовка — с 08:00 до 20:00, пн–пт" in text


def test_areas_without_own_fields_split_object_output(tmp_path: Path) -> None:
    raw = json.loads((SKILL / "examples" / "asphalt-plant.json").read_text(encoding="utf-8"))
    for circuit in raw["circuits"]:
        circuit.pop("daily_min"), circuit.pop("daily_max"), circuit.pop("shifts")
    spec = _load(raw, tmp_path)
    sim = gen.simulate(spec)
    for day, total in sim.production.items():
        assert all(sim.outputs[a.key][day] == round(total / 2) for a in spec.areas)
    assert all(a.shifts["start_hour"] == spec.obj["shifts"]["start_hour"] for a in spec.areas)


@pytest.mark.parametrize(
    ("change", "message"),
    [
        (lambda raw: raw["circuits"].__setitem__(0, {**raw["circuits"][0], "daily_min": 10, "daily_max": 20}),
         "у двух контуров и больше"),
        (lambda raw: raw["resources"]["gas"].__setitem__("circuit", "nope"), "нет контура «nope»"),
        (lambda raw: raw["resources"]["heat"].__setitem__("circuit", "roasting"), "только у газа и воды"),
    ],
)
def test_spec_errors(change, message: str, tmp_path: Path) -> None:
    raw = _one_circuit(_raw("no-circuits.json"))
    change(raw)
    with pytest.raises(gen.SpecError, match=message):
        _load(raw, tmp_path)


def test_resource_without_circuit_follows_areas_in_product_unit(coffee) -> None:
    spec, sim, _ = coffee
    # Вода без `circuit`: технология — от суммы участков в единице продукции (обе в кг), газ — только обжарка.
    water = gen._drivers(spec, sim, spec.resources["water"])
    gas = gen._drivers(spec, sim, spec.resources["gas"])
    assert [driver[0] for driver in water] == [sim.outputs["roasting"], sim.outputs["packing"]]
    assert [driver[0] for driver in gas] == [sim.outputs["roasting"]]


def test_areas_with_different_units_get_own_sheets(tmp_path: Path) -> None:
    raw = json.loads((SKILL / "examples" / "asphalt-plant.json").read_text(encoding="utf-8"))
    raw["circuits"][1]["driver_unit"] = "шт"
    spec = _load(raw, tmp_path)
    out = tmp_path / "pack"
    gen.build(spec, out)
    sheets = _xlsx(out / "upload" / gen.PRODUCTION_LATER / "06_Суточный_выпуск.xlsx")
    assert [title for title, _ in sheets] == ["Выпуск, т", "Выпуск, шт"]
    assert [rows[3][2] for _, rows in sheets] == ["Выпуск, т", "Выпуск, шт"]
    assert [{row[1] for row in rows[4:]} for _, rows in sheets] == [
        {"Сушка и нагрев каменных материалов"}, {"Смешение и выдача смеси"}]


@pytest.mark.parametrize(
    ("change", "message"),
    [
        (lambda raw: raw["circuits"][0].pop("daily_max"), "оба поля"),
        (lambda raw: raw["circuits"][0].__setitem__("daily_min", 5000), "daily_min ≤ daily_max"),
        (lambda raw: raw["circuits"][1].__setitem__("title", "обжарка!"), "должны различаться"),
    ],
)
def test_area_spec_errors(change, message: str, tmp_path: Path) -> None:
    raw = json.loads((SKILL / "examples" / "coffee-roastery.json").read_text(encoding="utf-8"))
    change(raw)
    with pytest.raises(gen.SpecError, match=message):
        _load(raw, tmp_path)


def test_production_goes_second_when_areas(coffee) -> None:
    """Разбор не видит контуров из открытых карточек прошлых ходов: при участках отчёты о выпуске лежат
    отдельно и прикладываются вторым сообщением, после подтверждения модели. Оборудование по контурам разбор
    ставит сам (агент, docs#678 срез I) — ручного шага «отнеси оборудование к контурам» больше нет."""
    _, _, out = coffee
    later = out / "upload" / gen.PRODUCTION_LATER
    assert sorted(p.name for p in later.iterdir()) == ["05_Отчёт_о_выпуске_по_месяцам.xlsx", "06_Суточный_выпуск.xlsx"]
    assert not list((out / "upload").glob("0[56]_*"))
    text = (out / "КАК ЗАГРУЗИТЬ.md").read_text(encoding="utf-8")
    first = text.index(f"КРОМЕ папки `upload/{gen.PRODUCTION_LATER}/`")
    confirm = text.index("энергомодель с контурами «Обжарка», «Фасовка»")
    second = text.index("«Вот отчёты о выпуске по участкам, разнеси выпуск по контурам»")
    assert first < confirm < second
    assert "Отнеси оборудование к контурам" not in text
    assert "«Поставь окна работы контуров по графикам участков из разговора с технологом»" in text
    assert "ВСЕ файлы из папки `upload/` (можно одним сообщением)" not in text
    steps = [int(line.split(".")[0]) for line in text.splitlines() if line[:1].isdigit()]
    assert steps == list(range(1, len(steps) + 1))  # шаги по порядку, без пропусков и повторов


def test_production_in_upload_without_areas(tmp_path: Path) -> None:
    spec = _load(_raw("no-circuits.json"), tmp_path)
    gen.build(spec, tmp_path / "pack")
    assert (tmp_path / "pack" / "upload" / "05_Отчёт_о_выпуске_по_месяцам.xlsx").exists()
    assert not (tmp_path / "pack" / "upload" / gen.PRODUCTION_LATER).exists()
    text = (tmp_path / "pack" / "КАК ЗАГРУЗИТЬ.md").read_text(encoding="utf-8")
    assert "ВСЕ файлы из папки `upload/` (можно одним сообщением)" in text
    assert "Отнеси оборудование к контурам" not in text
    assert text == gen.HOW_TO + gen.resource_how_to(spec)


# --- Связи по правилам модели ----------------------------------------------------------------------------


@pytest.fixture(scope="module")
def zhbi(tmp_path_factory: pytest.TempPathFactory):
    """Эталон «Примера» — `sample-packs/zhbi-pribrezhny/spec.json` агента (копия в `fixtures/`)."""
    spec = gen.load_spec(FIXTURES / "zhbi-pribrezhny.json")
    out = tmp_path_factory.mktemp("zhbi") / "pack"
    gen.build(spec, out)
    return spec, out


def test_reference_spec_passes_link_rules(zhbi) -> None:
    spec, out = zhbi
    by_key = {e.key: e for e in spec.equipment}
    assert spec.links and all(gen.link_problem(link, by_key, has_circuits=True) is None for link in spec.links)
    assert {link.kind for link in spec.links} <= {"metric", "unmeasured"}
    summary = json.loads((out / "summary.json").read_text(encoding="utf-8"))
    assert set(summary["production_by_circuit"]) == {"concrete", "rebar", "molding"}
    assert "технологическая" not in (out / "ЧТО ЗАЛОЖЕНО.md").read_text(encoding="utf-8")


def test_transcript_names_every_link_target(zhbi) -> None:
    spec, out = zhbi
    text = "\n".join(_docx(out / "upload" / "07_Разговор_с_технологом.docx"))
    by_key = {e.key: e for e in spec.equipment}
    lines = [gen.link_line(link, by_key, spec.product, spec.areas) for link in spec.links]
    assert all(line in text and by_key[link.target].name in line for link, line in zip(spec.links, lines))
    assert "Продукт идёт по цепочке" not in text


def _zhbi_raw() -> dict:
    return _raw("zhbi-pribrezhny.json")


def _add_link(raw: dict, **link) -> None:
    raw["links"].append(link)


def _set(raw: dict, key: str, **fields) -> None:
    item = next(e for e in raw["equipment"] if e["key"] == key)
    for name, value in fields.items():
        if value is None:
            item.pop(name, None)
        else:
            item[name] = value


@pytest.mark.parametrize(
    ("change", "message"),
    [
        # Связи эталона до 0.4.0, которые черновик агента отклонял.
        (lambda raw: _add_link(raw, **{"from": "vibro", "to": "crane", "kind": "metric", "method": "window"}),
         "связь процесс → процесс правилами модели не предусмотрена"),
        (lambda raw: _add_link(raw, **{"from": "straightener", "to": "welder", "kind": "process"}),
         "вид «process» правила модели Полиметрики не принимают"),
        (lambda raw: _set(raw, "compressor", circuit=None),
         (r"узлы в разных контурах \(mixer \(бетоносмеситель двухвальный\): concrete, compressor \(компрессор "
          r"винтовой\): без контура\)")),
        (lambda raw: _set(raw, "aspiration", profile="process"), "связь процесс → процесс"),
        # Прочие правила.
        (lambda raw: _add_link(raw, **{"from": "welder", "to": "steam_chamber", "kind": "unmeasured"}),
         "связей между контурами нет"),
        (lambda raw: _add_link(raw, **{"from": "mixer", "to": "lighting", "kind": "unmeasured"}),
         "связей между контурами нет"),
        (lambda raw: _add_link(raw, **{"from": "output", "to": "boiler_pump", "kind": "metric",
                                       "method": "regression"}),
         "от выпуска связь идёт только к технологическому оборудованию"),
        (lambda raw: _add_link(raw, **{"from": "compressor", "to": "conveyor", "kind": "unmeasured"}),
         "обратная связь система → процесс"),
        (lambda raw: _set(raw, "crane", circuit=None), "у crane .* нет контура, а в модели контуры есть"),
        (lambda raw: raw["links"][0].__setitem__("method", "window_const"), "метод: regression, window"),
        (lambda raw: raw["links"][1].update(kind="unmeasured", method=None),
         "вид связи: metric"),
    ],
)
def test_forbidden_links_are_spec_errors(change, message: str, tmp_path: Path) -> None:
    raw = _zhbi_raw()
    change(raw)
    with pytest.raises(gen.SpecError, match=message):
        _load(raw, tmp_path)


def test_say_must_name_link_target(tmp_path: Path) -> None:
    raw = _zhbi_raw()
    link = next(item for item in raw["links"] if item["to"] == "compressor")
    link["say"] = "Компрессоры нужны только под замес, без замесов их можно глушить."
    with pytest.raises(gen.SpecError, match="say: в реплике нет названия «Компрессор винтовой» из ведомости"):
        _load(raw, tmp_path)
    link["say"] = "А компрессор ВИНТОВОЙ нужен только под замес."  # регистр и знаки не важны
    assert _load(raw, tmp_path).links


@pytest.mark.parametrize("name", ["examples/coffee-roastery.json", "examples/asphalt-plant.json",
                                  "tests/fixtures/zhbi-pribrezhny.json"])
def test_default_links_follow_model_rules(name: str, tmp_path: Path) -> None:
    raw = json.loads((SKILL / name).read_text(encoding="utf-8"))
    raw.pop("links")
    spec = _load(raw, tmp_path)
    by_key = {e.key: e for e in spec.equipment}
    assert spec.links
    assert [gen.link_problem(link, by_key, has_circuits=True) for link in spec.links] == [None] * len(spec.links)
    # Технологическое оборудование без счётчика связи не получает: процесс → процесс правила не принимают.
    unmetered = {e.key for e in spec.equipment if e.profile == "process" and not e.metered}
    assert not unmetered & {link.target for link in spec.links}


def test_examples_pass_link_rules() -> None:
    for name in ("coffee-roastery.json", "asphalt-plant.json"):
        spec = gen.load_spec(SKILL / "examples" / name)
        assert spec.links and all(link.kind != "process" for link in spec.links)


# --- Счётчик электроэнергии на вводе ---------------------------------------------------------------------


def _with_input_meter(raw: dict, meter: str = "Меркурий 234") -> dict:
    raw = copy.deepcopy(raw)
    raw["supply"]["meter"] = meter
    return raw


@pytest.fixture(scope="module")
def zhbi_input(tmp_path_factory: pytest.TempPathFactory):
    """Эталон ЖБИ со счётчиком на вводе: две точки поставки, оборудование без приборов, окно показаний 60 суток
    через границу месяцев."""
    folder = tmp_path_factory.mktemp("zhbi-input")
    spec = _load(_with_input_meter(_zhbi_raw()), folder)
    out = folder / "pack"
    gen.build(spec, out)
    return spec, gen.simulate(spec), out


def _number(text: str) -> float:
    return float(text.replace(" ", "").replace(",", "."))


def _readings(out: Path) -> dict[str, list[tuple[str, float, float]]]:
    """Общий файл показаний ЭЭ: прибор → [(время, кВт, показание кВт·ч)] по порядку строк."""
    found: dict[str, list[tuple[str, float, float]]] = defaultdict(list)
    lines = (out / "telemetry" / "показания_приборов_учёта.csv").read_text(encoding="utf-8").splitlines()
    for line in lines[1:]:
        device, ts, kw, kwh = line.split(";")
        found[device].append((ts, _number(kw), _number(kwh)))
    return found


def test_input_series_is_metered_plus_unmetered_plus_other(zhbi_input) -> None:
    spec, sim, out = zhbi_input
    power = gen.input_power(sim)
    assert all(power[i] == pytest.approx(sum(s[i] for s in sim.units.values()) + sim.other[i], abs=1e-9)
               for i in range(len(sim.hours)))
    # По файлу показаний, каждый час окна: ввод = приборы оборудования + оборудование без прибора + база и потери.
    readings = _readings(out)
    vvod = readings.pop(gen.INPUT_DEVICE)
    first = len(sim.hours) - 24 * spec.telemetry_days
    assert len(vvod) == 24 * spec.telemetry_days
    hidden = [f"{e.key}#{n}" for e in spec.equipment if not e.metered for n in range(1, e.count + 1)]
    assert hidden  # аспирация и завеса — без приборов
    for k, (ts, kw, _) in enumerate(vvod):
        i = first + k
        assert ts == sim.hours[i].ts.strftime("%d.%m.%Y %H:%M")
        seen = sum(rows[k][1] for rows in readings.values())
        unseen = sum(sim.units[key][i] for key in hidden) + sim.other[i]
        assert kw == pytest.approx(seen + unseen, abs=0.005 * (len(readings) + 1))
    # Показание — нарастающий итог мощности.
    assert all(b[2] - a[2] == pytest.approx(b[1], abs=0.1) for a, b in pairwise(vvod))


def test_input_month_matches_bill(zhbi_input) -> None:
    spec, sim, out = zhbi_input
    bills = {(b.year, b.month): b.kwh for b in gen.bills(spec, sim)}
    by_month: dict[tuple[int, int], float] = defaultdict(float)
    hours: dict[tuple[int, int], int] = defaultdict(int)
    for hour, kw in zip(sim.hours, gen.input_power(sim)):
        by_month[(hour.ts.year, hour.ts.month)] += kw
        hours[(hour.ts.year, hour.ts.month)] += 1
    assert len(bills) == spec.bill_months
    assert {key: round(kwh) for key, kwh in by_month.items()} == bills
    # По файлу: месяц, все часы которого попали в окно показаний, сходится со счётом с точностью округления
    # кВт до сотых. В окне 60 суток до 13.11 таких месяцев два: октябрь целиком и неполный ноябрь (как и счёт).
    csv_month: dict[tuple[int, int], list[float]] = defaultdict(list)
    for ts, kw, _ in _readings(out)[gen.INPUT_DEVICE]:
        day = datetime.strptime(ts, "%d.%m.%Y %H:%M")
        csv_month[(day.year, day.month)].append(kw)
    full = [key for key, values in csv_month.items() if len(values) == hours[key]]
    assert full == [(2026, 10), (2026, 11)]
    for key in full:
        assert sum(csv_month[key]) == pytest.approx(bills[key], abs=0.005 * hours[key] + 0.5)


def test_input_meter_in_inventory_and_device_list(zhbi_input) -> None:
    spec, _, out = zhbi_input
    point = spec.supply["supply_points"][0]
    node = f"Счётчик на вводе {point}"
    (_, inventory), = _xlsx(out / "upload" / "03_Ведомость_оборудования.xlsx")
    names = [row[0] for row in inventory[4:]]
    # Ввод — первой строкой узлов учёта, перед вводами газа, тепла и воды; без мощности, как они.
    assert names[-4:] == [node, "Узел учёта газа", "Узел учёта тепла (ИТП)", "Водомер на вводе"]
    assert inventory[4:][-4] == [node, "Меркурий 234", 1, None, point]
    (_, devices), = _xlsx(out / "telemetry" / "перечень_приборов_учёта.xlsx")
    assert [gen.INPUT_DEVICE, "Меркурий 234", node, point] in devices
    # Агент относит строку к классу «Ввод» по синониму «счётчик на вводе» подряд (`match_class`), а прибор —
    # к узлу по словам имени (`match_node`): «vvod» есть у синонима «ввод» класса и нет у оборудования.
    assert "счётчик на вводе" in node.lower()
    assert {w for w in re.split(r"[^a-z0-9]+", gen.INPUT_DEVICE) if len(w) >= 3} == {"vvod"}
    assert not any("vvod" in e.key or "ввод" in e.name.lower() for e in spec.equipment)


def test_input_meter_location_from_spec(tmp_path: Path) -> None:
    raw = _with_input_meter(_raw("no-circuits.json"))
    raw["supply"]["meter_location"] = "Электрощитовая"
    spec = _load(raw, tmp_path)
    assert gen.input_meter(spec) == {"device": "ee-vvod-1", "node": f"Счётчик на вводе {gen.supply_point(spec)}",
                                     "model": "Меркурий 234", "location": "Электрощитовая"}


def test_summary_visible_share(zhbi_input) -> None:
    spec, _, out = zhbi_input
    seen = json.loads((out / "summary.json").read_text(encoding="utf-8"))["input_meter"]
    assert seen["device"] == gen.INPUT_DEVICE and seen["days"] == spec.telemetry_days
    assert 0.75 <= seen["visible_share"] <= 0.95
    assert seen["visible_share"] == pytest.approx(seen["metered_kwh"] / seen["input_kwh"], abs=0.001)
    invisible = {item["title"]: item for item in seen["invisible"]}
    assert set(invisible) == {"Аспирация БСУ", "Тепловая завеса ворот формовочного цеха", gen.OTHER_TITLE}
    assert seen["metered_kwh"] + sum(item["kwh"] for item in invisible.values()) == pytest.approx(
        seen["input_kwh"], abs=len(invisible) + 1)
    by_hour = seen["visible_share_by_hour"]
    # Ночью оборудование стоит, а база и потери идут: доля видимого ночью ниже, чем днём.
    assert len(by_hour) == 24 and max(by_hour[0:6]) < min(by_hour[9:17])
    answers = (out / "ЧТО ЗАЛОЖЕНО.md").read_text(encoding="utf-8")
    assert "## Счётчик на вводе" in answers
    assert f"видят {gen._pct(seen['visible_share'])} потребления ввода" in answers
    assert all(f"| {title} |" in answers for title in invisible)
    assert "Счётчик один, на точке поставки из счетов" in answers  # у эталона две точки поставки
    how = (out / "КАК ЗАГРУЗИТЬ.md").read_text(encoding="utf-8")
    assert f"Счётчик на вводе `{gen.INPUT_DEVICE}` — в том же файле" in how
    steps = [int(line.split(".")[0]) for line in how.splitlines() if line[:1].isdigit()]
    assert steps == list(range(1, len(steps) + 1))


@pytest.mark.parametrize("name", ["no-circuits.json", "zhbi-pribrezhny.json"])
def test_no_input_meter_without_field(name: str, tmp_path: Path) -> None:
    spec = _load(_raw(name), tmp_path)
    out = tmp_path / "pack"
    gen.build(spec, out)
    assert gen.INPUT_DEVICE not in _readings(out)
    (_, inventory), = _xlsx(out / "upload" / "03_Ведомость_оборудования.xlsx")
    assert not any(str(row[0]).startswith("Счётчик на вводе") for row in inventory)
    (_, devices), = _xlsx(out / "telemetry" / "перечень_приборов_учёта.xlsx")
    assert not any(row[0] == gen.INPUT_DEVICE for row in devices)
    assert "input_meter" not in json.loads((out / "summary.json").read_text(encoding="utf-8"))
    assert "## Счётчик на вводе" not in (out / "ЧТО ЗАЛОЖЕНО.md").read_text(encoding="utf-8")
    assert gen.INPUT_DEVICE not in (out / "КАК ЗАГРУЗИТЬ.md").read_text(encoding="utf-8")


def test_input_meter_keeps_equipment_readings(zhbi_input, tmp_path: Path) -> None:
    """Счётчик на вводе дописывается в конец файла: показания приборов оборудования те же, что без него."""
    _, _, out = zhbi_input
    spec = _load(_zhbi_raw(), tmp_path)
    gen.build(spec, tmp_path / "pack")
    without = _readings(tmp_path / "pack")
    with_input = _readings(out)
    assert list(with_input)[-1] == gen.INPUT_DEVICE
    with_input.pop(gen.INPUT_DEVICE)
    assert with_input == without


@pytest.mark.parametrize(("meter", "message"), [("", "марка счётчика на вводе строкой"),
                                               (234, "марка счётчика на вводе строкой"),
                                               ("Корректор ЕК-270", "без «Ирвис», «корректор», «архив»")])
def test_input_meter_spec_errors(meter, message: str, tmp_path: Path) -> None:
    with pytest.raises(gen.SpecError, match=message):
        _load(_with_input_meter(_raw("no-circuits.json"), meter), tmp_path)
