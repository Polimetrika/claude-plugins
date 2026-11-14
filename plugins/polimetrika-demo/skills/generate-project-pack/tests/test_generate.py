"""Тесты генератора демо-пакета: совместимость без участков и выпуск, график и ресурсы по участкам.

Запуск из корня репозитория: `pytest plugins/polimetrika-demo/skills/generate-project-pack/tests`.
Нужны openpyxl, python-docx; reportlab — по желанию (без него счета выходят в DOCX, на тесты это не влияет).
"""

from __future__ import annotations

import copy
import hashlib
import importlib.util
import json
import os
import sys
from collections import defaultdict
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


#: `GENERATE_PY` — другой файл генератора: так считались эталонные отпечатки ниже (генератор 0.2.1).
gen = _module(Path(os.environ.get("GENERATE_PY", SKILL / "scripts" / "generate.py")))

#: Отпечатки пакета генератора 0.2.1 (до участков) на `fixtures/no-circuits.json`: без контуров и с одним
#: контуром пакет не меняется ни в одной цифре и ни в одной строке.
GOLDEN = {
    "no_circuits": "6126036ef978a99c1c0f16195994244998ff77eb78f25d76c40e94a4cd043f0e",
    "one_circuit": "b342d00a9620099863160e828f350b6b5b804e0dd81b73244cf877145ed6f987",
}


def _raw(name: str) -> dict:
    return json.loads((FIXTURES / name).read_text(encoding="utf-8"))


def _one_circuit(raw: dict) -> dict:
    raw = copy.deepcopy(raw)
    raw["circuits"] = [{"key": "roasting", "title": "Обжарка", "driver_unit": "кг"}]
    for item in raw["equipment"]:
        if item["key"] in ("roaster", "exhaust"):
            item["circuit"] = "roasting"
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


def digest(out: Path) -> str:
    """Отпечаток содержимого пакета без того, что зависит от машины: пути архива, шрифта PDF, метаданных
    DOCX/XLSX. Счета — по `summary.json` (кВт·ч, суммы), документы — по тексту и ячейкам."""
    summary = json.loads((out / "summary.json").read_text(encoding="utf-8"))
    for key in ("upload_zip", "pdf_font", "upload_files"):
        summary.pop(key)
    parts: list = [summary]
    for path in sorted(out.rglob("*")):
        if path.suffix in (".csv", ".md"):
            parts.append((path.name, path.read_text(encoding="utf-8")))
        elif path.suffix == ".xlsx":
            parts.append((path.name, _xlsx(path)))
        elif path.suffix == ".docx" and path.parent.name == "upload":
            parts.append((path.name, _docx(path)))
    text = json.dumps(parts, ensure_ascii=False, default=str, sort_keys=True)
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


# --- Совместимость ---------------------------------------------------------------------------------


@pytest.mark.parametrize("case", ["no_circuits", "one_circuit"])
def test_pack_without_areas_is_unchanged(case: str, tmp_path: Path) -> None:
    raw = _raw("no-circuits.json")
    spec = _load(raw if case == "no_circuits" else _one_circuit(raw), tmp_path)
    gen.build(spec, tmp_path / "pack")
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
    """Выпуск, разобранный раньше контуров модели, агент складывает в объект: при участках отчёты о выпуске
    лежат отдельно и прикладываются вторым сообщением, после подтверждения модели."""
    _, _, out = coffee
    later = out / "upload" / gen.PRODUCTION_LATER
    assert sorted(p.name for p in later.iterdir()) == ["05_Отчёт_о_выпуске_по_месяцам.xlsx", "06_Суточный_выпуск.xlsx"]
    assert not list((out / "upload").glob("0[56]_*"))
    text = (out / "КАК ЗАГРУЗИТЬ.md").read_text(encoding="utf-8")
    first = text.index(f"КРОМЕ папки `upload/{gen.PRODUCTION_LATER}/`")
    confirm = text.index("энергомодель с контурами «Обжарка», «Фасовка»")
    # Агент сам не относит узлы к контурам: без этого шага удельный по контурам пуст (стенд 14.11).
    phrase = "«Отнеси оборудование к контурам по описанию объекта: что на каком участке работает»"
    equipment = text.index(f"\n   {phrase}.\n")  # фраза целиком на своей строке
    second = text.index("«Вот отчёты о выпуске по участкам, разнеси выпуск по контурам»")
    assert first < confirm < equipment < second
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
