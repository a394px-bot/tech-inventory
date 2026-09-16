# -*- coding: utf-8 -*-
"""Проверки подстановки полей из базы: только ноутбуки.

Правила, принятые заказчиком:
- серийный и инвентарный номер (и модель) подтягиваются из справочника ТОЛЬКО для
  категории «Ноутбук»; для остальной техники таких подстановок нет;
- при выборе другого ноутбука поля должны ЗАМЕНИТЬСЯ на его данные (раньше оставался
  серийник другой техники);
- после сохранения позиции форма очищается, включая поиск по базе, — иначе данные
  внесённого ноутбука подставлялись бы при вводе следующей позиции.

Запуск:  python test_laptop_fill.py   (exit 0 — все проверки зелёные)
"""
import os
import subprocess
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from test_support import (  # noqa: E402
    add_laptop_reference,
    app_run,
    db_query,
    find,
    find_button,
    has_text,
    new_db_path,
    prepare_db,
)

HERE = Path(__file__).parent
CHILD_SCENARIO = "--scenario-save-laptop"
PASSED = []


def check(name, cond, extra=""):
    PASSED.append(bool(cond))
    print(("OK  " if cond else "FAIL") + f" {name}" + (f" | {extra}" if extra else ""))


def field_value(at, label_part):
    widget = find("text_input", at, label_part)
    return str(widget.value) if widget is not None else None


def scenario_save_laptop():
    """Отдельный процесс: сохраняем ноутбук и смотрим, очистилась ли форма."""
    db = new_db_path()
    params = prepare_db(db)
    add_laptop_reference(db, "Lenovo ThinkPad", "PF-SAVE-1", "INV-SAVE-1")
    at = app_run(db, params)
    find("selectbox", at, "Поиск по инвентарному номеру").select("INV-SAVE-1")
    at.run()
    find_button(at, "Сохранить позицию").click()
    at.run()
    saved = db_query(
        db, "SELECT model, serial_number, inv_number FROM equipment "
            "WHERE serial_number='PF-SAVE-1'"
    )
    check("ноутбук сохранён с подставленными данными",
          saved == [("Lenovo ThinkPad", "PF-SAVE-1", "INV-SAVE-1")], str(saved))
    check("показано подтверждение с новой формулировкой",
          has_text(at, "категория сброшена"),
          "; ".join(str(v.value)[:90] for v in at.success))
    search = find("selectbox", at, "Поиск по инвентарному номеру")
    check("поиск по базе очищен", search is not None and str(search.value) == "",
          f"значение: {getattr(search, 'value', None)!r}")
    # Очистку полей проверяем по поведению приложения: тест-движок хранит свою
    # копию введённого значения и показывает её даже после очистки, а приложение
    # видит поля пустыми — что и подтверждает следующий шаг.
    find_button(at, "Сохранить позицию").click()
    at.run()
    again = db_query(
        db, "SELECT COUNT(*) FROM equipment WHERE serial_number='PF-SAVE-1'"
    )[0][0]
    check("повторное сохранение не создаёт дубль", again == 1, f"строк: {again}")
    check("после сохранения поля пустые — приложение просит их заполнить",
          has_text(at, "хотя бы одно поле"),
          "; ".join(str(v.value)[:90] for v in at.error))
    print()
    print("ИТОГ:", sum(PASSED), "/", len(PASSED),
          "проверок OK" if all(PASSED) else "— ЕСТЬ ПРОВАЛЫ")
    sys.exit(0 if all(PASSED) else 1)


if len(sys.argv) > 1 and sys.argv[1] == CHILD_SCENARIO:
    scenario_save_laptop()


db = new_db_path()
params = prepare_db(db)
add_laptop_reference(db, "Lenovo ThinkPad", "PF-111", "INV-111")
add_laptop_reference(db, "HP ProBook", "PF-222", "INV-222")

print("=== 1. Подстановка при выборе ноутбука из базы ===")
at = app_run(db, params)
check("кабинет открыт", has_text(at, "Работает:"),
      "; ".join(str(e.value)[:80] for e in at.exception))
check("по умолчанию выбрана категория «Ноутбук»",
      find("selectbox", at, "Категория техники") is not None
      and str(find("selectbox", at, "Категория техники").value) == "Ноутбук",
      str(find("selectbox", at, "Категория техники").value))
search = find("selectbox", at, "Поиск по инвентарному номеру")
check("есть поиск по инвентарному номеру", search is not None)
search.select("INV-111")
at.run()
check("подставлена модель первого ноутбука",
      field_value(at, "Модель ноутбука") == "Lenovo ThinkPad",
      str(field_value(at, "Модель ноутбука")))
check("подставлен серийный номер первого ноутбука",
      field_value(at, "Серийный номер") == "PF-111",
      str(field_value(at, "Серийный номер")))
check("подставлен инвентарный номер первого ноутбука",
      field_value(at, "Инвентарный номер") == "INV-111",
      str(field_value(at, "Инвентарный номер")))

print("=== 2. Выбор другого ноутбука заменяет данные (это был глюк) ===")
find("selectbox", at, "Поиск по инвентарному номеру").select("INV-222")
at.run()
check("модель заменилась на второй ноутбук",
      field_value(at, "Модель ноутбука") == "HP ProBook",
      str(field_value(at, "Модель ноутбука")))
check("серийник заменился (в поле не осталось чужого)",
      field_value(at, "Серийный номер") == "PF-222",
      str(field_value(at, "Серийный номер")))
check("инвентарный заменился", field_value(at, "Инвентарный номер") == "INV-222",
      str(field_value(at, "Инвентарный номер")))

print("=== 3. Поиск по серийному номеру тоже работает ===")
find("selectbox", at, "Поиск по инвентарному номеру").select("")
at.run()
find("selectbox", at, "Или поиск по серийному номеру").select("PF-111")
at.run()
check("по серийнику подставились данные первого ноутбука",
      field_value(at, "Серийный номер") == "PF-111"
      and field_value(at, "Инвентарный номер") == "INV-111",
      f"{field_value(at, 'Серийный номер')} / {field_value(at, 'Инвентарный номер')}")

print("=== 4. Для остальной техники подстановки нет ===")
find("selectbox", at, "Категория техники").select("МФУ/Принтер")
at.run()
check("нет поиска по базе для МФУ/Принтера",
      find("selectbox", at, "Поиск по инвентарному номеру") is None
      and find("selectbox", at, "Или поиск по серийному номеру") is None)

print("=== 5. Сохранение ноутбука: форма очищается (отдельный процесс) ===")
child = subprocess.run(
    [sys.executable, str(Path(__file__).resolve()), CHILD_SCENARIO],
    capture_output=True, text=True, encoding="utf-8", errors="replace",
    env={**os.environ, "PYTHONIOENCODING": "utf-8"}, cwd=str(HERE),
)
child_lines = [line for line in (child.stdout or "").splitlines() if line.strip()]
for line in child_lines:
    if line.startswith(("OK", "FAIL")):
        check("(отдельно) " + line[4:], line.startswith("OK"), line[4:])
check("сценарий сохранения ноутбука прошёл", child.returncode == 0,
      child_lines[-1] if child_lines else (child.stderr or "")[-200:])

print()
print("ИТОГ:", sum(PASSED), "/", len(PASSED),
      "проверок OK" if all(PASSED) else "— ЕСТЬ ПРОВАЛЫ")
sys.exit(0 if all(PASSED) else 1)
