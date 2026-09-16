# -*- coding: utf-8 -*-
"""Проверки выгрузок сводки: таблица «Техника под контролем».

По просьбе заказчика разбивка «всего / исправны / неисправны» по ИБП,
стабилизаторам напряжения, усилителям сотовой связи и роутерам Huawei нужна
в СКАЧИВАЕМЫХ отчётах (HTML-дашборд и Excel), а не в сводке на экране.

Запуск:  python test_dashboard.py   (exit 0 — все проверки зелёные)

Часть проверок — прямые вызовы функций приложения (импорт в «холостом» режиме
Streamlit, без интерфейса); часть — через тест-движок, что блока нет на экране.
"""
import os
import sys
import tempfile
from pathlib import Path

import pandas as pd

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Базу подменяем ДО импорта приложения: оно настраивается на неё при загрузке.
os.environ["DATABASE_URL"] = "sqlite:///" + (
    Path(tempfile.mkdtemp()) / "test_inventory.db"
).as_posix()

import app  # noqa: E402  (импорт после подмены базы)
from test_support import add_item, app_run, has_text, login_admin, new_db_path, prepare_db  # noqa: E402

PASSED = []


def check(name, cond, extra=""):
    PASSED.append(bool(cond))
    print(("OK  " if cond else "FAIL") + f" {name}" + (f" | {extra}" if extra else ""))


def sample_data():
    """Техника для проверки: ноутбуки 3/1/2, роутеры 2/1/1, усилители 3/2/1,
    стабилизаторы 1/1/0, ИБП 3/2/1 (одна строка с количеством 2).

    Колонки — как в базе: HTML-дашборд читает весь набор полей позиции.
    """
    rows = [
        ("Ноутбук", 1, "Удовлетворительно"),
        ("Ноутбук", 1, "Не удовлетворительно: не включается"),
        ("Ноутбук", 1, "Не удовлетворительно: разбит экран"),
        ("Роутер Huawei", 1, "Удовлетворительно"),
        ("Роутер Huawei", 1, "Не удовлетворительно: не ловит сеть"),
        ("Усилитель сотовой связи", 1, "Удовлетворительно"),
        ("Усилитель сотовой связи", 1, "Удовлетворительно"),
        ("Усилитель сотовой связи", 1, "Не удовлетворительно: нет кабеля"),
        ("Стабилизатор напряжения", 1, "Удовлетворительно"),
        ("ИБП", 2, "Удовлетворительно"),
        ("ИБП", 1, "Не удовлетворительно: не держит заряд"),
    ]
    return pd.DataFrame(
        [
            {
                "id": index + 1,
                "party": "Партия № 1",
                "category": category,
                "model": f"{category}-{index + 1}",
                "serial_number": f"SN-{index + 1}",
                "inv_number": f"INV-{index + 1}",
                "quantity": quantity,
                "condition": condition,
                "engineer": "Иванов",
                "date_updated": "тест",
            }
            for index, (category, quantity, condition) in enumerate(rows)
        ]
    )


data = sample_data()

print("=== 1. Расчёт по категориям ===")
check("ноутбуки 3/1/2", app.category_summary(data, "Ноутбук") == {
    "total": 3, "ok": 1, "bad": 2}, str(app.category_summary(data, "Ноутбук")))
check("роутеры Huawei 2/1/1",
      app.category_summary(data, "Роутер Huawei") == {"total": 2, "ok": 1, "bad": 1})
check("усилители сотовой связи 3/2/1",
      app.category_summary(data, "Усилитель сотовой связи") == {
          "total": 3, "ok": 2, "bad": 1})
check("стабилизаторы напряжения 1/1/0",
      app.category_summary(data, "Стабилизатор напряжения") == {
          "total": 1, "ok": 1, "bad": 0})
check("ИБП 3/2/1 (строка с количеством 2 считается дважды)",
      app.category_summary(data, "ИБП") == {"total": 3, "ok": 2, "bad": 1},
      str(app.category_summary(data, "ИБП")))
check("категория без техники даёт нули",
      app.category_summary(data, "Сотовый телефон") == {"total": 0, "ok": 0, "bad": 0})

print("=== 2. Таблица для выгрузок ===")
table = app.control_summary_table(data)
check("колонки: Категория / Всего / Исправны / Неисправны",
      list(table.columns) == ["Категория", "Всего", "Исправны", "Неисправны"],
      str(list(table.columns)))
values = {
    row["Категория"]: (row["Всего"], row["Исправны"], row["Неисправны"])
    for _, row in table.iterrows()
}
check("в таблице ровно четыре категории", len(values) == 4, str(sorted(values)))
check("роутеры Huawei 2/1/1", values.get("Роутер Huawei") == (2, 1, 1))
check("усилители сотовой связи 3/2/1",
      values.get("Усилитель сотовой связи") == (3, 2, 1))
check("стабилизаторы напряжения 1/1/0",
      values.get("Стабилизатор напряжения") == (1, 1, 0))
check("ИБП 3/2/1", values.get("ИБП") == (3, 2, 1))

print("=== 3. HTML-дашборд (скачиваемый) содержит таблицу ===")
html = app.build_dashboard_html(data)
check("в HTML есть раздел «Техника под контролем»", "Техника под контролем" in html)
check("в HTML есть все четыре категории",
      all(category in html for category in
          ("Роутер Huawei", "Усилитель сотовой связи",
           "Стабилизатор напряжения", "ИБП")))
check("в HTML есть строка роутеров с числами 2/1/1",
      "<td>2</td><td>1</td><td>1</td>" in html.replace(" ", ""))

print("=== 4. На экране сводки этих карточек нет (по просьбе заказчика) ===")
db = new_db_path()
prepare_db(db)
add_item(db, "Партия № 1", "ИБП", model="UPS-1", serial="I1")
os.environ["ADMIN_PASSWORD"] = "тест-админ"
at = login_admin(app_run(db), "тест-админ")
check("сводка открылась", not at.exception,
      "; ".join(str(e.value)[:120] for e in at.exception))
check("карточек «Всего ИБП» на экране нет", not has_text(at, "Всего ИБП"))
check("подсказка про выгрузку на месте",
      has_text(at, "в скачиваемых отчётах"))
del os.environ["ADMIN_PASSWORD"]

print()
print("ИТОГ:", sum(PASSED), "/", len(PASSED),
      "проверок OK" if all(PASSED) else "— ЕСТЬ ПРОВАЛЫ")
sys.exit(0 if all(PASSED) else 1)
