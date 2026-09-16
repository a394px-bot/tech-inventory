# -*- coding: utf-8 -*-
"""Проверки сводки администратора: карточки «всего / исправны / неисправны».

По каждой категории под контролем (ноутбуки, роутеры Huawei, усилители сотовой
связи, стабилизаторы напряжения, ИБП) в сводке должно быть три числа:
сколько всего единиц, сколько исправных и сколько неисправных.

Запуск:  python test_dashboard.py   (exit 0 — все проверки зелёные)
"""
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from test_support import (  # noqa: E402
    PARTY,
    add_item,
    app_run,
    category_stats,
    has_text,
    login_admin,
    new_db_path,
    prepare_db,
)

PASSED = []


def check(name, cond, extra=""):
    PASSED.append(bool(cond))
    print(("OK  " if cond else "FAIL") + f" {name}" + (f" | {extra}" if extra else ""))


print("=== Готовим технику разных категорий ===")
db = new_db_path()
prepare_db(db)
# Ноутбуки: 3 позиции — 1 исправен, 2 неисправны
add_item(db, PARTY, "Ноутбук", model="Lap-ok", serial="L1")
add_item(db, PARTY, "Ноутбук", model="Lap-bad-1", serial="L2",
         condition="Не удовлетворительно: не включается")
add_item(db, PARTY, "Ноутбук", model="Lap-bad-2", serial="L3",
         condition="Не удовлетворительно: разбит экран")
# Роутеры Huawei: 2 — 1 исправен, 1 неисправен
add_item(db, PARTY, "Роутер Huawei", model="Роутер Huawei", serial="R1")
add_item(db, PARTY, "Роутер Huawei", model="Роутер Huawei", serial="R2",
         condition="Не удовлетворительно: не ловит сеть")
# Усилители сотовой связи: 3 — 2 исправны, 1 неисправен
add_item(db, PARTY, "Усилитель сотовой связи", model="УС-1", serial="U1")
add_item(db, PARTY, "Усилитель сотовой связи", model="УС-2", serial="U2")
add_item(db, PARTY, "Усилитель сотовой связи", model="УС-3", serial="U3",
         condition="Не удовлетворительно: нет кабеля питания")
# Стабилизатор напряжения: 1 исправный
add_item(db, PARTY, "Стабилизатор напряжения", model="Stab-1", serial="S1")
# ИБП: строка с количеством 2 (исправные) + 1 неисправный = 3 единицы
add_item(db, PARTY, "ИБП", model="UPS-1", serial="I1", quantity=2)
add_item(db, PARTY, "ИБП", model="UPS-2", serial="I2",
         condition="Не удовлетворительно: не держит заряд")

os.environ["ADMIN_PASSWORD"] = "тест-админ"
at = login_admin(app_run(db), "тест-админ")
check("сводка открылась", not at.exception,
      "; ".join(str(e.value)[:120] for e in at.exception))
check("блок «Техника под контролем» есть", has_text(at, "Техника под контролем"))

print("=== Ноутбуки (как было) ===")
check("ноутбуки: всего 3, исправны 1, неисправны 2",
      category_stats(at, "Всего ноутбуков") == (3, 1, 2),
      str(category_stats(at, "Всего ноутбуков")))

print("=== Новые категории в сводке ===")
check("роутеры Huawei: всего 2, исправны 1, неисправны 1",
      category_stats(at, "Всего роутеров Huawei") == (2, 1, 1),
      str(category_stats(at, "Всего роутеров Huawei")))
check("усилители сотовой связи: всего 3, исправны 2, неисправны 1",
      category_stats(at, "Всего усилителей сотовой связи") == (3, 2, 1),
      str(category_stats(at, "Всего усилителей сотовой связи")))
check("стабилизаторы напряжения: всего 1, исправен 1, неисправных 0",
      category_stats(at, "Всего стабилизаторов напряжения") == (1, 1, 0),
      str(category_stats(at, "Всего стабилизаторов напряжения")))
check("ИБП: 3 единицы (одна строка с количеством 2), исправны 2, неисправен 1",
      category_stats(at, "Всего ИБП") == (3, 2, 1),
      str(category_stats(at, "Всего ИБП")))

print("=== Списанное в архив в сводке не считается ===")
from test_support import db_execute  # noqa: E402

db_execute(db, "UPDATE equipment SET archived = 1 WHERE serial_number = 'R1'")
at2 = login_admin(app_run(db), "тест-админ")
check("списанный роутер убран из карточки",
      category_stats(at2, "Всего роутеров Huawei") == (1, 0, 1),
      str(category_stats(at2, "Всего роутеров Huawei")))
del os.environ["ADMIN_PASSWORD"]

print()
print("ИТОГ:", sum(PASSED), "/", len(PASSED),
      "проверок OK" if all(PASSED) else "— ЕСТЬ ПРОВАЛЫ")
sys.exit(0 if all(PASSED) else 1)
