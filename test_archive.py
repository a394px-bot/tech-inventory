# -*- coding: utf-8 -*-
"""Проверки списания в архив вместо безвозвратного удаления.

Правило: инженер не может стереть позицию — он списывает её в архив (запись,
кто и когда виден в журнале). Списанная техника уходит из списков, отчётов и
не считается в лимитах. Совсем удалять может только администратор.

Запуск:  python test_archive.py   (exit 0 — все проверки зелёные)
"""
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from test_support import (  # noqa: E402
    ENGINEER,
    PARTY,
    add_item,
    app_run,
    dataframe_contains,
    dataframe_with_marker_contains,
    db_execute,
    db_query,
    find,
    find_button,
    has_text,
    login_admin,
    metric_value,
    new_db_path,
    prepare_db,
)

PASSED = []


def check(name, cond, extra=""):
    PASSED.append(bool(cond))
    print(("OK  " if cond else "FAIL") + f" {name}" + (f" | {extra}" if extra else ""))


print("=== 1. У инженера нет удаления, есть списание в архив ===")
db = new_db_path()
params = prepare_db(db)
add_item(db, PARTY, "Усилитель сотовой связи", model="УС-архив", serial="SN-1")
at = app_run(db, params)
check("кабинет открыт", has_text(at, "Работает:"),
      "; ".join(str(e.value)[:80] for e in at.exception))
check("кнопки «Удалить позицию» у инженера нет",
      find_button(at, "Удалить позицию") is None)
check("есть кнопка «Списать в архив»",
      find_button(at, "Списать в архив") is not None,
      "; ".join(b.label for b in at.button))

print("=== 2. Без подтверждения списание не проходит ===")
find_button(at, "Списать в архив").click()
at.run()
archived = db_query(db, "SELECT archived FROM equipment WHERE serial_number='SN-1'")[0][0]
check("без галочки позиция не списана",
      has_text(at, "Подтверждаю списание") and not archived, f"archived={archived}")

print("=== 3. С подтверждением — списывается и подписывается ===")
checkbox = find("checkbox", at, "Подтверждаю списание")
check("галочка подтверждения на месте", checkbox is not None)
if checkbox is not None:
    checkbox.check()
find_button(at, "Списать в архив").click()
at.run()
row = db_query(
    db,
    "SELECT archived, archived_at, archived_by FROM equipment "
    "WHERE serial_number='SN-1'",
)[0]
check("позиция помечена списанной", row[0] == 1, str(row))
check("записано, кто списал", row[2] == ENGINEER, str(row[2]))
check("записано, когда списали", bool(row[1]), str(row[1]))
actions = [r[0] for r in db_query(db, "SELECT action FROM audit_log")]
check("в журнале есть запись о списании", "Списание" in actions, str(actions))
check("запись в базе не удалена", len(db_query(
    db, "SELECT id FROM equipment WHERE serial_number='SN-1'")) == 1)
check("списанная позиция исчезла из списка инженера",
      not dataframe_contains(at, "SN-1"))
check("инженер видит, что техники нет",
      has_text(at, "нет зарегистрированной техники") or not at.dataframe,
      "; ".join(str(v.value)[:60] for v in at.info))

print("=== 4. Списание не мешает заново зарегистрировать тот же серийник ===")
add_item(db, PARTY, "Усилитель сотовой связи", model="УС-новый", serial="SN-1")
at2 = app_run(db, params)
check("новая позиция с тем же серийником видна инженеру",
      dataframe_contains(at2, "УС-новый"))
check("списанная позиция в списке не мешается",
      not dataframe_contains(at2, "УС-архив"))

print("=== 5. Карточка позиции показывает статус ===")
item_id = db_query(db, "SELECT id FROM equipment WHERE serial_number='SN-1' AND archived=1")[0][0]
at_card = app_run(db, {"item": str(item_id)})
check("в карточке видно, что позиция списана",
      has_text(at_card, "Списана в архив") and has_text(at_card, ENGINEER),
      "; ".join(str(v.value)[:60] for v in at_card.markdown)[:100])

print("=== 6. Реестр администратора: списанной позиции нет ===")
os.environ["ADMIN_PASSWORD"] = "тест-админ"
at_a = login_admin(app_run(db), "тест-админ")
check("админ вошёл (сводка открылась)",
      metric_value(at_a, "Всего единиц техники") is not None and not at_a.exception,
      "; ".join(str(e.value)[:80] for e in at_a.exception))
check("в реестре есть действующая позиция", dataframe_contains(at_a, "УС-новый"))
check("списанной позиции в РЕЕСТРЕ нет",
      not dataframe_with_marker_contains(at_a, "Обновлено", "УС-архив"))
check("но в журнале действий запись о списании осталась",
      dataframe_with_marker_contains(at_a, "Действие", "УС-архив"))

print("=== 7. Списание не мешает работать остальной технике ===")
db2 = new_db_path()
params2 = prepare_db(db2)
for index in range(3):  # ноутбуков 3 при лимите 2 — превышение
    add_item(db2, PARTY, "Ноутбук", model=f"Lenovo-{index}", serial=f"L{index}")
at_b = login_admin(app_run(db2), "тест-админ")
check("превышение лимита видно", metric_value(at_b, "Превышений лимита") == "1",
      str(metric_value(at_b, "Превышений лимита")))
db_execute(db2, "UPDATE equipment SET archived=1 WHERE serial_number='L0'")
at_b2 = login_admin(app_run(db2), "тест-админ")
check("списанный ноутбук в лимите не считается",
      metric_value(at_b2, "Превышений лимита") == "0",
      str(metric_value(at_b2, "Превышений лимита")))
check("остальная техника на месте", metric_value(at_b2, "Всего единиц техники") == "2",
      str(metric_value(at_b2, "Всего единиц техники")))

print("=== 8. Старая база: колонки архива добавляются, позиции остаются в строю ===")
db3 = new_db_path()
params3 = prepare_db(db3)
add_item(db3, PARTY, "МФУ/Принтер", model="HP-1", serial="OLD-2")
at_c = app_run(db3, params3)
columns = {row[1] for row in db_query(db3, "PRAGMA table_info(equipment)")}
check("в базе появились колонки архива",
      {"archived", "archived_at", "archived_by"} <= columns, str(sorted(columns)))
state = db_query(db3, "SELECT archived FROM equipment WHERE serial_number='OLD-2'")[0][0]
check("старая позиция помечена как «в строю» (не потерялась)", state == 0, str(state))
check("старая позиция видна инженеру", dataframe_contains(at_c, "OLD-2"))
del os.environ["ADMIN_PASSWORD"]

print()
print("ИТОГ:", sum(PASSED), "/", len(PASSED),
      "проверок OK" if all(PASSED) else "— ЕСТЬ ПРОВАЛЫ")
sys.exit(0 if all(PASSED) else 1)
