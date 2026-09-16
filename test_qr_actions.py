# -*- coding: utf-8 -*-
"""Проверки действий по QR-коду: подтвердить нахождение или переместить технику.

Правила, согласованные с заказчиком:
- QR-код ведёт в карточку позиции по прежнему адресу (?item=ID) — сами коды не меняем;
- в карточке можно подтвердить, что техника находится в этой партии, или переместить её
  в другую партию / на «Базу»;
- подтвердить может партия, где техника сейчас, или администратор;
- другая партия может принять технику к себе;
- администратор может подтвердить и переместить в любую партию;
- любое действие выполняется по паролю и подписывается фамилией.

Запуск:  python test_qr_actions.py   (exit 0 — все проверки зелёные)
"""
import os
import sys
import tempfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Базу подменяем ДО импорта приложения (оно читает настройки при загрузке).
os.environ["DATABASE_URL"] = "sqlite:///" + (
    Path(tempfile.mkdtemp()) / "test_inventory.db"
).as_posix()

import app  # noqa: E402
from test_support import (  # noqa: E402
    PARTY,
    PASSWORD,
    add_item,
    add_party_password,
    app_run,
    db_query,
    find,
    find_button,
    has_text,
    new_db_path,
    prepare_db,
)

PARTY_B = "Партия № 2"
PASSWORD_B = "Второй-пароль"
PASSED = []


def check(name, cond, extra=""):
    PASSED.append(bool(cond))
    print(("OK  " if cond else "FAIL") + f" {name}" + (f" | {extra}" if extra else ""))


print("=== 1. Адрес в QR-коде не изменился ===")
app.app_base_url = lambda: "https://пример.рф"
check("QR ведёт на карточку позиции (?item=ID)",
      app.item_qr_url(42) == "https://пример.рф/?item=42", app.item_qr_url(42))


def open_card(db, item_id, params=None):
    return app_run(db, {"item": str(item_id)}, params=params) if params else app_run(
        db, {"item": str(item_id)}
    )


def prepare_item():
    db = new_db_path()
    prepare_db(db)
    add_party_password(db, PARTY_B, PASSWORD_B)
    add_item(db, PARTY, "МФУ/Принтер", model="HP-проверка", serial="QR-1")
    item_id = db_query(db, "SELECT id FROM equipment WHERE serial_number='QR-1'")[0][0]
    return db, item_id


def fill_and_go(at, actor, password, name, what=None, destination=None, receiver=None):
    find("selectbox", at, "Кто выполняет действие").select(actor)
    at.run()
    if what:
        find("radio", at, "Что сделать").set_value(what)
        at.run()
    if destination:
        find("selectbox", at, "Куда передать" if "передать" in " ".join(
            str(s.label) for s in at.selectbox
        ) else "Куда переместить").select(destination)
    if receiver:
        find("text_input", at, "Фамилия принимающего").set_value(receiver)
    find("text_input", at, "Ваша фамилия").set_value(name)
    find("text_input", at, "Пароль").set_value(password)
    find_button(at, "Выполнить").click()
    at.run()
    return at


print("=== 2. Карточка по QR открывается и предлагает действия ===")
db, item_id = prepare_item()
at = open_card(db, item_id)
check("карточка открылась", has_text(at, "Карточка позиции"),
      "; ".join(str(e.value)[:80] for e in at.exception))
check("есть блок действий", has_text(at, "Подтвердить или переместить"))
check("видно, где техника сейчас", has_text(at, f"числится в партии «{PARTY}»"))

print("=== 3. Действие без пароля не проходит ===")
at = fill_and_go(at, PARTY, "", "Иванов")
check("без пароля отказ", has_text(at, "Введите пароль"),
      "; ".join(str(v.value)[:90] for v in at.error))
check("подтверждения в базе нет",
      db_query(db, "SELECT confirmed_at FROM equipment WHERE id=?", (item_id,))[0][0] is None)

print("=== 4. Неверный пароль не проходит ===")
at = fill_and_go(at, PARTY, "не тот пароль", "Иванов")
check("неверный пароль — отказ", has_text(at, "Неверный пароль партии"),
      "; ".join(str(v.value)[:90] for v in at.error))

print("=== 5. Своя партия подтверждает нахождение ===")
at = fill_and_go(at, PARTY, PASSWORD, "Иванов")
row = db_query(db, "SELECT confirmed_at, confirmed_by FROM equipment WHERE id=?",
               (item_id,))[0]
check("подтверждение записано с фамилией", bool(row[0]) and row[1] == "Иванов", str(row))
actions = [r[0] for r in db_query(db, "SELECT action FROM audit_log")]
check("в журнале запись «Подтверждение»", "Подтверждение" in actions, str(actions))
at2 = open_card(db, item_id)
check("в карточке видно, кто и когда подтвердил", has_text(at2, "Иванов"))

print("=== 6. Чужая партия подтвердить не может ===")
db2, item_id2 = prepare_item()
at = open_card(db2, item_id2)
at = fill_and_go(at, PARTY_B, PASSWORD_B, "Петров")
check("чужой партии подтверждение запрещено и предложено принять",
      has_text(at, "может та партия, где техника сейчас"),
      "; ".join(str(v.value)[:90] for v in at.error))
check("подтверждения не появилось",
      db_query(db2, "SELECT confirmed_at FROM equipment WHERE id=?",
               (item_id2,))[0][0] is None)

print("=== 7. Другая партия принимает технику к себе ===")
at = fill_and_go(at, PARTY_B, PASSWORD_B, "Петров", what="Переместить",
                 receiver="Сидоров")
party_now = db_query(db2, "SELECT party FROM equipment WHERE id=?", (item_id2,))[0][0]
check("позиция перешла в партию принявшего", party_now == PARTY_B, str(party_now))
history = db_query(
    db2, "SELECT from_where, to_where, engineer, receiver FROM history WHERE equipment_id=?",
    (item_id2,),
)
check("в истории перемещений записаны откуда/куда и кто",
      history == [(PARTY, PARTY_B, "Петров", "Сидоров")], str(history))
check("в журнале действий есть «Перемещение»",
      "Перемещение" in [r[0] for r in db_query(db2, "SELECT action FROM audit_log")])

print("=== 8. Своя партия передаёт технику на Базу ===")
db3, item_id3 = prepare_item()
at = open_card(db3, item_id3)
at = fill_and_go(at, PARTY, PASSWORD, "Иванов", what="Переместить",
                 destination="База", receiver="Кладовщик")
party_now = db_query(db3, "SELECT party FROM equipment WHERE id=?", (item_id3,))[0][0]
check("техника ушла на Базу", party_now == "База", str(party_now))

print("=== 9. Администратор может переместить в любую партию ===")
db4, item_id4 = prepare_item()
os.environ["ADMIN_PASSWORD"] = "админ-пароль"
at = open_card(db4, item_id4)
at = fill_and_go(at, "Администратор", "админ-пароль", "Альфия",
                 what="Переместить", destination=PARTY_B, receiver="Петров")
party_now = db_query(db4, "SELECT party FROM equipment WHERE id=?", (item_id4,))[0][0]
check("админ переместил технику", party_now == PARTY_B, str(party_now))
check("запись в журнале подписана администратором",
      db_query(db4, "SELECT engineer FROM audit_log ORDER BY id DESC LIMIT 1")[0][0]
      == "Альфия")

print("=== 10. Администратор с неверным паролем — отказ ===")
db5, item_id5 = prepare_item()
at = open_card(db5, item_id5)
at = fill_and_go(at, "Администратор", "неверный", "Альфия",
                 what="Переместить", destination=PARTY_B, receiver="Петров")
check("неверный пароль администратора — отказ",
      has_text(at, "Неверный пароль администратора"),
      "; ".join(str(v.value)[:90] for v in at.error))
check("позиция не изменилась",
      db_query(db5, "SELECT party FROM equipment WHERE id=?", (item_id5,))[0][0] == PARTY)
del os.environ["ADMIN_PASSWORD"]

print()
print("ИТОГ:", sum(PASSED), "/", len(PASSED),
      "проверок OK" if all(PASSED) else "— ЕСТЬ ПРОВАЛЫ")
sys.exit(0 if all(PASSED) else 1)
