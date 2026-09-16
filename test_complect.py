# -*- coding: utf-8 -*-
"""Проверки комплектности усилителей сотовой связи.

Правило: у категории «Усилитель сотовой связи» отмечается комплектность
(«Комплект» / «Не комплект»); при «Не комплект» комментарий обязателен —
что есть, а чего не хватает.

Запуск:  python test_complect.py   (exit 0 — все проверки зелёные)

Тесты идут через настоящее приложение (streamlit.testing.v1.AppTest) на отдельной
временной базе SQLite, поэтому рабочие данные не затрагиваются.

Почему сценарий «Комплект» вынесен в отдельный процесс: тестовый движок Streamlit
после сохранения (внутри приложения вызывается rerun) перестаёт отдавать скрипту
значения полей формы в том же процессе — значения приходят пустыми, и сохранение
падает на проверке «заполните хотя бы одно поле». К приложению это отношения не
имеет: в браузере состояние виджетов своё на каждую сессию. Поэтому сценарий
гоняем чистым процессом, чтобы проверка была честной.
"""
import hashlib
import hmac
import os
import secrets
import sqlite3
import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from streamlit.testing.v1 import AppTest  # noqa: E402

HERE = Path(__file__).parent
PARTY = "Партия № 1"
PASSWORD = "Тест-1234"
ENGINEER = "Иванов"
CATEGORY = "Усилитель сотовой связи"
CHILD_SCENARIO = "--scenario-komplekt"
PASSED = []


def check(name, cond, extra=""):
    PASSED.append(bool(cond))
    print(("OK  " if cond else "FAIL") + f" {name}" + (f" | {extra}" if extra else ""))


def hash_like_app(password, salt=None):
    """Формат как в hash_party_password() из app.py."""
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), bytes.fromhex(salt), 120_000
    )
    return f"pbkdf2_sha256${salt}${digest.hex()}"


def party_sig(password_hash, party, fio):
    """Подпись ссылки — как party_token() в app.py."""
    return hmac.new(
        password_hash.encode("utf-8"),
        f"{party}|{fio}".encode("utf-8"),
        hashlib.sha256,
    ).hexdigest()[:32]


def new_db_path():
    return Path(tempfile.mkdtemp()) / "test_inventory.db"


def db_query(db_path, sql, params=()):
    con = sqlite3.connect(db_path)
    try:
        return con.execute(sql, params).fetchall()
    finally:
        con.close()


def prepare_db(db_path):
    """Пароль партии + подтверждённая на сегодня фамилия: кабинет открыт сразу."""
    password_hash = hash_like_app(PASSWORD)
    con = sqlite3.connect(db_path)
    cur = con.cursor()
    cur.execute(
        """CREATE TABLE IF NOT EXISTS equipment (
            id INTEGER PRIMARY KEY AUTOINCREMENT, party VARCHAR, category VARCHAR,
            model VARCHAR, serial_number VARCHAR, inv_number VARCHAR,
            quantity INTEGER, condition VARCHAR, engineer VARCHAR,
            date_updated VARCHAR)"""
    )
    cur.execute(
        """CREATE TABLE IF NOT EXISTS party_access (
            id INTEGER PRIMARY KEY AUTOINCREMENT, party VARCHAR UNIQUE,
            password_hash VARCHAR, updated_at VARCHAR)"""
    )
    cur.execute(
        """CREATE TABLE IF NOT EXISTS party_session (
            id INTEGER PRIMARY KEY AUTOINCREMENT, party VARCHAR, fio VARCHAR,
            confirmed_date VARCHAR, confirmed_at VARCHAR)"""
    )
    cur.execute(
        "INSERT OR REPLACE INTO party_access (party, password_hash, updated_at) "
        "VALUES (?,?,?)",
        (PARTY, password_hash, "тест"),
    )
    cur.execute(
        "INSERT INTO party_session (party, fio, confirmed_date, confirmed_at) "
        "VALUES (?,?,?,?)",
        (PARTY, ENGINEER, date.today().isoformat(), "тест"),
    )
    con.commit()
    con.close()
    return {
        "party": PARTY,
        "fio": ENGINEER,
        "sig": party_sig(password_hash, PARTY, ENGINEER),
    }


def app_run(db_path, params=None):
    os.environ["DATABASE_URL"] = "sqlite:///" + db_path.as_posix()
    at = AppTest.from_file(str(HERE / "app.py"), default_timeout=120)
    if params:
        for key, value in params.items():
            at.query_params[key] = value
    at.run()
    return at


def texts(at, kind):
    out = [getattr(e, "value", "") for e in getattr(at, kind)]
    out += [getattr(e, "value", "") for e in getattr(at.sidebar, kind)]
    return out


def has_text(at, needle):
    for kind in (
        "title", "header", "subheader", "markdown", "caption", "info", "warning",
        "error", "success", "text",
    ):
        for value in texts(at, kind):
            if needle.lower() in str(value).lower():
                return True
    return False


def find(kind, at, label_part):
    """Находит виджет по части подписи — и в основной области, и в сайдбаре."""
    for container in (at, at.sidebar):
        for widget in getattr(container, kind):
            if label_part.lower() in str(getattr(widget, "label", "")).lower():
                return widget
    return None


def find_button(at, label_part):
    for container in (at, at.sidebar):
        for button in container.button:
            if label_part.lower() in str(button.label).lower():
                return button
    return None


def dataframe_contains(at, needle):
    """Ищет текст в таблицах на экране (str(DataFrame) обрезает колонки)."""
    for frame in at.dataframe:
        data = frame.value
        if not hasattr(data, "values"):
            continue
        flat = " | ".join(str(cell) for cell in data.values.ravel())
        if needle.lower() in flat.lower():
            return True
    return False


def open_cabinet(db_path, params):
    at = app_run(db_path, params)
    check("кабинет открыт (пароль партии + подтверждённая фамилия)",
          has_text(at, "Работает:"),
          "; ".join(str(e.value)[:60] for e in at.exception))
    return at


def fill_and_save(at, complect, comment="", serial="SN-1"):
    """Заполняет форму добавления усилителя и жмёт «Сохранить позицию»."""
    find("selectbox", at, "Категория техники").select(CATEGORY)
    at.run()
    find("radio", at, "Комплектность").set_value(complect)
    at.run()
    if comment:
        find("text_input", at, "Что есть и чего не хватает").set_value(comment)
    find("text_input", at, "Серийный номер").set_value(serial)
    find_button(at, "Сохранить позицию").click()
    at.run()
    return at


def scenario_komplekt():
    """Отдельный процесс: «Комплект» сохраняется без комментария."""
    db_k = new_db_path()
    params_k = prepare_db(db_k)
    at_k = app_run(db_k, params_k)
    check("кабинет открыт (пароль партии + подтверждённая фамилия)",
          has_text(at_k, "Работает:"),
          "; ".join(str(e.value)[:60] for e in at_k.exception))
    fill_and_save(at_k, "Комплект", serial="SN-TEST-2")
    rows = db_query(
        db_k,
        "SELECT complect, complect_comment FROM equipment "
        "WHERE serial_number='SN-TEST-2'",
    )
    check("сохранён «Комплект» без комментария",
          rows == [("Комплект", None)], str(rows))
    check("ошибок на экране нет",
          not [v.value for v in at_k.error], "; ".join(v.value for v in at_k.error))
    print()
    print("ИТОГ:", sum(PASSED), "/", len(PASSED),
          "проверок OK" if all(PASSED) else "— ЕСТЬ ПРОВАЛЫ")
    sys.exit(0 if all(PASSED) else 1)


if len(sys.argv) > 1 and sys.argv[1] == CHILD_SCENARIO:
    scenario_komplekt()


print("=== Подготовка кабинета партии ===")
db = new_db_path()
params = prepare_db(db)

print("=== 1. Усилитель: «Не комплект» без комментария не сохраняется ===")
at = open_cabinet(db, params)
check("форма «Добавить» доступна",
      find("selectbox", at, "Категория техники") is not None)
find("selectbox", at, "Категория техники").select(CATEGORY)
at.run()
check("для усилителя появилось поле комплектности",
      find("radio", at, "Комплектность") is not None)
find("radio", at, "Комплектность").set_value("Не комплект")
at.run()
check("при «Не комплект» запрашивается комментарий",
      find("text_input", at, "Что есть и чего не хватает") is not None)
find("text_input", at, "Серийный номер").set_value("SN-TEST-1")
find_button(at, "Сохранить позицию").click()
at.run()
check("без описания позиция не сохраняется",
      has_text(at, "напишите в комментарии"), "; ".join(texts(at, "error"))[:80])
saved = db_query(
    db, "SELECT count(*) FROM equipment WHERE serial_number='SN-TEST-1'"
)[0][0]
check("в базе записи нет", saved == 0, f"строк: {saved}")

print("=== 2. Усилитель: «Не комплект» с описанием сохраняется ===")
find("text_input", at, "Что есть и чего не хватает").set_value("нет кабеля 5 м")
find_button(at, "Сохранить позицию").click()
at.run()
rows = db_query(
    db,
    "SELECT complect, complect_comment FROM equipment "
    "WHERE serial_number='SN-TEST-1'",
)
check("сохранено «Не комплект» вместе с комментарием",
      rows == [("Не комплект", "нет кабеля 5 м")], str(rows))

print("=== 3. Усилитель: «Комплект» (отдельный процесс) ===")
child = subprocess.run(
    [sys.executable, str(Path(__file__).resolve()), CHILD_SCENARIO],
    capture_output=True, text=True, encoding="utf-8", errors="replace",
    env={**os.environ, "PYTHONIOENCODING": "utf-8"}, cwd=str(HERE),
)
child_lines = [line for line in (child.stdout or "").splitlines() if line.strip()]
for line in child_lines:
    if line.startswith(("OK", "FAIL")):
        check("(отдельно) " + line[4:], line.startswith("OK"), line[4:])
check("сценарий «Комплект» прошёл в отдельном процессе", child.returncode == 0,
      child_lines[-1] if child_lines else (child.stderr or "")[-160:])

print("=== 4. Другие категории комплектность не спрашивают ===")
at = open_cabinet(db, params)
find("selectbox", at, "Категория техники").select("МФУ/Принтер")
at.run()
check("для «МФУ/Принтер» поля комплектности нет",
      find("radio", at, "Комплектность") is None)

print("=== 5. Комплектность видна в списке партии ===")
check("в таблице видно «не комплект» с описанием",
      dataframe_contains(at, "нет кабеля 5 м"), f"таблиц: {len(at.dataframe)}")
check("в таблице есть колонка «Комплектность»",
      any("Комплектность" in list(getattr(f.value, "columns", []))
          for f in at.dataframe))

print("=== 6. Карточка позиции (QR) показывает комплектность ===")
item_id = db_query(
    db, "SELECT id FROM equipment WHERE serial_number='SN-TEST-1'"
)[0][0]
at_card = app_run(db, {"item": str(item_id)})
check("в карточке видна комплектность и чего не хватает",
      has_text(at_card, "Комплектность") and has_text(at_card, "нет кабеля 5 м"))

print("=== 7. Старая база: колонки добавляются, данные целы ===")
db_old = new_db_path()
con = sqlite3.connect(db_old)
con.execute(
    """CREATE TABLE equipment (
        id INTEGER PRIMARY KEY AUTOINCREMENT, party VARCHAR, category VARCHAR,
        model VARCHAR, serial_number VARCHAR, inv_number VARCHAR,
        quantity INTEGER, condition VARCHAR, engineer VARCHAR,
        date_updated VARCHAR)"""
)
con.execute(
    "INSERT INTO equipment (party, category, model, serial_number, inv_number, "
    "quantity, condition, engineer, date_updated) VALUES "
    "('Партия № 1','Усилитель сотовой связи','УС-1','OLD-1','-',1,"
    "'Удовлетворительно','Иванов','тест')"
)
con.commit()
con.close()
at_old = app_run(db_old)
columns = {row[1] for row in db_query(db_old, "PRAGMA table_info(equipment)")}
check("в старой базе появились колонки комплектности",
      {"complect", "complect_comment"} <= columns, str(sorted(columns)))
check("старая запись не потерялась",
      db_query(db_old, "SELECT count(*) FROM equipment")[0][0] == 1)
check("приложение на старой базе не падает", not at_old.exception,
      "; ".join(str(e.value)[:80] for e in at_old.exception))

print()
print("ИТОГ:", sum(PASSED), "/", len(PASSED),
      "проверок OK" if all(PASSED) else "— ЕСТЬ ПРОВАЛЫ")
sys.exit(0 if all(PASSED) else 1)
