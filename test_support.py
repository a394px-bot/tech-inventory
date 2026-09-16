# -*- coding: utf-8 -*-
"""Общие помощники для проверок приложения «Учёт оргтехники».

Здесь всё, что нужно тестам: временная база, вход в кабинет партии по подписанной
ссылке, поиск виджетов на экране и чтение значений из базы. Рабочие данные не
затрагиваются: каждый тест поднимает приложение на своей временной базе.
"""
import hashlib
import hmac
import os
import secrets
import sqlite3
import tempfile
from datetime import date
from pathlib import Path

from streamlit.testing.v1 import AppTest

HERE = Path(__file__).parent
PARTY = "Партия № 1"
PASSWORD = "Тест-1234"
ENGINEER = "Иванов"

OLD_EQUIPMENT_TABLE = """CREATE TABLE IF NOT EXISTS equipment (
    id INTEGER PRIMARY KEY AUTOINCREMENT, party VARCHAR, category VARCHAR,
    model VARCHAR, serial_number VARCHAR, inv_number VARCHAR,
    quantity INTEGER, condition VARCHAR, engineer VARCHAR,
    date_updated VARCHAR)"""


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


def db_execute(db_path, sql, params=()):
    con = sqlite3.connect(db_path)
    try:
        con.execute(sql, params)
        con.commit()
    finally:
        con.close()


def prepare_db(db_path, party=PARTY, password=PASSWORD, engineer=ENGINEER):
    """Пароль партии + подтверждённая на сегодня фамилия: кабинет открыт сразу.

    Таблица позиций создаётся в «старом» виде (без колонок, добавленных позже) —
    так заодно проверяется миграция.
    """
    password_hash = hash_like_app(password)
    con = sqlite3.connect(db_path)
    cur = con.cursor()
    cur.execute(OLD_EQUIPMENT_TABLE)
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
        (party, password_hash, "тест"),
    )
    cur.execute(
        "INSERT INTO party_session (party, fio, confirmed_date, confirmed_at) "
        "VALUES (?,?,?,?)",
        (party, engineer, date.today().isoformat(), "тест"),
    )
    con.commit()
    con.close()
    return {
        "party": party,
        "fio": engineer,
        "sig": party_sig(password_hash, party, engineer),
    }


def add_item(db_path, party, category, model="Не указана", serial="-", inv="-",
             quantity=1, condition="Удовлетворительно", engineer=ENGINEER):
    """Кладёт позицию в базу напрямую (быстрее, чем проходить форму)."""
    db_execute(
        db_path,
        "INSERT INTO equipment (party, category, model, serial_number, inv_number, "
        "quantity, condition, engineer, date_updated) VALUES (?,?,?,?,?,?,?,?,?)",
        (party, category, model, serial, inv, quantity, condition, engineer, "тест"),
    )


def add_laptop_reference(db_path, model, serial, inv):
    """Добавляет ноутбук в справочник (таблица, из которой идёт подстановка полей)."""
    db_execute(
        db_path,
        """CREATE TABLE IF NOT EXISTS laptop_reference (
            id INTEGER PRIMARY KEY AUTOINCREMENT, model VARCHAR,
            serial_number VARCHAR, inv_number VARCHAR)""",
    )
    db_execute(
        db_path,
        "INSERT INTO laptop_reference (model, serial_number, inv_number) "
        "VALUES (?,?,?)",
        (model, serial, inv),
    )


def add_party_password(db_path, party, password):
    """Задаёт пароль ещё одной партии (для проверок прав)."""
    db_execute(
        db_path,
        """CREATE TABLE IF NOT EXISTS party_access (
            id INTEGER PRIMARY KEY AUTOINCREMENT, party VARCHAR UNIQUE,
            password_hash VARCHAR, updated_at VARCHAR)""",
    )
    db_execute(
        db_path,
        "INSERT OR REPLACE INTO party_access (party, password_hash, updated_at) "
        "VALUES (?,?,?)",
        (party, hash_like_app(password), "тест"),
    )


def app_run(db_path, params=None, script_path=None):
    os.environ["DATABASE_URL"] = "sqlite:///" + db_path.as_posix()
    at = AppTest.from_file(
        str(script_path or (HERE / "app.py")), default_timeout=150
    )
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


def dataframe_with_marker_contains(at, marker_column, needle):
    """Ищет текст только в той таблице, у которой есть колонка-признак.

    Нужно, чтобы различать таблицы на одной странице: у реестра есть колонка
    «Обновлено», у журнала действий — «Действие», и запись о списании обязана
    остаться в журнале, хотя из реестра позиция уходит.
    """
    for frame in at.dataframe:
        data = frame.value
        if not hasattr(data, "columns") or marker_column not in list(data.columns):
            continue
        flat = " | ".join(str(cell) for cell in data.values.ravel())
        if needle.lower() in flat.lower():
            return True
    return False


def metric_value(at, label_part):
    """Значение метрики по части подписи (для сводки администратора)."""
    for metric in at.metric:
        if label_part.lower() in str(metric.label).lower():
            return str(metric.value)
    return None


def category_stats(at, total_label):
    """Читает из сводки карточки «всего / исправны / неисправны» по категории.

    total_label — подпись под первым числом, например «Всего ИБП».
    Возвращает (всего, исправны, неисправны) или None, если блок не найден.
    """
    import re

    html = " ".join(str(getattr(block, "value", "")) for block in at.markdown)
    pattern = (
        r'ls-val">(\d+)</div>\s*<div class="ls-label">'
        + re.escape(total_label)
        + r"</div>"
        r'.*?ls-val">(\d+)</div>\s*<div class="ls-label">Исправны</div>'
        r'.*?ls-val">(\d+)</div>\s*<div class="ls-label">Неисправны</div>'
    )
    found = re.search(pattern, html, re.S)
    return tuple(int(value) for value in found.groups()) if found else None


def login_admin(at, password):
    """Переключает режим на «Администратор» и вводит пароль."""
    at.sidebar.selectbox[0].select("Администратор")
    at.run()
    at.sidebar.text_input[0].set_value(password)
    at.run()
    return at
