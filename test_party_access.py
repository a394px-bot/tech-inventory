# -*- coding: utf-8 -*-
"""Проверки доступа по партиям: пароль партии, подтверждение фамилии, подпись ссылки.

Запуск:  python test_party_access.py
Результат: exit 0 — все проверки зелёные, exit 1 — есть провалы.

Тесты поднимают настоящее приложение через streamlit.testing.v1.AppTest на
ОТДЕЛЬНОЙ временной базе SQLite (переменная DATABASE_URL), поэтому рабочие данные
в inventory_v4.db и в облачной базе не затрагиваются.
"""
import hashlib
import os
import secrets
import shutil
import sqlite3
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
PASSED = []


def check(name, cond, extra=""):
    PASSED.append(bool(cond))
    print(("OK  " if cond else "FAIL") + f" {name}" + (f" | {extra}" if extra else ""))


def hash_like_app(password, salt=None):
    """Тот же формат, что hash_party_password() в app.py: pbkdf2_sha256$соль$хэш."""
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), bytes.fromhex(salt), 120_000
    )
    return f"pbkdf2_sha256${salt}${digest.hex()}"


def new_db_path():
    return Path(tempfile.mkdtemp()) / "test_inventory.db"


def prepare_db(db_path, party=None, password=None, engineer=None):
    """Готовит базу: пароль партии и одну позицию с ответственным инженером."""
    con = sqlite3.connect(db_path)
    cur = con.cursor()
    cur.execute(
        """CREATE TABLE IF NOT EXISTS equipment (
            id INTEGER PRIMARY KEY AUTOINCREMENT, party VARCHAR, category VARCHAR,
            model VARCHAR, serial_number VARCHAR, inv_number VARCHAR,
            quantity INTEGER, condition VARCHAR, engineer VARCHAR,
            date_updated VARCHAR)"""
    )
    if party and password:
        cur.execute(
            """CREATE TABLE IF NOT EXISTS party_access (
                id INTEGER PRIMARY KEY AUTOINCREMENT, party VARCHAR UNIQUE,
                password_hash VARCHAR, updated_at VARCHAR)"""
        )
        cur.execute(
            "INSERT OR REPLACE INTO party_access (party, password_hash, updated_at) "
            "VALUES (?,?,?)",
            (party, hash_like_app(password), "тест"),
        )
    if party and engineer:
        cur.execute(
            "INSERT INTO equipment (party, category, model, quantity, condition, "
            "engineer, date_updated) VALUES (?,?,?,?,?,?,?)",
            (party, "Ноутбук", "Lenovo T480", 1, "Удовлетворительно", engineer, "тест"),
        )
    con.commit()
    con.close()


def db_query(db_path, sql, params=()):
    con = sqlite3.connect(db_path)
    try:
        return con.execute(sql, params).fetchall()
    finally:
        con.close()


def app_run(db_path, params=None, script_path=None):
    """Запускает приложение на указанной базе; params — то, что стоит в ссылке."""
    os.environ["DATABASE_URL"] = "sqlite:///" + db_path.as_posix()
    at = AppTest.from_file(
        str(script_path or (HERE / "app.py")), default_timeout=120
    )
    if params:
        for key, value in params.items():
            at.query_params[key] = value
    at.run()
    return at


def texts(at, kind):
    """Все тексты элементов этого вида — и в основной области, и в сайдбаре."""
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


def find_button(at, label):
    for collection in (at.button, at.sidebar.button):
        for button in collection:
            if label.lower() in str(button.label).lower():
                return button
    return None


def local_admin_password():
    """Пароль администратора из локального .streamlit/secrets.toml (в репозиторий
    он не попадает). Файла нет — проверки админки пропускаем, а не падаем."""
    path = HERE / ".streamlit" / "secrets.toml"
    if not path.exists():
        return ""
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line.startswith("admin_password"):
            return line.split("=", 1)[1].strip().strip('"').strip("'")
    return ""


def login(at, party, password):
    """Выбирает партию, вводит пароль и нажимает «Войти»."""
    at.sidebar.selectbox[1].select(party)
    at.sidebar.text_input[0].set_value(password)
    button = find_button(at, "Войти")
    button.click()
    at.run()
    return at


print("=== 1. Чистая база: пароль партии не задан ===")
db = new_db_path()
at = app_run(db)
check("без входа показан экран входа, а не кабинет",
      has_text(at, "Чтобы начать работу") and not has_text(at, "Работает:"))
at = login(at, PARTY, "какой-то-пароль")
check("пароль партии не задан — вход отклонён с подсказкой",
      has_text(at, "пароль ещё не задан"))
check("кабинет закрыт", not has_text(at, "Работает:"))

print("=== 2. Верный пароль партии ===")
db = new_db_path()
prepare_db(db, party=PARTY, password=PASSWORD, engineer=ENGINEER)
at = app_run(db)
at = login(at, PARTY, PASSWORD)
check("после верного пароля открылось окно подтверждения фамилии",
      has_text(at, "Кто сегодня работает"))
check("до подтверждения фамилии кабинет ещё закрыт", not has_text(at, "Работает:"))

print("=== 3. Неверный пароль ===")
db_bad = new_db_path()
prepare_db(db_bad, party="Партия № 2", password="Верный-пароль")
at_bad = app_run(db_bad)
at_bad = login(at_bad, "Партия № 2", "неверный-пароль")
check("неверный пароль — отказ, кабинет закрыт",
      has_text(at_bad, "Неверный пароль партии") and not has_text(at_bad, "Кто сегодня работает"))

print("=== 4. Подтверждение фамилии ===")
# Тестовый движок не возвращает приложению параметры ссылки между прогонами
# (в браузере они сохраняются в адресной строке), поэтому переносим их сами —
# тем самым проверяем и то, что ссылка подписана и переживает перезагрузку.
params_login = {k: v for k, v in dict(at.query_params).items()}
print("     ссылка после входа:", sorted(params_login.keys()))
at = app_run(db, params_login)
confirm = find_button(at, "Подтвердить и продолжить")
check("кнопка подтверждения на месте", confirm is not None)
if confirm is not None:
    confirm.click()
    at.run()
check("на экране подтверждения нет исключений", not at.exception,
      "; ".join(str(e.value)[:80] for e in at.exception))
params_after = {k: v for k, v in dict(at.query_params).items()}
check("после подтверждения фамилия попала в подписанную ссылку",
      "fio" in params_after, str(params_after.get("fio")))
rows = db_query(
    db, "SELECT party, fio, confirmed_date FROM party_session WHERE party=?", (PARTY,)
)
check("подтверждение записано в базу на сегодня",
      rows == [(PARTY, ENGINEER, date.today().isoformat())], str(rows))

print("=== 5. Повторный вход по той же ссылке (без повторных вопросов) ===")
at2 = app_run(db, params_after)
check("кабинет открыт и подписан фамилией",
      has_text(at2, "Работает:") and ENGINEER in " ".join(texts(at2, "caption")))
check("вопрос о фамилии повторно не задаётся",
      not has_text(at2, "Кто сегодня работает"))
check("в кабинете виден список позиций этой партии",
      any("Lenovo T480" in str(getattr(df, "value", "")) for df in at2.dataframe),
      f"таблиц на экране: {len(at2.dataframe)}")

print("=== 6. Подделанная ссылка (вписали чужую партию) ===")
db3 = new_db_path()
prepare_db(db3, party="Партия № 3", password="Секрет-3")
at3 = app_run(db3, {"party": "Партия № 5", "fio": "Петров", "sig": "0" * 32})
check("подпись не совпала — в чужую партию не пускает",
      not has_text(at3, "Работает:") and not has_text(at3, "Кто сегодня работает")
      and has_text(at3, "Чтобы начать работу"))

print("=== 7. Пароль в базе хранится хэшем ===")
stored = db_query(db, "SELECT password_hash FROM party_access WHERE party=?", (PARTY,))
stored_hash = stored[0][0] if stored else ""
check("в базе не открытый пароль, а хэш",
      stored_hash.startswith("pbkdf2_sha256$") and PASSWORD not in stored_hash,
      stored_hash[:24])

print("=== 8. Приложение не падает ===")
check("на экране входа нет исключений", not at3.exception)
check("в кабинете нет исключений", not at.exception and not at_bad.exception)

print("=== 9. Вкладка администратора «Пароли партий» ===")
admin_pwd = local_admin_password()
if not admin_pwd:
    print("SKIP: локальный секрет admin_password не найден — проверка пропущена")
else:
    db4 = new_db_path()
    at4 = app_run(db4)
    at4.sidebar.selectbox[0].select("Администратор")
    at4.run()
    at4.sidebar.text_input[0].set_value(admin_pwd)
    at4.run()
    check("админ вошёл: вкладка «Пароли партий» видна",
          has_text(at4, "Пароли партий") and not at4.exception,
          "; ".join(str(e.value)[:80] for e in at4.exception))
    gen_all = find_button(at4, "Сгенерировать пароли для всех партий")
    check("кнопка массовой выдачи паролей на месте", gen_all is not None)
    if gen_all is not None:
        gen_all.click()
        at4.run()
        issued = db_query(db4, "SELECT count(*) FROM party_access")[0][0]
        check("пароли выданы всем 31 партии", issued == 31, f"записей: {issued}")
        check("показано предупреждение «пароли показываются один раз»",
              has_text(at4, "показываются один раз"))

print("=== 10. Контейнер БЕЗ файла секретов (как на хостинге Timeweb) ===")
# В контейнере secrets.toml нет, а st.secrets без файла не отдаёт значение по
# умолчанию, а выбрасывает StreamlitSecretNotFoundError — из-за этого в проде
# падала админка (Traceback на строке ADMIN_PASSWORD = st.secrets.get(...)).
# Здесь запускаем приложение из папки без .streamlit: пароль идёт из переменной
# окружения, страница не падает.
sandbox = Path(tempfile.mkdtemp())
shutil.copy(HERE / "app.py", sandbox / "app.py")
os.environ["ADMIN_PASSWORD"] = "пароль-из-переменной"
at5 = app_run(new_db_path(), script_path=sandbox / "app.py")
check("приложение запускается без файла секретов", not at5.exception,
      "; ".join(str(e.value)[:120] for e in at5.exception))
at5.sidebar.selectbox[0].select("Администратор")
at5.run()
check("админка открывается без secrets.toml (раньше падала с Traceback)",
      not at5.exception, "; ".join(str(e.value)[:120] for e in at5.exception))
at5.sidebar.text_input[0].set_value("пароль-из-переменной")
at5.run()
check("пароль из переменной окружения ADMIN_PASSWORD принимается",
      has_text(at5, "Пароли партий") and not at5.exception,
      "; ".join(str(e.value)[:120] for e in at5.exception))
del os.environ["ADMIN_PASSWORD"]

print()
print("ИТОГ:", sum(PASSED), "/", len(PASSED),
      "проверок OK" if all(PASSED) else "— ЕСТЬ ПРОВАЛЫ")
sys.exit(0 if all(PASSED) else 1)
