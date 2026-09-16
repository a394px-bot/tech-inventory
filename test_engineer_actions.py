# -*- coding: utf-8 -*-
"""Проверки кнопок инженера: выход из кабинета, удаление позиции, очистка формы.

Правила, принятые заказчиком:
- у инженера есть выход из кабинета;
- вкладки «Переместить» у инженера нет — перемещения делают не они;
- ошибочную позицию инженер удаляет, и факт удаления остаётся в истории (кто/когда);
- после сохранения позиции поля формы очищаются, и показывается подтверждение —
  чтобы нельзя было случайно сохранить одно и то же дважды.

Запуск:  python test_engineer_actions.py   (exit 0 — все проверки зелёные)

Сценарий с сохранением позиции выполняется отдельным процессом: тестовый движок
Streamlit после сохранения (внутри приложения вызывается rerun) перестаёт отдавать
скрипту значения полей формы. К браузеру это отношения не имеет — там состояние
виджетов своё на каждую сессию.
"""
import os
import subprocess
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from test_support import (  # noqa: E402
    ENGINEER,
    PARTY,
    add_item,
    app_run,
    dataframe_with_marker_contains,
    db_query,
    find,
    find_button,
    has_text,
    login_admin,
    metric_value,
    new_db_path,
    prepare_db,
)

HERE = Path(__file__).parent
CHILD_SCENARIO = "--scenario-add"
PASSED = []


def check(name, cond, extra=""):
    PASSED.append(bool(cond))
    print(("OK  " if cond else "FAIL") + f" {name}" + (f" | {extra}" if extra else ""))


def scenario_add():
    """Отдельный процесс: сохраняем позицию и смотрим, что стало с формой."""
    db = new_db_path()
    params = prepare_db(db)
    at = app_run(db, params)
    find("selectbox", at, "Категория техники").select("МФУ/Принтер")
    at.run()
    find("text_input", at, "Серийный номер").set_value("SN-NEW-1")
    find_button(at, "Сохранить позицию").click()
    at.run()

    saved = db_query(db, "SELECT COUNT(*) FROM equipment WHERE serial_number='SN-NEW-1'")[0][0]
    check("позиция сохранена", saved == 1, f"строк: {saved}")
    check("показано подтверждение, что позиция добавлена",
          has_text(at, "Позиция добавлена") and has_text(at, "SN-NEW-1"),
          "; ".join(str(v.value)[:80] for v in at.success))
    # Очистку полей проверяем по поведению: тестовый движок держит прежнее значение
    # поля в своей памяти, а приложение при повторном сохранении видит пустые поля —
    # именно это и защищает от случайного дубля.
    find_button(at, "Сохранить позицию").click()
    at.run()
    still_one = db_query(
        db, "SELECT COUNT(*) FROM equipment WHERE serial_number='SN-NEW-1'"
    )[0][0]
    check("повторное нажатие не создаёт вторую запись", still_one == 1,
          f"строк: {still_one}")
    check("после сохранения поля пустые — повторное сохранение просит их заполнить",
          has_text(at, "хотя бы одно поле"),
          "; ".join(str(v.value)[:80] for v in at.error))
    print()
    print("ИТОГ:", sum(PASSED), "/", len(PASSED),
          "проверок OK" if all(PASSED) else "— ЕСТЬ ПРОВАЛЫ")
    sys.exit(0 if all(PASSED) else 1)


if len(sys.argv) > 1 and sys.argv[1] == CHILD_SCENARIO:
    scenario_add()


print("=== 1. Кнопки инженера: удаление, а не архив ===")
db = new_db_path()
params = prepare_db(db)
add_item(db, PARTY, "МФУ/Принтер", model="HP-ошибочный", serial="SN-DEL")
at = app_run(db, params)
check("кабинет открыт", has_text(at, "Работает:"),
      "; ".join(str(e.value)[:80] for e in at.exception))
check("есть кнопка «Удалить позицию»", find_button(at, "Удалить позицию") is not None,
      "; ".join(b.label for b in at.button))
check("кнопки «Списать в архив» нет",
      find_button(at, "Списать в архив") is None)

print("=== 2. Вкладки инженера: перемещение на месте, списания нет ===")
check("вкладка «Переместить» вернулась",
      any("Переместить" in str(tab.label) for tab in at.tabs),
      "; ".join(str(getattr(t, "label", "")) for t in at.tabs))
check("есть вкладки «Список» и «Добавить»",
      any("Список" in str(t.label) for t in at.tabs)
      and any("Добавить" in str(t.label) for t in at.tabs))
check("поля перемещения на месте (выбор позиции и получателя)",
      find("selectbox", at, "Куда переместить") is not None
      and find("text_input", at, "Фамилия принимающего") is not None)

print("=== 3. Удаление без подтверждения не срабатывает ===")
find_button(at, "Удалить позицию").click()
at.run()
left = db_query(db, "SELECT COUNT(*) FROM equipment WHERE serial_number='SN-DEL'")[0][0]
check("без галочки позиция на месте", left == 1, f"строк: {left}")
check("приложение просит подтверждение", has_text(at, "Подтверждаю удаление"))

print("=== 4. С подтверждением — удаляет и пишет в историю ===")
checkbox = find("checkbox", at, "Подтверждаю удаление")
if checkbox is not None:
    checkbox.check()
find_button(at, "Удалить позицию").click()
at.run()
left = db_query(db, "SELECT COUNT(*) FROM equipment WHERE serial_number='SN-DEL'")[0][0]
check("позиция удалена", left == 0, f"строк: {left}")
journal = db_query(
    db, "SELECT action, engineer FROM audit_log ORDER BY id DESC LIMIT 1"
)
check("в истории запись об удалении с фамилией",
      journal == [("Удаление", ENGINEER)], str(journal))
check("удалённой позиции нет в списке", not has_text(at, "HP-ошибочный"))

print("=== 5. Выход из кабинета инженера ===")
logout = find_button(at, "Выйти из кабинета")
check("кнопка выхода на месте", logout is not None)
if logout is not None:
    logout.click()
    at.run()
check("после выхода снова экран входа",
      has_text(at, "Чтобы начать работу") and not has_text(at, "Работает:"))
check("ссылка с партией очищена", not dict(at.query_params),
      str(dict(at.query_params)))

print("=== 6. Уведомление и очистка формы после сохранения (отдельный процесс) ===")
child = subprocess.run(
    [sys.executable, str(Path(__file__).resolve()), CHILD_SCENARIO],
    capture_output=True, text=True, encoding="utf-8", errors="replace",
    env={**os.environ, "PYTHONIOENCODING": "utf-8"}, cwd=str(HERE),
)
child_lines = [line for line in (child.stdout or "").splitlines() if line.strip()]
for line in child_lines:
    if line.startswith(("OK", "FAIL")):
        check("(отдельно) " + line[4:], line.startswith("OK"), line[4:])
check("сценарий сохранения прошёл", child.returncode == 0,
      child_lines[-1] if child_lines else (child.stderr or "")[-200:])

print("=== 7. Администратор: списание в архив, выход и реестр ===")
os.environ["ADMIN_PASSWORD"] = "тест-админ"
db2 = new_db_path()
prepare_db(db2)
add_item(db2, PARTY, "Монитор", model="Dell-1", serial="MON-1")
add_item(db2, PARTY, "Монитор", model="Dell-2", serial="MON-2")
at_a = login_admin(app_run(db2), "тест-админ")
check("админ вошёл, реестр с данными не падает",
      metric_value(at_a, "Всего единиц техники") == "2" and not at_a.exception,
      "; ".join(str(e.value)[:100] for e in at_a.exception))
archive_button = find_button(at_a, "Списать в архив")
check("у администратора есть списание в архив", archive_button is not None)
if archive_button is not None:
    archive_button.click()
    at_a.run()
row = db_query(
    db2, "SELECT archived, archived_by, archived_at FROM equipment "
         "WHERE serial_number='MON-1'"
)[0]
check("позиция помечена списанной администратором",
      row[0] == 1 and row[1] == "Администратор" and bool(row[2]), str(row))
check("запись в базе осталась (не удалена)",
      db_query(db2, "SELECT COUNT(*) FROM equipment WHERE serial_number='MON-1'")[0][0] == 1)
check("списанная позиция ушла из реестра",
      not dataframe_with_marker_contains(at_a, "Обновлено", "MON-1"))
check("в истории есть запись о списании",
      dataframe_with_marker_contains(at_a, "Действие", "Списание"))
check("списанная техника не считается в сводке",
      metric_value(at_a, "Всего единиц техники") == "1",
      str(metric_value(at_a, "Всего единиц техники")))
logout_admin = find_button(at_a, "Выйти")
check("у администратора есть кнопка выхода", logout_admin is not None)
if logout_admin is not None:
    logout_admin.click()
    at_a.run()
check("после выхода админка снова просит пароль",
      has_text(at_a, "Введите пароль администратора")
      and not has_text(at_a, "Всего единиц техники"))

print("=== 8. История администратора видит правки инженера ===")
db3 = new_db_path()
params3 = prepare_db(db3)
add_item(db3, PARTY, "Монитор", model="Dell-2", serial="MON-2")
at3 = app_run(db3, params3)
find("text_input", at3, "Модель").set_value("Dell-2 исправленный")
find_button(at3, "Сохранить изменения").click()
at3.run()
at_admin = login_admin(app_run(db3), "тест-админ")
check("правка инженера видна в истории администратора",
      dataframe_with_marker_contains(at_admin, "Действие", "Dell-2 исправленный")
      or dataframe_with_marker_contains(at_admin, "Действие", "Изменение"),
      "журнал действий")
del os.environ["ADMIN_PASSWORD"]

print()
print("ИТОГ:", sum(PASSED), "/", len(PASSED),
      "проверок OK" if all(PASSED) else "— ЕСТЬ ПРОВАЛЫ")
sys.exit(0 if all(PASSED) else 1)
