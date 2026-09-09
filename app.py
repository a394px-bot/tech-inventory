from datetime import datetime, timezone
import io
import os

import pandas as pd
import streamlit as st
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from sqlalchemy import Column, Integer, String, create_engine, inspect, text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker


def utc_now_str():
    """Текущее время в UTC в едином формате записей."""
    return datetime.now(timezone.utc).strftime("%b %d, %Y %H:%M UTC")


# Русские заголовки колонок для отображения таблиц
COL_RU = {
    "id": "ID",
    "date": "Дата",
    "party": "Партия",
    "category": "Категория",
    "model": "Модель",
    "serial_number": "Серийный номер",
    "inv_number": "Инвентарный номер",
    "quantity": "Кол-во",
    "condition": "Состояние",
    "engineer": "Ответственный",
    "date_updated": "Обновлено",
    "equipment_info": "Техника",
    "from_where": "Откуда",
    "to_where": "Куда",
    "receiver": "Принял(а)",
}


def validate_transfer_surnames(own_surname, receiver_surname, engineer_fio, match_fio=True):
    """Проверяет обязательные фамилии при перемещении. Возвращает (ok, текст ошибки).

    match_fio=True — фамилия исполнителя сверяется с ФИО из меню (режим инженера);
    match_fio=False — только обязательность полей (режим администратора).
    """
    own = (own_surname or "").strip()
    recv = (receiver_surname or "").strip()
    fio = (engineer_fio or "").strip()

    if not own:
        return False, "Укажите фамилию исполнителя — это поле обязательно."
    if not recv:
        return False, "Укажите фамилию принимающего — это поле обязательно."

    if not match_fio:
        return True, ""

    fio_words = [w.lower() for w in fio.split()]
    if not fio_words:
        return False, "Сначала укажите ваше ФИО в боковом меню слева."
    if own.lower() not in fio_words:
        return False, (
            f"Ваша фамилия «{own}» не совпадает с ФИО из бокового меню "
            f"(«{fio}»). Проверьте и повторите."
        )
    return True, ""

# --- НАСТРОЙКА БАЗЫ ДАННЫХ ---
# Приоритет: 1) переменная окружения DATABASE_URL; 2) секрет database_url
# (Streamlit Cloud: App → Settings → Secrets) — постоянный Postgres;
# 3) локальный SQLite-файл (запасной режим).
DB_FILE = "inventory_v4.db"
DATABASE_URL = os.environ.get("DATABASE_URL", "") or st.secrets.get(
    "database_url", ""
)
if DATABASE_URL:
    engine = create_engine(DATABASE_URL, echo=False, pool_pre_ping=True)
    IS_SQLITE = False
else:
    engine = create_engine(f"sqlite:///{DB_FILE}", echo=False)
    IS_SQLITE = True
SessionLocal = sessionmaker(bind=engine)
Base = declarative_base()


class Equipment(Base):
    __tablename__ = "equipment"
    id = Column(Integer, primary_key=True, autoincrement=True)
    party = Column(String)
    category = Column(String)
    model = Column(String)
    serial_number = Column(String)
    inv_number = Column(String)
    quantity = Column(Integer)
    condition = Column(String)
    engineer = Column(String)
    date_updated = Column(String)


class LaptopReference(Base):
    __tablename__ = "laptop_reference"
    id = Column(Integer, primary_key=True, autoincrement=True)
    model = Column(String)
    serial_number = Column(String)
    inv_number = Column(String)


class CustomModels(Base):
    __tablename__ = "custom_models"
    id = Column(Integer, primary_key=True, autoincrement=True)
    category = Column(String)
    model = Column(String)


class History(Base):
    __tablename__ = "history"
    id = Column(Integer, primary_key=True, autoincrement=True)
    date = Column(String)
    equipment_info = Column(String)
    from_where = Column(String)
    to_where = Column(String)
    engineer = Column(String)
    receiver = Column(String)


Base.metadata.create_all(bind=engine)


def migrate_history_schema():
    """Добавляет колонку receiver в существующую таблицу истории без потери данных.
    Работает и в SQLite, и в Postgres (через SQLAlchemy Inspector)."""
    if not inspect(engine).has_table("history"):
        return
    cols = [c["name"] for c in inspect(engine).get_columns("history")]
    if "receiver" not in cols:
        with engine.begin() as conn:
            conn.execute(text("ALTER TABLE history ADD COLUMN receiver VARCHAR"))


migrate_history_schema()


# Автоматическая загрузка ноутбуков из Excel
def init_laptop_db():
    session = SessionLocal()
    if session.query(LaptopReference).count() == 0:
        file_path = "Ноутбуки перемещения 29.08.2024 г..xlsx"
        if os.path.exists(file_path):
            try:
                df = pd.read_excel(file_path, sheet_name="Ноутбуки")
                for _, row in df.iterrows():
                    model = str(row.get("Модель", ""))
                    if model == "nan" or not model:
                        continue

                    condition_text = str(
                        row.get("Состояние", "")
                    ).upper() + str(row.get("Обьект заказчика", "")).upper()
                    if (
                        "СПИСАТЬ" in condition_text
                        or "У СИСАДМИНА НА СПИСАНИЕ" in condition_text
                    ):
                        continue

                    serial = str(row.get("Завадской серийный номер", ""))
                    if serial == "nan" or not serial:
                        serial = "-"

                    inv = str(row.get("Инвентарный номер", ""))
                    if inv == "nan" or not inv:
                        inv = "-"

                    ref_item = LaptopReference(
                        model=model, serial_number=serial, inv_number=inv
                    )
                    session.add(ref_item)
                session.commit()
            except Exception as e:
                print(f"Ошибка импорта ноутбуков: {e}")
    session.close()


init_laptop_db()

# --- ИНТЕРФЕЙС STREAMLIT ---
st.set_page_config(
    page_title="Учет оргтехники", page_icon="💻", layout="centered"
)

# --- ЕДИНЫЙ СТИЛЬ ИНТЕРФЕЙСА ---
# Цвета берутся из переменных темы Streamlit (var(--...)),
# поэтому приложение одинаково хорошо выглядит в светлой и тёмной теме.
st.markdown(
    """
    <style>
    /* Базовые шрифты и отступы */
    html, body, [data-testid="stAppViewContainer"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto,
            "Helvetica Neue", Arial, sans-serif;
    }
    [data-testid="stAppViewContainer"] {
        background: var(--background-color);
    }
    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 4rem;
        max-width: 960px;
    }

    /* Акцентная полоса сверху */
    [data-testid="stAppViewContainer"]::before {
        content: "";
        position: fixed;
        top: 0; left: 0; right: 0;
        height: 4px;
        z-index: 99999;
        background: linear-gradient(90deg, #2f6fed, #7c5ce0, #2f6fed);
    }

    /* Заголовки */
    h1, h2, h3 { letter-spacing: -0.01em; }
    h1 { font-weight: 800 !important; }
    h3 { font-weight: 700 !important; }

    /* Шапка приложения (hero) */
    .hero {
        display: flex;
        align-items: center;
        gap: 1rem;
        padding: 1.1rem 1.25rem;
        margin-bottom: 1.2rem;
        background: var(--secondary-background-color);
        border: 1px solid rgba(128, 128, 128, 0.18);
        border-radius: 18px;
    }
    .hero-icon {
        font-size: 2.1rem;
        line-height: 1;
        flex-shrink: 0;
    }
    .hero h1 {
        margin: 0;
        font-size: 1.5rem;
        color: var(--text-color);
    }
    .hero-sub {
        margin-top: 0.15rem;
        font-size: 0.85rem;
        color: var(--text-color);
        opacity: 0.65;
    }

    /* Бейдж партии (цвет задан явно — не зависит от темы Streamlit) */
    .party-chip {
        display: inline-flex;
        align-items: center;
        gap: 0.4rem;
        padding: 0.3rem 0.9rem;
        border-radius: 999px;
        font-weight: 700;
        font-size: 0.95rem;
        color: #2f6fed;
        background: rgba(47, 111, 237, 0.12);
        border: 1px solid rgba(47, 111, 237, 0.32);
        margin-bottom: 0.6rem;
    }

    /* Основные кнопки и активные вкладки — фирменный синий */
    button[kind="primary"], button[data-testid="stBaseButton-primary"] {
        background-color: #2f6fed !important;
    }
    button[kind="primary"]:hover, button[data-testid="stBaseButton-primary"]:hover {
        background-color: #285fd0 !important;
    }
    [data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"] {
        color: #2f6fed !important;
        border-bottom: 2px solid #2f6fed !important;
    }

    /* Кнопки */
    .stButton button, .stDownloadButton button, .stFormSubmitButton button {
        border-radius: 12px;
        font-weight: 600;
        transition: transform 0.12s ease, box-shadow 0.12s ease;
    }
    .stButton button:hover, .stDownloadButton button:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 16px -8px rgba(0, 0, 0, 0.35);
    }
    .stButton button:active, .stDownloadButton button:active {
        transform: translateY(0);
    }

    /* Подсказка только для мобильных (меню спрятано за гамбургером) */
    .device-hint {
        display: none;
    }

    /* Вкладки */
    [data-testid="stTabs"] [data-baseweb="tab-list"] {
        gap: 0.25rem;
    }
    [data-testid="stTabs"] button[data-baseweb="tab"] {
        font-weight: 600;
        border-radius: 10px 10px 0 0;
        padding: 0.55rem 1rem;
    }

    /* Скрываем служебный футер Streamlit */
    [data-testid="stFooter"] { display: none; }

    /* ---------- МОБИЛЬНАЯ ВЕРСИЯ ---------- */
    @media (max-width: 768px) {
        .hero { padding: 0.9rem 1rem; border-radius: 14px; }
        .hero h1 { font-size: 1.25rem; }

        .device-hint {
            display: block;
            padding: 0.75rem 0.9rem;
            margin-bottom: 1rem;
            background: #fff4d6;
            color: #5c4400;
            border: 1px solid #ffd54f;
            border-radius: 12px;
            font-size: 0.9rem;
            line-height: 1.45;
        }

        /* Все поля и кнопки — 16px, чтобы iOS не увеличивал масштаб */
        .stButton button, .stDownloadButton button {
            width: 100%;
            font-size: 16px;
            padding: 0.65rem 0.5rem;
        }
        input, textarea, select, [data-baseweb="select"] > div {
            font-size: 16px !important;
        }
        [data-testid="stTabs"] button[data-baseweb="tab"] {
            font-size: 0.85rem;
            padding: 0.45rem 0.6rem;
        }
        .block-container { padding-top: 0.9rem; }
    }
    </style>
""",
    unsafe_allow_html=True,
)

# --- ШАПКА ПРИЛОЖЕНИЯ ---
st.markdown(
    """
    <div class="hero">
        <div class="hero-icon">💻</div>
        <div>
            <h1>Учёт оргтехники полевых инженеров</h1>
            <div class="hero-sub">Инвентаризация, партии и перемещения техники</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

parties = [f"Партия № {i}" for i in range(1, 32)]

cats_with_identifiers = [
    "Ноутбук",
    "Роутер Huawei",
    "Коммутатор/хаб для обменника",
    "МФУ/Принтер",
    "Усилитель сотовой связи",
    "Сотовый телефон",
    "Стабилизатор напряжения",
    "ИБП",
]

cats_with_qty = [
    "ПК/Системный блок",
    "Монитор",
    "Сетевой фильтр",
    "Lan кабель 50 метров",
]

session = SessionLocal()

# --- БОКОВАЯ ПАНЕЛЬ НАВИГАЦИИ ---
st.sidebar.header("Параметры сеанса")
role = st.sidebar.selectbox("Режим работы:", ["Инженер", "Администратор"])

current_engineer = ""
selected_party = ""

if role == "Инженер":
    # Восстановление партии и ФИО из адресной строки (?party=...&fio=...):
    # инженер, вернувшийся по закладке/истории браузера, попадает сразу в свою
    # партию с уже заполненным ФИО.
    q_party = (st.query_params.get_all("party") or [""])[0]
    q_fio = (st.query_params.get_all("fio") or [""])[0]

    party_index = parties.index(q_party) if q_party in parties else 0
    selected_party = st.sidebar.selectbox(
        "Выберите вашу партию:", parties, index=party_index
    )

    if (
        "last_party" not in st.session_state
        or st.session_state["last_party"] != selected_party
    ):
        st.session_state["last_party"] = selected_party
        # ФИО подставляем только если в URL та же партия (тот же контекст)
        st.session_state["engineer_name"] = (
            q_fio if q_party == selected_party else ""
        )

    current_engineer = st.sidebar.text_input(
        "ФИО ответственного инженера:",
        value=st.session_state["engineer_name"],
        key="eng_input_field",
    )
    st.session_state["engineer_name"] = current_engineer

    # Сохраняем выбор в адресную строку, чтобы он пережил перезагрузку страницы
    if (st.query_params.get_all("party") or [""])[0] != selected_party:
        st.query_params["party"] = selected_party
    url_fio = (st.query_params.get_all("fio") or [""])[0]
    if current_engineer and url_fio != current_engineer:
        st.query_params["fio"] = current_engineer
    elif not current_engineer and "fio" in st.query_params:
        del st.query_params["fio"]

if role == "Инженер":
    st.markdown(
        f'<span class="party-chip">📍 {selected_party}</span>',
        unsafe_allow_html=True,
    )

    if not current_engineer:
        st.markdown(
            '<div class="device-hint">⚠️ <b>Внимание!</b> Нажмите на стрелочку '
            '<b>&gt; (меню)</b> в верхнем левом углу экрана и укажите ваше '
            '<b>ФИО</b> для продолжения работы!</div>',
            unsafe_allow_html=True,
        )
    else:
        tab1, tab2, tab3 = st.tabs(
            ["📋 Список", "➕ Добавить", "🚚 Переместить"]
        )

        with tab1:
            data = pd.read_sql(
                session.query(Equipment)
                .filter(Equipment.party == selected_party)
                .statement,
                engine,
            )
            if not data.empty:
                st.dataframe(
                    data[
                        [
                            "category",
                            "model",
                            "serial_number",
                            "inv_number",
                            "quantity",
                            "condition",
                            "engineer",
                            "date_updated",
                        ]
                    ].rename(columns=COL_RU),
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.info("У данной партии пока нет зарегистрированной техники.")

        with tab2:
            st.markdown("### Добавить оргтехнику")
            all_categories = cats_with_identifiers + cats_with_qty
            cat = st.selectbox("Категория техники", all_categories)

            # Выбор состояния и обязательный комментарий при неудовлетворительном
            cond_option = st.selectbox(
                "Состояние", ["Удовлетворительно", "Не удовлетворительно"]
            )
            if cond_option == "Не удовлетворительно":
                cond_comment = st.text_input(
                    "⚠️ Опишите неисправность (обязательно):"
                )
            else:
                cond_comment = "Удовлетворительно"

            # Логика для техники с серийными/инвентарными номерами
            if cat in cats_with_identifiers:
                model, serial, inv = "", "", ""

                if cat == "Ноутбук":
                    laptops_db = session.query(LaptopReference).all()
                    inv_list = [
                        l.inv_number for l in laptops_db if l.inv_number != "-"
                    ]
                    ser_list = [
                        l.serial_number
                        for l in laptops_db
                        if l.serial_number != "-"
                    ]

                    sel_inv = st.selectbox(
                        "Поиск по инвентарному номеру (из базы)", [""] + inv_list
                    )
                    sel_ser = st.selectbox(
                        "Или поиск по серийному номеру (из базы)", [""] + ser_list
                    )

                    auto_model, auto_serial, auto_inv = "", "", ""
                    if sel_inv:
                        match = (
                            session.query(LaptopReference)
                            .filter(LaptopReference.inv_number == sel_inv)
                            .first()
                        )
                        if match:
                            auto_model, auto_serial, auto_inv = (
                                match.model,
                                match.serial_number,
                                match.inv_number,
                            )
                    elif sel_ser:
                        match = (
                            session.query(LaptopReference)
                            .filter(LaptopReference.serial_number == sel_ser)
                            .first()
                        )
                        if match:
                            auto_model, auto_serial, auto_inv = (
                                match.model,
                                match.serial_number,
                                match.inv_number,
                            )

                    model = st.text_input(
                        "Модель ноутбука",
                        value=auto_model,
                        placeholder="Введите или выберите выше",
                    )
                    serial = st.text_input("Серийный номер", value=auto_serial)
                    inv = st.text_input("Инвентарный номер", value=auto_inv)

                elif cat == "Роутер Huawei":
                    model = "Роутер Huawei"
                    serial = st.text_input("Серийный номер")
                    inv = st.text_input("Инвентарный номер")

                else:
                    custom_db = (
                        session.query(CustomModels)
                        .filter(CustomModels.category == cat)
                        .all()
                    )
                    saved_models = [m.model for m in custom_db]

                    model = st.selectbox(
                        "Модель (из сохраненных или новая)",
                        [""] + saved_models,
                    )
                    if not model:
                        model = st.text_input("Или введите модель вручную:")

                    serial = st.text_input("Серийный номер")
                    inv = st.text_input("Инвентарный номер")

                if st.button("Сохранить позицию", use_container_width=True):
                    if (
                        cond_option == "Не удовлетворительно"
                        and not cond_comment.strip()
                    ):
                        st.error(
                            "❌ Ошибка: Обязательно укажите описание неисправности!"
                        )
                    else:
                        is_model_valid = bool(
                            model
                            and model != "Роутер Huawei"
                            and model != "Не указана"
                        )
                        is_serial_valid = bool(serial and serial != "-")
                        is_inv_valid = bool(inv and inv != "-")

                        if not (
                            is_model_valid or is_serial_valid or is_inv_valid
                        ):
                            st.error(
                                "❌ Ошибка: Заполните хотя бы одно поле (Модель, Серийный или Инвентарный номер)!"
                            )
                        else:
                            # Защита от дублей: один серийный номер — одна позиция
                            duplicate_item = None
                            if serial and serial != "-":
                                duplicate_item = (
                                    session.query(Equipment)
                                    .filter(Equipment.serial_number == serial)
                                    .first()
                                )
                            if duplicate_item:
                                st.error(
                                    f"❌ Позиция с серийным номером «{serial}» уже зарегистрирована в партии «{duplicate_item.party}». Дубликат не сохранён."
                                )
                            else:
                                if (
                                    cat not in ["Ноутбук", "Роутер Huawei"]
                                    and model
                                ):
                                    existing = (
                                        session.query(CustomModels)
                                        .filter(
                                            CustomModels.category == cat,
                                            CustomModels.model == model,
                                        )
                                        .first()
                                    )
                                    if not existing:
                                        session.add(
                                            CustomModels(
                                                category=cat, model=model
                                            )
                                        )
                                        session.commit()

                                new_item = Equipment(
                                    party=selected_party,
                                    category=cat,
                                    model=model
                                    if model
                                    else "Роутер Huawei"
                                    if cat == "Роутер Huawei"
                                    else "Не указана",
                                    serial_number=serial if serial else "-",
                                    inv_number=inv if inv else "-",
                                    quantity=1,
                                    condition=cond_comment,
                                    engineer=current_engineer,
                                    date_updated=utc_now_str(),
                                )
                                session.add(new_item)
                                session.commit()
                                st.success("Успешно добавлено!")
                                st.rerun()

            else:
                qty = st.number_input("Количество (шт.)", min_value=1, value=1)

                if st.button(
                    "Сохранить количество", use_container_width=True
                ):
                    if (
                        cond_option == "Не удовлетворительно"
                        and not cond_comment.strip()
                    ):
                        st.error(
                            "❌ Ошибка: Обязательно укажите описание неисправности!"
                        )
                    else:
                        new_item = Equipment(
                            party=selected_party,
                            category=cat,
                            model=cat,
                            serial_number="-",
                            inv_number="-",
                            quantity=qty,
                            condition=cond_comment,
                            engineer=current_engineer,
                            date_updated=utc_now_str(),
                        )
                        session.add(new_item)
                        session.commit()
                        st.success("Успешно добавлено!")
                        st.rerun()

        with tab3:
            st.markdown("### Перемещение техники")
            data = pd.read_sql(
                session.query(Equipment)
                .filter(Equipment.party == selected_party)
                .statement,
                engine,
            )
            if not data.empty:
                item_options = {
                    f"{row.category} — {row.model} (Сер: {row.serial_number})": row.id
                    for index, row in data.iterrows()
                }
                selected_item_label = st.selectbox(
                    "Выберите позицию", list(item_options.keys())
                )
                destination = st.selectbox(
                    "Куда переместить?", ["База"] + parties
                )

                st.markdown("#### Подтверждение ответственных")
                st.caption(
                    "Оба поля обязательны: перемещение фиксируется на вас и на принимающего."
                )
                confirm_surname = st.text_input(
                    "Ваша фамилия (кто выполняет перемещение)",
                    placeholder="Например: Иванов",
                )
                receiver_surname = st.text_input(
                    "Фамилия принимающего (кому передаёте технику)",
                    placeholder="Например: Петров",
                )

                if st.button(
                    "Подтвердить перемещение", use_container_width=True
                ):
                    ok, err = validate_transfer_surnames(
                        confirm_surname, receiver_surname, current_engineer
                    )
                    if not ok:
                        st.error("❌ " + err)
                    else:
                        item_id = item_options[selected_item_label]
                        item_to_move = (
                            session.query(Equipment)
                            .filter(Equipment.id == item_id)
                            .first()
                        )
                        if item_to_move:
                            if item_to_move.party == destination:
                                st.error(
                                    "❌ Техника уже находится в этой партии. Выберите другое место."
                                )
                            else:
                                old_party = item_to_move.party
                                item_to_move.party = destination
                                item_to_move.date_updated = utc_now_str()

                                history_entry = History(
                                    date=utc_now_str(),
                                    equipment_info=f"{item_to_move.category} {item_to_move.model}",
                                    from_where=old_party,
                                    to_where=destination,
                                    engineer=current_engineer,
                                    receiver=receiver_surname.strip(),
                                )
                                session.add(history_entry)
                                session.commit()
                                st.success("Перемещение выполнено успешно!")
                                st.rerun()
            else:
                st.info("Нет доступной техники для перемещения.")

elif role == "Администратор":
    st.sidebar.subheader("🔒 Авторизация администратора")

    # Пароль хранится в секретах Streamlit (App → Settings → Secrets → admin_password),
    # а не в коде. Если секрет не настроен — вход невозможен.
    ADMIN_PASSWORD = st.secrets.get("admin_password", "")

    if not ADMIN_PASSWORD:
        st.error(
            "Пароль администратора не настроен. Добавьте в секреты приложения "
            "(`App → Settings → Secrets`) ключ `admin_password`, затем обновите страницу."
        )
        st.stop()

    password = st.sidebar.text_input(
        "Введите пароль администратора", type="password"
    )

    if password == ADMIN_PASSWORD:
        st.success("Добро пожаловать в панель администратора!")
        all_data = pd.read_sql(session.query(Equipment).statement, engine)
        history_data = pd.read_sql(session.query(History).statement, engine)

        tab_dash, tab_all, tab_hist, tab_move = st.tabs(
            ["📊 Сводка", "📁 Реестр", "📜 История", "🚚 Переместить"]
        )

        with tab_dash:
            st.markdown("### Сводная таблица по всем партиям")
            if not all_data.empty:
                total_units = int(all_data["quantity"].sum())
                defective = all_data[
                    all_data["condition"].astype(str).str.strip().str.lower()
                    != "удовлетворительно"
                ]
                bad_units = int(defective["quantity"].sum())

                c1, c2, c3 = st.columns(3)
                c1.metric("Всего единиц техники", total_units)
                c2.metric("Требуют внимания", bad_units)
                c3.metric("Партий с техникой", all_data["party"].nunique())

                summary = (
                    all_data.groupby(["party", "category"])["quantity"]
                    .sum()
                    .reset_index()
                )
                pivot_summary = summary.pivot(
                    index="party", columns="category", values="quantity"
                ).fillna(0)
                pivot_summary.loc["ИТОГО по категориям"] = pivot_summary.sum(axis=0)
                pivot_summary["ИТОГО по партии"] = pivot_summary.sum(axis=1)
                st.dataframe(pivot_summary, use_container_width=True)

                # Полный отчёт в Excel: сводка по партиям, итоги по позициям,
                # перечень техники в неудовлетворительном состоянии
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    pivot_summary.to_excel(
                        writer, sheet_name="Сводка по партиям"
                    )
                    pos_total = (
                        all_data.groupby(["category", "model"])["quantity"]
                        .sum()
                        .reset_index()
                        .rename(columns=COL_RU)
                    )
                    pos_total.sort_values(
                        by=["Категория", "Кол-во"],
                        ascending=[True, False],
                        inplace=True,
                    )
                    pos_total.to_excel(
                        writer, sheet_name="Всего по позициям", index=False
                    )
                    def_cols = [
                        "party", "category", "model", "serial_number",
                        "inv_number", "quantity", "condition", "engineer",
                        "date_updated",
                    ]
                    defective[def_cols].rename(columns=COL_RU).to_excel(
                        writer, sheet_name="Неудовлетворительные", index=False
                    )
                    # Оформление: жирная шапка и автоширина колонок на всех листах
                    for ws in writer.book.worksheets:
                        for cell in ws[1]:
                            cell.font = Font(bold=True)
                        for col_cells in ws.columns:
                            max_len = max(
                                (len(str(c.value)) if c.value is not None else 0)
                                for c in col_cells
                            )
                            ws.column_dimensions[
                                get_column_letter(col_cells[0].column)
                            ].width = min(max_len + 2, 42)
                output.seek(0)
                st.download_button(
                    label="📥 Скачать сводный отчёт в Excel (3 листа)",
                    data=output,
                    file_name="inventory_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("В базе пока нет записей.")

        with tab_all:
            st.markdown("### Общая база данных оргтехники")
            if not all_data.empty:
                st.dataframe(
                    all_data.rename(columns=COL_RU),
                    use_container_width=True,
                    hide_index=True,
                )

                # Выгрузка общего списка в Excel
                output_all = io.BytesIO()
                with pd.ExcelWriter(output_all, engine="openpyxl") as writer:
                    all_data.to_excel(writer, index=False, sheet_name="Реестр")
                output_all.seek(0)
                st.download_button(
                    label="📥 Скачать полный реестр в Excel",
                    data=output_all,
                    file_name="all_equipment_registry.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                st.markdown("---")
                st.markdown("### Удаление позиции из реестра")
                del_options = {
                    f"ID {row.id} | Партия: {row.party} | {row.category} — {row.model} (Сер: {row.serial_number})": row.id
                    for _, row in all_data.iterrows()
                }
                selected_del_label = st.selectbox(
                    "Выберите позицию для удаления:", list(del_options.keys())
                )
                if st.button("🗑️ Удалить выбранную позицию"):
                    item_id_to_del = del_options[selected_del_label]
                    session.query(Equipment).filter(
                        Equipment.id == item_id_to_del
                    ).delete()
                    session.commit()
                    st.success("Позиция успешно удалена!")
                    st.rerun()

                if IS_SQLITE and os.path.exists(DB_FILE):
                    st.markdown("---")
                    st.markdown("### 💾 Резервная копия базы")
                    st.caption(
                        "Скачивайте копию регулярно: при перезапуске приложения "
                        "на Streamlit Cloud временные файлы слота могут быть удалены. "
                        "Копию можно восстановить локально или перенести в Postgres "
                        "скриптом migrate_db.py."
                    )
                    with open(DB_FILE, "rb") as db_file:
                        st.download_button(
                            label="Скачать резервную копию базы (.db)",
                            data=db_file.read(),
                            file_name="inventory_backup.db",
                            mime="application/octet-stream",
                        )
            else:
                st.info("Реестр пуст.")

        with tab_hist:
            st.markdown("### Журнал перемещений")
            if not history_data.empty:
                st.dataframe(
                    history_data.rename(columns=COL_RU),
                    use_container_width=True,
                    hide_index=True,
                )

                hist_out = io.BytesIO()
                with pd.ExcelWriter(hist_out, engine="openpyxl") as writer:
                    history_data.rename(columns=COL_RU).to_excel(
                        writer, sheet_name="История перемещений", index=False
                    )
                hist_out.seek(0)
                st.download_button(
                    label="📥 Скачать историю перемещений в Excel",
                    data=hist_out,
                    file_name="transfer_history.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                st.markdown("---")
                st.markdown("### Удаление записей истории")
                hist_del_options = {
                    f"ID {row.id} | Дата: {row.date} | {row.equipment_info} ({row.from_where} ➔ {row.to_where})": row.id
                    for _, row in history_data.iterrows()
                }
                selected_hist_label = st.selectbox(
                    "Выберите запись истории для удаления:",
                    list(hist_del_options.keys()),
                )
                if st.button("🗑️ Удалить выбранную запись истории"):
                    hist_id_to_del = hist_del_options[selected_hist_label]
                    session.query(History).filter(
                        History.id == hist_id_to_del
                    ).delete()
                    session.commit()
                    st.success("Запись истории успешно удалена!")
                    st.rerun()
            else:
                st.info("История перемещений пуста.")

        with tab_move:
            st.markdown("### Перемещение техники")
            if not all_data.empty:
                item_options = {
                    f"ID {row.id} | {row.party} | {row.category} — {row.model} (Сер: {row.serial_number})": row.id
                    for _, row in all_data.iterrows()
                }
                selected_item_label = st.selectbox(
                    "Выберите позицию", list(item_options.keys())
                )
                destination = st.selectbox(
                    "Куда переместить?", ["База"] + parties
                )

                st.markdown("#### Подтверждение ответственных")
                st.caption(
                    "Оба поля обязательны: перемещение фиксируется на вас и на принимающего."
                )
                admin_surname = st.text_input(
                    "Ваша фамилия (кто выполняет перемещение)",
                    placeholder="Например: Иванов",
                )
                receiver_surname = st.text_input(
                    "Фамилия принимающего (кому передаёте технику)",
                    placeholder="Например: Петров",
                )

                if st.button(
                    "Подтвердить перемещение", use_container_width=True
                ):
                    ok, err = validate_transfer_surnames(
                        admin_surname,
                        receiver_surname,
                        current_engineer,
                        match_fio=False,
                    )
                    if not ok:
                        st.error("❌ " + err)
                    else:
                        item_id = item_options[selected_item_label]
                        item_to_move = (
                            session.query(Equipment)
                            .filter(Equipment.id == item_id)
                            .first()
                        )
                        if item_to_move:
                            if item_to_move.party == destination:
                                st.error(
                                    "❌ Позиция уже находится в «"
                                    + destination
                                    + "». Выберите другое место."
                                )
                            else:
                                old_party = item_to_move.party
                                item_to_move.party = destination
                                item_to_move.date_updated = utc_now_str()

                                history_entry = History(
                                    date=utc_now_str(),
                                    equipment_info=(
                                        f"{item_to_move.category} {item_to_move.model}"
                                    ),
                                    from_where=old_party,
                                    to_where=destination,
                                    engineer=admin_surname.strip(),
                                    receiver=receiver_surname.strip(),
                                )
                                session.add(history_entry)
                                session.commit()
                                st.success("Перемещение выполнено успешно!")
                                st.rerun()
            else:
                st.info("В реестре нет техники для перемещения.")

    else:
        if password != "":
            st.error("Неверный пароль администратора!")
        st.warning("Введите пароль администратора в боковой панели слева.")

session.close()