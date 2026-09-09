import os
from datetime import datetime
import pandas as pd
import streamlit as st
from sqlalchemy import Column, Integer, String, create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

# --- НАСТРОЙКА БАЗЫ ДАННЫХ ---
DB_FILE = "inventory_v4.db"
engine = create_engine(f"sqlite:///{DB_FILE}", echo=False)
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


Base.metadata.create_all(bind=engine)


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
    page_title="Учет оргтехники полевых партий", page_icon="💻", layout="wide"
)

st.title("💻 Система учета оргтехники полевых инженеров")

parties = [f"Партия № {i}" for i in range(1, 32)]

cats_with_identifiers = [
    "Ноутбук",
    "Роутер Huawei",
    "Коммутатор",
    "МФУ",
    "Усилитель сотовой связи",
    "Телефон",
    "Стабилизатор напряжения",
    "ИБП",
]

cats_with_qty = [
    "ПК",
    "Монитор",
    "Сетевой фильтр",
    "Lan кабель 50 метров",
]

session = SessionLocal()

# --- БОКОВАЯ ПАНЕЛЬ НАВИГАЦИИ ---
st.sidebar.header("Параметры сеанса")
role = st.sidebar.selectbox("Режим работы:", ["Инженер", "Администратор"])

current_engineer = ""
if role == "Инженер":
    selected_party = st.sidebar.selectbox("Выберите вашу партию:", parties)
    current_engineer = st.sidebar.text_input(
        "ФИО ответственного инженера (вводится 1 раз):", value=""
    )

if role == "Инженер":
    st.subheader(f"Управление техникой для: **{selected_party}**")

    if not current_engineer:
        st.warning(
            "⚠️ Пожалуйста, укажите ваше ФИО в боковой панели слева для продолжения работы!"
        )
    else:
        tab1, tab2, tab3 = st.tabs(
            ["📋 Список техники", "➕ Добавить позицию", "🚚 Перемещение"]
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
                    ],
                    use_container_width=True,
                )
            else:
                st.info("У данной партии пока нет зарегистрированной техники.")

        with tab2:
            st.markdown("### Добавить оргтехнику")
            all_categories = cats_with_identifiers + cats_with_qty
            cat = st.selectbox("Выберите категорию техники", all_categories)

            # Логика для техники с серийными/инвентарными номерами
            if cat in cats_with_identifiers:
                model, serial, inv = "", "", "",

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

                    col1, col2 = st.columns(2)
                    with col1:
                        sel_inv = st.selectbox(
                            "Поиск по инвентарному номеру", [""] + inv_list
                        )
                    with col2:
                        sel_ser = st.selectbox(
                            "Или поиск по серийному номеру", [""] + ser_list
                        )

                    if sel_inv:
                        match = (
                            session.query(LaptopReference)
                            .filter(
                                LaptopReference.inv_number == sel_inv
                            )
                            .first()
                        )
                        if match:
                            model, serial, inv = (
                                match.model,
                                match.serial_number,
                                match.inv_number,
                            )
                    elif sel_ser:
                        match = (
                            session.query(LaptopReference)
                            .filter(
                                LaptopReference.serial_number == sel_ser
                            )
                            .first()
                        )
                        if match:
                            model, serial, inv = (
                                match.model,
                                match.serial_number,
                                match.inv_number,
                            )

                    st.info(
                        f"📌 Данные ноутбука: Модель: **{model or 'Не выбрана'}** | Серийник: **{serial or '-'}** | Инвентарник: **{inv or '-'}**"
                    )
                else:
                    # Подгружаем сохраненные модели для категории + стандартные
                    custom_db = (
                        session.query(CustomModels)
                        .filter(CustomModels.category == cat)
                        .all()
                    )
                    saved_models = [m.model for m in custom_db]

                    model = st.selectbox(
                        "Выберите или введите новую модель",
                        [""] + saved_models,
                    )
                    if not model:
                        model = st.text_input(
                            "Или введите наименование/модель вручную:"
                        )

                    serial = st.text_input("Серийный номер")
                    inv = st.text_input("Инвентарный номер")

                condition = st.selectbox(
                    "Состояние",
                    ["Отлично", "Хорошо", "Удовлетворительно", "Неисправно"],
                )

                if st.button("Сохранить позицию"):
                    if cat != "Ноутбук" and model:
                        # Сохраняем кастомную модель в базу для будущих подсказок
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
                                CustomModels(category=cat, model=model)
                            )
                            session.commit()

                    new_item = Equipment(
                        party=selected_party,
                        category=cat,
                        model=model if model else "Не указана",
                        serial_number=serial if serial else "-",
                        inv_number=inv if inv else "-",
                        quantity=1,
                        condition=condition,
                        engineer=current_engineer,
                        date_updated="Sep 9, 2026 08:34 UTC",
                    )
                    session.add(new_item)
                    session.commit()
                    st.success("Успешно добавлено!")
                    st.rerun()

            else:
                # Логика для техники с количеством (ПК, мониторы, кабели)
                qty = st.number_input("Количество (шт.)", min_value=1, value=1)
                condition = st.selectbox(
                    "Состояние",
                    ["Отлично", "Хорошо", "Удовлетворительно", "Неисправно"],
                )

                if st.button("Сохранить количество"):
                    new_item = Equipment(
                        party=selected_party,
                        category=cat,
                        model=cat,
                        serial_number="-",
                        inv_number="-",
                        quantity=qty,
                        condition=condition,
                        engineer=current_engineer,
                        date_updated="Sep 9, 2026 08:34 UTC",
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
                    f"{row.category} — Модель: {row.model} (Сер: {row.serial_number} / Инв: {row.inv_number})": row.id
                    for index, row in data.iterrows()
                }
                selected_item_label = st.selectbox(
                    "Выберите позицию", list(item_options.keys())
                )
                destination = st.selectbox(
                    "Куда переместить?", ["База"] + parties
                )

                if st.button("Подтвердить перемещение"):
                    item_id = item_options[selected_item_label]
                    item_to_move = (
                        session.query(Equipment)
                        .filter(Equipment.id == item_id)
                        .first()
                    )
                    if item_to_move:
                        old_party = item_to_move.party
                        item_to_move.party = destination
                        item_to_move.date_updated = "Sep 9, 2026 08:34 UTC"

                        history_entry = History(
                            date="Sep 9, 2026 08:34 UTC",
                            equipment_info=f"{item_to_move.category} {item_to_move.model}",
                            from_where=old_party,
                            to_where=destination,
                            engineer=current_engineer,
                        )
                        session.add(history_entry)
                        session.commit()
                        st.success("Перемещение выполнено успешно!")
                        st.rerun()
            else:
                st.info("Нет доступной техники для перемещения.")

elif role == "Администратор":
    st.sidebar.subheader("🔒 Авторизация администратора")
    password = st.sidebar.text_input(
        "Введите пароль администратора", type="password"
    )

    ADMIN_PASSWORD = "62133165215Mig./"

    if password == ADMIN_PASSWORD:
        st.success("Добро пожаловать в панель администратора!")
        all_data = pd.read_sql(session.query(Equipment).statement, engine)
        history_data = pd.read_sql(session.query(History).statement, engine)

        tab_dash, tab_all, tab_hist = st.tabs(
            ["📊 Сводка / Дашборд", "📁 Реестр всей техники", "📜 История перемещений"]
        )

        with tab_dash:
            st.markdown("### Сводная таблица по всем партиям")
            if not all_data.empty:
                summary = (
                    all_data.groupby(["party", "category"])["quantity"]
                    .sum()
                    .reset_index()
                )
                pivot_summary = summary.pivot(
                    index="party", columns="category", values="quantity"
                ).fillna(0)
                st.dataframe(pivot_summary, use_container_width=True)
            else:
                st.warning("В базе пока нет записей.")

        with tab_all:
            st.markdown("### Общая база данных оргтехники")
            if not all_data.empty:
                st.dataframe(all_data, use_container_width=True)

        with tab_hist:
            st.markdown("### Журнал перемещений между партиями и базой")
            if not history_data.empty:
                st.dataframe(history_data, use_container_width=True)
            else:
                st.info("История перемещений пуста.")

    else:
        if password != "":
            st.error("Неверный пароль администратора!")
        st.warning(
            "Введите пароль в боковой панели слева для доступа к административной панели."
        )

session.close()