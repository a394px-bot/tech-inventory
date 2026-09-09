from datetime import datetime
import io
import os
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
    page_title="Учет оргтехники", page_icon="💻", layout="centered"
)

# Добавляем мобильные стили CSS
st.markdown(
    """
    <style>
    @media (max-width: 768px) {
        .stButton button {
            width: 100%;
            font-size: 16px;
            padding: 10px;
        }
        .stSelectbox, .stTextInput, .stNumberInput {
            font-size: 16px !important;
        }
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("💻 Учет оргтехники полевых инженеров")

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
    selected_party = st.sidebar.selectbox("Выберите вашу партию:", parties)

    if (
        "last_party" not in st.session_state
        or st.session_state["last_party"] != selected_party
    ):
        st.session_state["last_party"] = selected_party
        st.session_state["engineer_name"] = ""

    current_engineer = st.sidebar.text_input(
        "ФИО ответственного инженера:",
        value=st.session_state["engineer_name"],
        key="eng_input_field",
    )
    st.session_state["engineer_name"] = current_engineer

if role == "Инженер":
    st.markdown(f"### Партия: **{selected_party}**")

    if not current_engineer:
        st.warning(
            "⚠️ **Внимание!** Нажмите на стрелочку **> (меню)** в верхнем левом углу экрана и укажите ваше **ФИО** для продолжения работы!"
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
                    ],
                    use_container_width=True,
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
                                        CustomModels(category=cat, model=model)
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
                                date_updated="Sep 9, 2026 10:37 UTC",
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
                            date_updated="Sep 9, 2026 10:37 UTC",
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

                if st.button(
                    "Подтвердить перемещение", use_container_width=True
                ):
                    item_id = item_options[selected_item_label]
                    item_to_move = (
                        session.query(Equipment)
                        .filter(Equipment.id == item_id)
                        .first()
                    )
                    if item_to_move:
                        old_party = item_to_move.party
                        item_to_move.party = destination
                        item_to_move.date_updated = "Sep 9, 2026 10:37 UTC"

                        history_entry = History(
                            date="Sep 9, 2026 10:37 UTC",
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
            ["📊 Сводка", "📁 Реестр", "📜 История"]
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

                # Выгрузка дашборда в Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    pivot_summary.to_excel(writer, sheet_name="Сводка")
                output.seek(0)
                st.download_button(
                    label="📥 Скачать сводку в Excel",
                    data=output,
                    file_name="dashboard_summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("В базе пока нет записей.")

        with tab_all:
            st.markdown("### Общая база данных оргтехники")
            if not all_data.empty:
                st.dataframe(all_data, use_container_width=True)

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
            else:
                st.info("Реестр пуст.")

        with tab_hist:
            st.markdown("### Журнал перемещений")
            if not history_data.empty:
                st.dataframe(history_data, use_container_width=True)

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

    else:
        if password != "":
            st.error("Неверный пароль администратора!")
        st.warning("Введите пароль администратора в боковой панели слева.")

session.close()
