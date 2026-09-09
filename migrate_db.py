# -*- coding: utf-8 -*-
"""
Перенос данных из SQLite-файла в постоянную базу (например, Postgres).

Зачем: на Streamlit Cloud временное хранилище слота может очищаться при
перезапуске приложения. Постоянная база (Supabase/Neon и т.п.) решает проблему.

Как пользоваться:
    1. В админке приложения нажмите «Скачать резервную копию базы (.db)».
    2. Зарегистрируйте бесплатный Postgres (например, neon.tech) и получите
       строку подключения вида:
           postgresql://user:password@host/dbname
    3. Запустите перенос (нужны установленные пакеты: sqlalchemy, pandas):
           python migrate_db.py inventory_backup.db "postgresql://user:password@host/dbname"
    4. В Streamlit Cloud добавьте секрет: database_url = "postgresql://user:password@host/dbname"
    5. Обновите приложение — оно будет работать с постоянной базой.
       В requirements.txt репозитория добавьте: psycopg[binary]

Запуск: python migrate_db.py <файл.sqlite> "<строка подключения>"
"""
import sys

from sqlalchemy import Column, Integer, String, create_engine, inspect, text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker


# Схема таблиц — должна совпадать с app.py (держите в синхроне)
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


TABLES = [Equipment, LaptopReference, CustomModels, History]


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    src_path, target_url = sys.argv[1], sys.argv[2]

    source = create_engine(f"sqlite:///{src_path}")
    target = create_engine(target_url)

    with source.connect() as conn:
        # Обязательные проверки существования таблиц и их наполнения
        existing = [
            t.__tablename__
            for t in TABLES
            if inspect(source).has_table(t.__tablename__)
        ]
        if not existing:
            print(f"ОШИБКА: в файле {src_path} нет таблиц приложения.")
            sys.exit(1)

        # Создаём схему в целевой базе
        Base.metadata.create_all(bind=target)

        Session = sessionmaker(bind=target)
        target_session = Session()

        for model in TABLES:
            table_name = model.__tablename__
            rows = conn.execute(model.__table__.select()).mappings().all()
            if not rows:
                print(f"  {table_name}: записей нет, пропуск")
                continue
            # Вставляем с сохранением id (нужно для ссылок и удобства)
            for row in rows:
                target_session.execute(model.__table__.insert().values(**dict(row)))
            target_session.commit()

            # Обновляем счётчик автоинкремента в целевой базе (Postgres/SQLite)
            try:
                if target_url.startswith("postgresql"):
                    target_session.execute(
                        text(
                            "SELECT setval(pg_get_serial_sequence(:t, 'id'), "
                            "COALESCE((SELECT MAX(id) FROM " + table_name + "), 1))"
                        ),
                        {"t": table_name},
                    )
                    target_session.commit()
                elif target_url.startswith("sqlite"):
                    target_session.execute(
                        text(
                            f"UPDATE sqlite_sequence SET seq = "
                            f"(SELECT MAX(id) FROM {table_name}) "
                            f"WHERE name = '{table_name}'"
                        )
                    )
                    target_session.commit()
            except Exception as e:  # счётчик не критичен для данных
                print(f"  (предупреждение) не удалось обновить счётчик: {e}")

            print(f"  {table_name}: перенесено записей — {len(rows)}")

        target_session.close()
        print("\nПеренос завершён. Теперь добавьте в секреты Streamlit Cloud:")
        print(f'  database_url = "{target_url}"')


if __name__ == "__main__":
    main()
