from sqlalchemy import inspect, text
from sqlmodel import SQLModel, Session, create_engine

engine = create_engine("sqlite:///./crm.db", connect_args={"check_same_thread": False})


def init_db() -> None:
    SQLModel.metadata.create_all(engine)
    with engine.begin() as connection:
        inspector = inspect(connection)
        if "client" in inspector.get_table_names():
            columns = {column["name"] for column in inspector.get_columns("client")}
            if "depannage" not in columns:
                connection.execute(
                    text("ALTER TABLE client ADD COLUMN depannage TEXT DEFAULT 'non_refacturable'")
                )
            if "astreinte" not in columns:
                connection.execute(
                    text("ALTER TABLE client ADD COLUMN astreinte TEXT DEFAULT 'pas_d_astreinte'")
                )
            if "is_active" not in columns:
                connection.execute(
                    text("ALTER TABLE client ADD COLUMN is_active BOOLEAN DEFAULT 1")
                )


def get_session():
    with Session(engine) as session:
        yield session
