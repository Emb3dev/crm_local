from sqlmodel import SQLModel, create_engine, Session

engine = create_engine("sqlite:///./crm.db", connect_args={"check_same_thread": False})

def init_db():
    SQLModel.metadata.create_all(engine)
    with engine.begin() as conn:
        cols = {row[1] for row in conn.exec_driver_sql("PRAGMA table_info('client')")}
        if "company_name" not in cols:
            conn.exec_driver_sql("ALTER TABLE client ADD COLUMN company_name VARCHAR")
        if "astreinte" not in cols:
            conn.exec_driver_sql("ALTER TABLE client ADD COLUMN astreinte VARCHAR")
        if "technician_name" not in cols:
            conn.exec_driver_sql("ALTER TABLE client ADD COLUMN technician_name VARCHAR")
        filter_cols = {
            row[1]
            for row in conn.exec_driver_sql("PRAGMA table_info('filterline')")
        }
        if "quantity" not in filter_cols:
            conn.exec_driver_sql(
                "ALTER TABLE filterline ADD COLUMN quantity INTEGER DEFAULT 1"
            )
            conn.exec_driver_sql(
                "UPDATE filterline SET quantity = 1 WHERE quantity IS NULL"
            )
        service_cols = {
            row[1]
            for row in conn.exec_driver_sql("PRAGMA table_info('subcontractedservice')")
        }
        if "status" not in service_cols:
            conn.exec_driver_sql(
                "ALTER TABLE subcontractedservice ADD COLUMN status VARCHAR DEFAULT 'non_commence'"
            )
        if "frequency_interval" not in service_cols:
            conn.exec_driver_sql(
                "ALTER TABLE subcontractedservice ADD COLUMN frequency_interval INTEGER"
            )
        if "frequency_unit" not in service_cols:
            conn.exec_driver_sql(
                "ALTER TABLE subcontractedservice ADD COLUMN frequency_unit VARCHAR"
            )

def get_session():
    with Session(engine) as session:
        yield session
