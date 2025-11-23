from sqlalchemy import func
from sqlalchemy.exc import OperationalError
from sqlmodel import SQLModel, create_engine, Session, select

from defaults import DEFAULT_PRESTATION_GROUPS
from models import (
    Client,
    Entreprise,
    PrestationDefinition,
    Supplier,
    SupplierCategory,
    SupplierContact,
)

engine = create_engine("sqlite:///./crm.db", connect_args={"check_same_thread": False})


def _rebuild_filterline_table(conn, filter_cols):
    """Ensure the filterline table matches the expected schema."""

    filter_type_expr = "filter_type"
    if "filter_type" in filter_cols and "format_type" in filter_cols:
        filter_type_expr = "COALESCE(filter_type, format_type)"
    elif "filter_type" not in filter_cols and "format_type" in filter_cols:
        filter_type_expr = "format_type"

    quantity_expr = "COALESCE(quantity, 1)" if "quantity" in filter_cols else "1"
    order_week_expr = "order_week" if "order_week" in filter_cols else "NULL"
    created_at_expr = "created_at" if "created_at" in filter_cols else "CURRENT_TIMESTAMP"

    conn.exec_driver_sql("DROP TABLE IF EXISTS filterline_tmp")
    conn.exec_driver_sql(
        """
        CREATE TABLE filterline_tmp (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            site VARCHAR NOT NULL,
            equipment VARCHAR NOT NULL,
            efficiency VARCHAR,
            filter_type VARCHAR NOT NULL,
            dimensions VARCHAR,
            quantity INTEGER NOT NULL DEFAULT 1,
            order_week VARCHAR,
            created_at DATETIME
        )
        """
    )
    conn.exec_driver_sql(
        f"""
        INSERT INTO filterline_tmp (
            id,
            site,
            equipment,
            efficiency,
            filter_type,
            dimensions,
            quantity,
            order_week,
            created_at
        )
        SELECT
            id,
            site,
            equipment,
            efficiency,
            {filter_type_expr},
            dimensions,
            {quantity_expr},
            {order_week_expr},
            {created_at_expr}
        FROM filterline
        """
    )
    conn.exec_driver_sql("DROP TABLE filterline")
    conn.exec_driver_sql("ALTER TABLE filterline_tmp RENAME TO filterline")


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
        if "entreprise_id" not in cols:
            conn.exec_driver_sql(
                "ALTER TABLE client ADD COLUMN entreprise_id INTEGER REFERENCES entreprise(id)"
            )
        filter_cols = {
            row[1]
            for row in conn.exec_driver_sql("PRAGMA table_info('filterline')")
        }
        if "filter_type" not in filter_cols and "format_type" in filter_cols:
            try:
                conn.exec_driver_sql(
                    "ALTER TABLE filterline RENAME COLUMN format_type TO filter_type"
                )
            except OperationalError:
                _rebuild_filterline_table(conn, filter_cols)
            filter_cols = {
                row[1]
                for row in conn.exec_driver_sql("PRAGMA table_info('filterline')")
            }
        if "filter_type" in filter_cols and "format_type" in filter_cols:
            _rebuild_filterline_table(conn, filter_cols)
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
        if "order_week" not in filter_cols:
            conn.exec_driver_sql(
                "ALTER TABLE filterline ADD COLUMN order_week VARCHAR"
            )
        belt_cols = {
            row[1]
            for row in conn.exec_driver_sql("PRAGMA table_info('beltline')")
        }
        if "order_week" not in belt_cols:
            conn.exec_driver_sql(
                "ALTER TABLE beltline ADD COLUMN order_week VARCHAR"
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
        if "supplier_id" not in service_cols:
            conn.exec_driver_sql(
                "ALTER TABLE subcontractedservice ADD COLUMN supplier_id INTEGER REFERENCES supplier(id)"
            )
        user_cols = {
            row[1]
            for row in conn.exec_driver_sql("PRAGMA table_info('user')")
        }
        if "last_login_at" not in user_cols:
            conn.exec_driver_sql(
                "ALTER TABLE user ADD COLUMN last_login_at DATETIME"
            )
        if "last_logout_at" not in user_cols:
            conn.exec_driver_sql(
                "ALTER TABLE user ADD COLUMN last_logout_at DATETIME"
            )
        if "last_active_at" not in user_cols:
            conn.exec_driver_sql(
                "ALTER TABLE user ADD COLUMN last_active_at DATETIME"
            )
        if "is_online" not in user_cols:
            conn.exec_driver_sql(
                "ALTER TABLE user ADD COLUMN is_online BOOLEAN DEFAULT 0"
            )
        supplier_contact_cols = {
            row[1]
            for row in conn.exec_driver_sql("PRAGMA table_info('suppliercontact')")
        }
        if "description" not in supplier_contact_cols:
            conn.exec_driver_sql(
                "ALTER TABLE suppliercontact ADD COLUMN description VARCHAR"
            )
    with Session(engine) as session:
        clients_without = session.exec(
            select(Client).where(Client.entreprise_id.is_(None))
        ).all()
        for client in clients_without:
            company_name = (client.company_name or client.name or "").strip()
            if not company_name:
                continue
            entreprise = session.exec(
                select(Entreprise).where(Entreprise.nom == company_name)
            ).first()
            if not entreprise:
                statut_value = True
                if client.status:
                    statut_value = client.status == "actif"
                entreprise = Entreprise(
                    nom=company_name,
                    adresse_facturation=client.billing_address,
                    tag=client.tags,
                    statut=statut_value,
                )
                session.add(entreprise)
                session.flush()
            else:
                if client.billing_address and not entreprise.adresse_facturation:
                    entreprise.adresse_facturation = client.billing_address
                if client.tags and not entreprise.tag:
                    entreprise.tag = client.tags
                if client.status:
                    entreprise.statut = client.status == "actif"
            client.entreprise_id = entreprise.id
            client.company_name = entreprise.nom
        session.commit()

        definitions_count = session.exec(
            select(func.count(PrestationDefinition.id))
        ).one()
        if definitions_count == 0:
            for group in DEFAULT_PRESTATION_GROUPS:
                category = group["category"]
                for option in group.get("options", []):
                    definition = PrestationDefinition(
                        key=option["value"],
                        label=option["label"],
                        budget_code=option["budget_code"],
                        category=category,
                        position=option.get("position", 0),
                    )
                    session.add(definition)
            session.commit()
        supplier_category_count = session.exec(
            select(func.count(SupplierCategory.id))
        ).one()
        if supplier_category_count == 0:
            existing_categories = set()
            for supplier in session.exec(select(Supplier)).all():
                raw_categories = (supplier.categories or "").replace(";", ",")
                for part in raw_categories.split(","):
                    label = part.strip()
                    if label:
                        existing_categories.add(label)
            for label in sorted(existing_categories):
                session.add(SupplierCategory(label=label))
            session.commit()
    SQLModel.metadata.create_all(engine)

def get_session():
    with Session(engine) as session:
        yield session
