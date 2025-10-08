from typing import List, Optional

from sqlalchemy import or_
from sqlalchemy.orm import selectinload
from sqlmodel import Session, select

from models import Client, ClientCreate, ClientUpdate, Contact


def list_clients(session: Session, q: Optional[str] = None, limit: int = 50) -> List[Client]:
    stmt = select(Client).options(selectinload(Client.contacts)).order_by(Client.created_at.desc())
    if q:
        like = f"%{q}%"
        stmt = (
            stmt.join(Contact, isouter=True)
            .where(
                or_(
                    Client.company_name.ilike(like),
                    Client.billing_address.ilike(like),
                    Client.tag.ilike(like),
                    Contact.name.ilike(like),
                    Contact.email.ilike(like),
                    Contact.phone.ilike(like),
                )
            )
            .distinct()
        )
    return session.exec(stmt.limit(limit)).all()


def get_client(session: Session, client_id: int) -> Optional[Client]:
    stmt = select(Client).where(Client.id == client_id).options(selectinload(Client.contacts))
    return session.exec(stmt).one_or_none()


def create_client(session: Session, data: ClientCreate) -> Client:
    payload = data.model_dump(exclude_none=True)
    contacts_data = payload.pop("contacts", [])
    client = Client(**payload)
    client.contacts = [Contact(**c.model_dump(exclude_none=True)) for c in contacts_data]
    session.add(client)
    session.commit()
    session.refresh(client)
    session.refresh(client, attribute_names=["contacts"])
    return client


def update_client(session: Session, client_id: int, data: ClientUpdate) -> Optional[Client]:
    client = session.get(Client, client_id)
    if not client:
        return None

    updates = data.model_dump(exclude_unset=True, exclude={"contacts"})
    for key, value in updates.items():
        setattr(client, key, value)

    if data.contacts is not None:
        session.refresh(client, attribute_names=["contacts"])
        client.contacts.clear()
        for contact in data.contacts:
            client.contacts.append(Contact(**contact.model_dump(exclude_none=True)))

    session.add(client)
    session.commit()
    session.refresh(client)
    session.refresh(client, attribute_names=["contacts"])
    return client


def delete_client(session: Session, client_id: int) -> bool:
    client = session.get(Client, client_id)
    if not client:
        return False
    session.delete(client)
    session.commit()
    return True
