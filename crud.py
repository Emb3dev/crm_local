from typing import List, Optional
from sqlmodel import select, Session
from models import Client, ClientCreate, ClientUpdate

def list_clients(session: Session, q: Optional[str] = None, limit: int = 50) -> List[Client]:
    stmt = select(Client).order_by(Client.created_at.desc())
    if q:
        like = f"%{q}%"
        stmt = stmt.where(
            (Client.name.ilike(like)) | (Client.email.ilike(like)) | (Client.tags.ilike(like))
        )
    return session.exec(stmt.limit(limit)).all()

def get_client(session: Session, client_id: int) -> Optional[Client]:
    return session.get(Client, client_id)

def create_client(session: Session, data: ClientCreate) -> Client:
    c = Client.model_validate(data)
    session.add(c)
    session.commit()
    session.refresh(c)
    return c

def update_client(session: Session, client_id: int, data: ClientUpdate) -> Optional[Client]:
    c = session.get(Client, client_id)
    if not c: return None
    updates = data.model_dump(exclude_unset=True)
    for k, v in updates.items():
        setattr(c, k, v)
    session.add(c)
    session.commit()
    session.refresh(c)
    return c

def delete_client(session: Session, client_id: int) -> bool:
    c = session.get(Client, client_id)
    if not c: return False
    session.delete(c)
    session.commit()
    return True
