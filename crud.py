from typing import Dict, List, Optional

import re

from sqlalchemy import delete
from sqlalchemy.orm import selectinload
from sqlmodel import Session, select

from models import (
    BeltLine,
    BeltLineCreate,
    BeltLineUpdate,
    Client,
    ClientCreate,
    ClientUpdate,
    Contact,
    ContactCreate,
    FilterLine,
    FilterLineCreate,
    FilterLineUpdate,
    SubcontractedService,
    SubcontractedServiceCreate,
    SubcontractedServiceUpdate,
)

def list_clients(
    session: Session,
    q: Optional[str] = None,
    *,
    filters: Optional[Dict[str, str]] = None,
    limit: int = 50,
) -> List[Client]:
    stmt = (
        select(Client)
        .options(
            selectinload(Client.contacts),
            selectinload(Client.subcontractings),
        )
        .order_by(Client.created_at.desc())
    )
    effective_filters = filters or {}
    if q:
        like = f"%{q}%"
        stmt = stmt.outerjoin(Contact)
        stmt = stmt.where(
            (
                Client.name.ilike(like)
                | Client.company_name.ilike(like)
                | Client.email.ilike(like)
                | Client.phone.ilike(like)
                | Client.billing_address.ilike(like)
                | Client.depannage.ilike(like)
                | Client.astreinte.ilike(like)
                | Client.tags.ilike(like)
                | Contact.name.ilike(like)
                | Contact.email.ilike(like)
                | Contact.phone.ilike(like)
            )
        )
        stmt = stmt.distinct()

    status = effective_filters.get("status")
    if status:
        stmt = stmt.where(Client.status == status)

    depannage = effective_filters.get("depannage")
    if depannage:
        stmt = stmt.where(Client.depannage == depannage)

    astreinte = effective_filters.get("astreinte")
    if astreinte:
        stmt = stmt.where(Client.astreinte == astreinte)

    return session.exec(stmt.limit(limit)).all()


def list_client_choices(session: Session) -> List[Client]:
    stmt = select(Client).order_by(Client.company_name.asc(), Client.name.asc())
    return session.exec(stmt).all()

def get_client(session: Session, client_id: int) -> Optional[Client]:
    return session.get(Client, client_id)

def create_client(
    session: Session,
    data: ClientCreate,
    contacts: Optional[List[ContactCreate]] = None,
) -> Client:
    c = Client.model_validate(data)
    session.add(c)
    session.flush()

    for contact_data in contacts or []:
        contact = Contact(client_id=c.id, **contact_data.model_dump())
        session.add(contact)

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
    session.exec(delete(Contact).where(Contact.client_id == client_id))
    session.delete(c)
    session.commit()
    return True


def create_contact(session: Session, client_id: int, data: ContactCreate) -> Optional[Contact]:
    client = session.get(Client, client_id)
    if not client:
        return None
    contact = Contact(client_id=client_id, **data.model_dump())
    session.add(contact)
    session.commit()
    session.refresh(contact)
    return contact


def delete_contact(session: Session, client_id: int, contact_id: int) -> bool:
    contact = session.get(Contact, contact_id)
    if not contact or contact.client_id != client_id:
        return False
    session.delete(contact)
    session.commit()
    return True


def create_subcontracted_service(
    session: Session, client_id: int, data: SubcontractedServiceCreate
) -> Optional[SubcontractedService]:
    client = session.get(Client, client_id)
    if not client:
        return None
    record = SubcontractedService(client_id=client_id, **data.model_dump())
    session.add(record)
    session.commit()
    session.refresh(record)
    return record


def delete_subcontracted_service(
    session: Session, client_id: int, service_id: int
) -> bool:
    record = session.get(SubcontractedService, service_id)
    if not record or record.client_id != client_id:
        return False
    session.delete(record)
    session.commit()
    return True


def list_subcontracted_services(
    session: Session,
    q: Optional[str] = None,
    *,
    filters: Optional[Dict[str, str]] = None,
    limit: int = 200,
) -> List[SubcontractedService]:
    stmt = (
        select(SubcontractedService)
        .options(selectinload(SubcontractedService.client))
        .order_by(SubcontractedService.created_at.desc())
    )
    effective_filters = filters or {}
    if q:
        like = f"%{q}%"
        stmt = stmt.outerjoin(Client)
        stmt = stmt.where(
            (
                SubcontractedService.prestation_label.ilike(like)
                | SubcontractedService.category.ilike(like)
                | SubcontractedService.budget_code.ilike(like)
                | Client.company_name.ilike(like)
                | Client.name.ilike(like)
            )
        )
        stmt = stmt.distinct()

    category = effective_filters.get("category")
    if category:
        stmt = stmt.where(SubcontractedService.category == category)

    frequency = effective_filters.get("frequency")
    if frequency:
        stmt = stmt.where(SubcontractedService.frequency == frequency)

    return session.exec(stmt.limit(limit)).all()


def get_subcontracted_service(
    session: Session, service_id: int
) -> Optional[SubcontractedService]:
    stmt = (
        select(SubcontractedService)
        .where(SubcontractedService.id == service_id)
        .options(selectinload(SubcontractedService.client))
    )
    return session.exec(stmt).one_or_none()


def update_subcontracted_service(
    session: Session, service_id: int, data: SubcontractedServiceUpdate
) -> Optional[SubcontractedService]:
    record = session.get(SubcontractedService, service_id)
    if not record:
        return None

    updates = data.model_dump(exclude_unset=True)
    for key, value in updates.items():
        setattr(record, key, value)

    session.add(record)
    session.commit()
    session.refresh(record)
    return record


def _normalize_order_week(value: Optional[str]) -> Optional[str]:
    if value is None:
        return None
    stripped = value.strip()
    return stripped.upper() if stripped else None


def _normalize_filter_dimensions(
    dimensions: Optional[str],
    format_type: str,
) -> Optional[str]:
    if not dimensions:
        return None

    numbers = re.findall(r"\d+(?:[.,]\d+)?", dimensions)
    if format_type == "cousus_sur_fil":
        if len(numbers) >= 2:
            return f"{numbers[0]} x {numbers[1]}"
    else:
        if len(numbers) >= 3:
            return f"{numbers[0]} x {numbers[1]} x {numbers[2]}"

    return dimensions.strip()


def list_filter_lines(session: Session) -> List[FilterLine]:
    stmt = select(FilterLine).order_by(FilterLine.created_at.desc())
    return session.exec(stmt).all()


def create_filter_line(session: Session, data: FilterLineCreate) -> FilterLine:
    payload = data.model_dump()
    payload["order_week"] = _normalize_order_week(payload.get("order_week"))
    payload["dimensions"] = _normalize_filter_dimensions(
        payload.get("dimensions"), payload["format_type"]
    )

    record = FilterLine(**payload)
    session.add(record)
    session.commit()
    session.refresh(record)
    return record


def delete_filter_line(session: Session, line_id: int) -> bool:
    record = session.get(FilterLine, line_id)
    if not record:
        return False
    session.delete(record)
    session.commit()
    return True


def get_filter_line(session: Session, line_id: int) -> Optional[FilterLine]:
    return session.get(FilterLine, line_id)


def update_filter_line(
    session: Session, line_id: int, data: FilterLineUpdate
) -> Optional[FilterLine]:
    record = session.get(FilterLine, line_id)
    if not record:
        return None

    updates = data.model_dump(exclude_unset=True)

    if "order_week" in updates:
        updates["order_week"] = _normalize_order_week(updates.get("order_week"))

    if "dimensions" in updates:
        format_type = updates.get("format_type") or record.format_type
        updates["dimensions"] = _normalize_filter_dimensions(
            updates.get("dimensions"), format_type
        )

    for key, value in updates.items():
        setattr(record, key, value)

    session.add(record)
    session.commit()
    session.refresh(record)
    return record


def list_belt_lines(session: Session) -> List[BeltLine]:
    stmt = select(BeltLine).order_by(BeltLine.created_at.desc())
    return session.exec(stmt).all()


def create_belt_line(session: Session, data: BeltLineCreate) -> BeltLine:
    payload = data.model_dump()
    payload["order_week"] = _normalize_order_week(payload.get("order_week"))

    record = BeltLine(**payload)
    session.add(record)
    session.commit()
    session.refresh(record)
    return record


def get_belt_line(session: Session, line_id: int) -> Optional[BeltLine]:
    return session.get(BeltLine, line_id)


def update_belt_line(
    session: Session, line_id: int, data: BeltLineUpdate
) -> Optional[BeltLine]:
    record = session.get(BeltLine, line_id)
    if not record:
        return None

    updates = data.model_dump(exclude_unset=True)

    if "order_week" in updates:
        updates["order_week"] = _normalize_order_week(updates.get("order_week"))

    for key, value in updates.items():
        setattr(record, key, value)

    session.add(record)
    session.commit()
    session.refresh(record)
    return record


def delete_belt_line(session: Session, line_id: int) -> bool:
    record = session.get(BeltLine, line_id)
    if not record:
        return False
    session.delete(record)
    session.commit()
    return True
