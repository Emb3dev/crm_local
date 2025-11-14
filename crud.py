from typing import Dict, Iterable, List, Optional

import re

from sqlalchemy import delete, func
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
    Entreprise,
    EntrepriseCreate,
    EntrepriseUpdate,
    FilterLine,
    FilterLineCreate,
    FilterLineUpdate,
    SubcontractedService,
    SubcontractedServiceCreate,
    SubcontractedServiceUpdate,
    WorkloadCell,
    WorkloadCellUpdate,
    WorkloadSite,
    WorkloadSiteCreate,
    WorkloadSiteUpdate,
)

def get_entreprise(session: Session, entreprise_id: int) -> Optional[Entreprise]:
    return session.get(Entreprise, entreprise_id)


def get_entreprise_by_name(session: Session, name: str) -> Optional[Entreprise]:
    stmt = select(Entreprise).where(Entreprise.nom == name)
    return session.exec(stmt).first()


def list_entreprises(session: Session) -> List[Entreprise]:
    stmt = select(Entreprise).order_by(Entreprise.nom.asc())
    return session.exec(stmt).all()


def create_entreprise(session: Session, data: EntrepriseCreate) -> Entreprise:
    entreprise = Entreprise.model_validate(data)
    session.add(entreprise)
    session.commit()
    session.refresh(entreprise)
    return entreprise


def update_entreprise(
    session: Session, entreprise_id: int, data: EntrepriseUpdate
) -> Optional[Entreprise]:
    entreprise = session.get(Entreprise, entreprise_id)
    if not entreprise:
        return None
    updates = data.model_dump(exclude_unset=True)
    for key, value in updates.items():
        setattr(entreprise, key, value)
    session.add(entreprise)
    session.commit()
    session.refresh(entreprise)
    return entreprise


def ensure_entreprise(
    session: Session,
    *,
    name: str,
    adresse_facturation: Optional[str] = None,
    tag: Optional[str] = None,
    statut: Optional[bool] = None,
) -> Entreprise:
    normalized = name.strip()
    entreprise = get_entreprise_by_name(session, normalized)
    payload = {}
    if adresse_facturation is not None:
        payload["adresse_facturation"] = adresse_facturation or None
    if tag is not None:
        payload["tag"] = tag or None
    if statut is not None:
        payload["statut"] = statut

    if entreprise:
        if payload:
            update = EntrepriseUpdate(**payload)
            update_entreprise(session, entreprise.id, update)
            session.refresh(entreprise)
        return entreprise

    create_payload = EntrepriseCreate(
        nom=normalized,
        adresse_facturation=payload.get("adresse_facturation"),
        tag=payload.get("tag"),
        statut=payload.get("statut", True),
    )
    entreprise = create_entreprise(session, create_payload)
    return entreprise


def list_clients(
    session: Session,
    q: Optional[str] = None,
    *,
    filters: Optional[Dict[str, str]] = None,
    limit: int = 50,
) -> List[Client]:
    stmt = (
        select(Client)
        .join(Entreprise)
        .options(
            selectinload(Client.contacts),
            selectinload(Client.subcontractings),
            selectinload(Client.entreprise),
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
                | Client.email.ilike(like)
                | Client.phone.ilike(like)
                | Client.billing_address.ilike(like)
                | Client.depannage.ilike(like)
                | Client.astreinte.ilike(like)
                | Client.tags.ilike(like)
                | Entreprise.nom.ilike(like)
                | Entreprise.tag.ilike(like)
                | Contact.name.ilike(like)
                | Contact.email.ilike(like)
                | Contact.phone.ilike(like)
            )
        )
        stmt = stmt.distinct()

    status = effective_filters.get("status")
    if status:
        stmt = stmt.where(Entreprise.statut == (status == "actif"))

    depannage = effective_filters.get("depannage")
    if depannage:
        stmt = stmt.where(Client.depannage == depannage)

    astreinte = effective_filters.get("astreinte")
    if astreinte:
        stmt = stmt.where(Client.astreinte == astreinte)

    return session.exec(stmt.limit(limit)).all()


def list_client_choices(session: Session) -> List[Client]:
    stmt = (
        select(Client)
        .options(selectinload(Client.entreprise))
        .join(Entreprise)
        .order_by(Entreprise.nom.asc(), Client.name.asc())
    )
    return session.exec(stmt).all()

def get_client(session: Session, client_id: int) -> Optional[Client]:
    return session.get(Client, client_id)

def create_client(
    session: Session,
    data: ClientCreate,
    contacts: Optional[List[ContactCreate]] = None,
) -> Client:
    c = Client.model_validate(data)
    if c.entreprise_id:
        entreprise = session.get(Entreprise, c.entreprise_id)
        if entreprise:
            c.company_name = entreprise.nom
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
    entreprise_id = updates.pop("entreprise_id", None)
    if entreprise_id is not None:
        entreprise = session.get(Entreprise, entreprise_id)
        if entreprise:
            c.entreprise_id = entreprise_id
            c.company_name = entreprise.nom
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


def list_workload_sites(session: Session) -> List[WorkloadSite]:
    stmt = (
        select(WorkloadSite)
        .options(selectinload(WorkloadSite.cells))
        .order_by(WorkloadSite.position.asc(), WorkloadSite.id.asc())
    )
    return session.exec(stmt).all()


def create_workload_site(session: Session, data: WorkloadSiteCreate) -> WorkloadSite:
    normalized = data.name.strip()
    if not normalized:
        raise ValueError("Le nom du site est requis")

    existing = session.exec(
        select(WorkloadSite).where(WorkloadSite.name == normalized)
    ).first()
    if existing:
        raise ValueError("Ce site existe déjà")

    max_position = session.exec(
        select(func.max(WorkloadSite.position))
    ).one_or_none()
    if max_position is None:
        max_position = 0

    position = max_position + 1

    site = WorkloadSite(name=normalized, position=position)
    session.add(site)
    session.commit()
    session.refresh(site)
    return site


def rename_workload_site(
    session: Session, site_id: int, data: WorkloadSiteUpdate
) -> Optional[WorkloadSite]:
    site = session.get(WorkloadSite, site_id)
    if not site:
        return None

    updates = data.model_dump(exclude_unset=True)
    new_name = updates.get("name")

    if new_name is not None:
        normalized = new_name.strip()
        if not normalized:
            raise ValueError("Le nom du site est requis")
        exists = session.exec(
            select(WorkloadSite)
            .where(WorkloadSite.name == normalized)
            .where(WorkloadSite.id != site_id)
        ).first()
        if exists:
            raise ValueError("Ce nom est déjà utilisé")
        site.name = normalized

    if "position" in updates and updates["position"] is not None:
        site.position = updates["position"]

    session.add(site)
    session.commit()
    session.refresh(site)
    return site


def delete_workload_site(session: Session, site_id: int) -> bool:
    site = session.get(WorkloadSite, site_id)
    if not site:
        return False
    session.delete(site)
    session.commit()
    return True


def bulk_update_workload_cells(
    session: Session, updates: Iterable[WorkloadCellUpdate]
) -> int:
    items = list(updates)
    if not items:
        return 0

    site_ids = {item.site_id for item in items}
    if not site_ids:
        return 0

    existing_ids = set(
        session.exec(select(WorkloadSite.id).where(WorkloadSite.id.in_(site_ids))).all()
    )
    missing = site_ids - existing_ids
    if missing:
        raise ValueError("Site introuvable")

    for item in items:
        if not 0 <= item.day_index < 364:
            raise ValueError("Indice de jour invalide")

        normalized = (item.value or "").strip()
        cell = session.exec(
            select(WorkloadCell)
            .where(WorkloadCell.site_id == item.site_id)
            .where(WorkloadCell.day_index == item.day_index)
        ).first()

        if normalized:
            if cell:
                cell.value = normalized
            else:
                cell = WorkloadCell(
                    site_id=item.site_id, day_index=item.day_index, value=normalized
                )
                session.add(cell)
        else:
            if cell:
                session.delete(cell)

    session.commit()
    return len(items)


def replace_workload_plan(
    session: Session,
    site_names: Iterable[str],
    cells_map: Dict[str, Iterable[Optional[str]]],
) -> int:
    normalized_cells = {
        (name or "").strip(): list(values)
        for name, values in (cells_map or {}).items()
        if name is not None
    }

    session.exec(delete(WorkloadCell))
    session.exec(delete(WorkloadSite))
    session.commit()

    created = 0
    seen_names = set()
    for position, raw_name in enumerate(site_names):
        normalized = (raw_name or "").strip()
        if not normalized or normalized in seen_names:
            continue
        seen_names.add(normalized)

        site = WorkloadSite(name=normalized, position=position + 1)
        session.add(site)
        session.flush()

        values = normalized_cells.get(normalized, [])
        for day_index, value in enumerate(values[:364]):
            normalized_value = (value or "").strip()
            if normalized_value:
                session.add(
                    WorkloadCell(
                        site_id=site.id,
                        day_index=day_index,
                        value=normalized_value,
                    )
                )

        created += 1

    session.commit()
    return created


def list_subcontracted_services(
    session: Session,
    q: Optional[str] = None,
    *,
    filters: Optional[Dict[str, str]] = None,
    limit: int = 200,
) -> List[SubcontractedService]:
    stmt = (
        select(SubcontractedService)
        .options(
            selectinload(SubcontractedService.client).selectinload(Client.entreprise)
        )
        .order_by(SubcontractedService.created_at.desc())
    )
    effective_filters = filters or {}
    if q:
        like = f"%{q}%"
        stmt = stmt.outerjoin(Client).outerjoin(Entreprise)
        stmt = stmt.where(
            (
                SubcontractedService.prestation_label.ilike(like)
                | SubcontractedService.category.ilike(like)
                | SubcontractedService.budget_code.ilike(like)
                | Client.name.ilike(like)
                | Entreprise.nom.ilike(like)
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
