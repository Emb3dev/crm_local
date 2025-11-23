from typing import Dict, Iterable, List, Optional, Sequence

import re
from datetime import datetime

from sqlalchemy import delete, func, or_
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
    Supplier,
    SupplierContact,
    SupplierContactCreate,
    SupplierCreate,
    SupplierUpdate,
    SupplierCategory,
    FilterLine,
    FilterLineCreate,
    FilterLineUpdate,
    PrestationDefinition,
    PrestationDefinitionCreate,
    PrestationDefinitionUpdate,
    SubcontractedService,
    SubcontractedServiceComment,
    SubcontractedServiceCreate,
    SubcontractedServiceUpdate,
    User,
    UserCreate,
    UserLoginEvent,
    WorkloadCell,
    WorkloadCellUpdate,
    WorkloadSite,
    WorkloadSiteCreate,
    WorkloadSiteUpdate,
)


def list_prestation_definitions(session: Session) -> List[PrestationDefinition]:
    stmt = (
        select(PrestationDefinition)
        .order_by(
            PrestationDefinition.category.asc(),
            PrestationDefinition.position.asc(),
            PrestationDefinition.label.asc(),
        )
    )
    return session.exec(stmt).all()


def get_prestation_definition(
    session: Session, definition_id: int
) -> Optional[PrestationDefinition]:
    return session.get(PrestationDefinition, definition_id)


def get_prestation_definition_by_key(
    session: Session, key: str
) -> Optional[PrestationDefinition]:
    stmt = select(PrestationDefinition).where(PrestationDefinition.key == key)
    return session.exec(stmt).first()


def create_prestation_definition(
    session: Session, data: PrestationDefinitionCreate
) -> PrestationDefinition:
    definition = PrestationDefinition.model_validate(data)
    session.add(definition)
    session.commit()
    session.refresh(definition)
    return definition


def update_prestation_definition(
    session: Session, definition_id: int, data: PrestationDefinitionUpdate
) -> Optional[PrestationDefinition]:
    definition = session.get(PrestationDefinition, definition_id)
    if not definition:
        return None
    updates = data.model_dump(exclude_unset=True)
    for key, value in updates.items():
        setattr(definition, key, value)
    session.add(definition)
    session.commit()
    session.refresh(definition)
    return definition


def delete_prestation_definition(session: Session, definition_id: int) -> bool:
    definition = session.get(PrestationDefinition, definition_id)
    if not definition:
        return False
    session.delete(definition)
    session.commit()
    return True


def count_subcontracted_services_by_definition(
    session: Session,
) -> Dict[str, int]:
    stmt = (
        select(SubcontractedService.prestation_key, func.count(SubcontractedService.id))
        .group_by(SubcontractedService.prestation_key)
    )
    return {key: count for key, count in session.exec(stmt)}


def sync_subcontracted_services_from_definition(
    session: Session, definition: PrestationDefinition
) -> int:
    stmt = select(SubcontractedService).where(
        SubcontractedService.prestation_key == definition.key
    )
    services = session.exec(stmt).all()
    updated = 0
    for service in services:
        service.prestation_label = definition.label
        service.category = definition.category
        service.budget_code = definition.budget_code
        session.add(service)
        updated += 1
    if updated:
        session.commit()
    return updated

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


def find_clients_for_import(
    session: Session,
    *,
    company_name: Optional[str] = None,
    client_name: Optional[str] = None,
) -> List[Client]:
    stmt = select(Client).outerjoin(Entreprise)

    filters = []
    if company_name:
        normalized_company = company_name.strip().lower()
        if normalized_company:
            filters.append(func.lower(Entreprise.nom) == normalized_company)
            filters.append(func.lower(Client.company_name) == normalized_company)

    if filters:
        stmt = stmt.where(or_(*filters))

    if client_name:
        normalized_client = client_name.strip().lower()
        if normalized_client:
            stmt = stmt.where(func.lower(Client.name) == normalized_client)

    stmt = stmt.options(selectinload(Client.entreprise))
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


# =======================
# FOURNISSEURS
# =======================


def list_suppliers(
    session: Session,
    q: Optional[str] = None,
    *,
    supplier_type: Optional[str] = None,
    limit: int = 100,
) -> List[Supplier]:
    stmt = (
        select(Supplier)
        .options(selectinload(Supplier.contacts))
        .order_by(Supplier.created_at.desc())
    )
    if q:
        like = f"%{q}%"
        stmt = stmt.outerjoin(SupplierContact)
        stmt = stmt.where(
            (
                Supplier.name.ilike(like)
                | Supplier.our_code.ilike(like)
                | Supplier.categories.ilike(like)
                | SupplierContact.name.ilike(like)
                | SupplierContact.email.ilike(like)
                | SupplierContact.phone.ilike(like)
                | SupplierContact.description.ilike(like)
            )
        ).distinct()

    if supplier_type:
        stmt = stmt.where(Supplier.supplier_type == supplier_type)

    return session.exec(stmt.limit(limit)).all()


def get_supplier(session: Session, supplier_id: int) -> Optional[Supplier]:
    return session.get(Supplier, supplier_id)


def create_supplier(
    session: Session,
    data: SupplierCreate,
    contacts: Optional[List[SupplierContactCreate]] = None,
) -> Supplier:
    supplier = Supplier.model_validate(data)
    session.add(supplier)
    session.flush()

    for contact_data in contacts or []:
        contact = SupplierContact(supplier_id=supplier.id, **contact_data.model_dump())
        session.add(contact)

    session.commit()
    session.refresh(supplier)
    return supplier


def update_supplier(
    session: Session, supplier_id: int, data: SupplierUpdate
) -> Optional[Supplier]:
    supplier = session.get(Supplier, supplier_id)
    if not supplier:
        return None
    updates = data.model_dump(exclude_unset=True)
    for key, value in updates.items():
        setattr(supplier, key, value)
    session.add(supplier)
    session.commit()
    session.refresh(supplier)
    return supplier


def delete_supplier(session: Session, supplier_id: int) -> bool:
    supplier = session.get(Supplier, supplier_id)
    if not supplier:
        return False
    session.delete(supplier)
    session.commit()
    return True


def create_supplier_contact(
    session: Session, supplier_id: int, data: SupplierContactCreate
) -> Optional[SupplierContact]:
    supplier = session.get(Supplier, supplier_id)
    if not supplier:
        return None
    contact = SupplierContact(supplier_id=supplier_id, **data.model_dump())
    session.add(contact)
    session.commit()
    session.refresh(contact)
    return contact


def delete_supplier_contact(
    session: Session, supplier_id: int, contact_id: int
) -> bool:
    contact = session.get(SupplierContact, contact_id)
    if not contact or contact.supplier_id != supplier_id:
        return False
    session.delete(contact)
    session.commit()
    return True


def list_supplier_categories(session: Session) -> List[SupplierCategory]:
    stmt = select(SupplierCategory).order_by(SupplierCategory.label.asc())
    return session.exec(stmt).all()


def create_supplier_category(session: Session, label: str) -> SupplierCategory:
    normalized_label = (label or "").strip()
    if not normalized_label:
        raise ValueError("Le nom de la catégorie est requis")

    existing = session.exec(
        select(SupplierCategory).where(
            func.lower(SupplierCategory.label) == normalized_label.lower()
        )
    ).first()
    if existing:
        raise ValueError("Une catégorie portant ce nom existe déjà.")

    category = SupplierCategory(label=normalized_label)
    session.add(category)
    session.commit()
    session.refresh(category)
    return category


def update_supplier_category(
    session: Session, category_id: int, label: str
) -> SupplierCategory:
    normalized_label = (label or "").strip()
    if not normalized_label:
        raise ValueError("Le nom de la catégorie est requis")

    category = session.get(SupplierCategory, category_id)
    if not category:
        raise ValueError("Catégorie introuvable")

    previous_label = category.label

    existing = session.exec(
        select(SupplierCategory)
        .where(func.lower(SupplierCategory.label) == normalized_label.lower())
        .where(SupplierCategory.id != category_id)
    ).first()
    if existing:
        raise ValueError("Une catégorie portant ce nom existe déjà.")

    category.label = normalized_label
    session.add(category)
    session.commit()
    session.refresh(category)

    _propagate_supplier_category_update(
        session, previous_label=previous_label, new_label=normalized_label
    )
    return category


def _split_categories(value: Optional[str]) -> List[str]:
    if not value:
        return []
    cleaned = value.replace(";", ",")
    return [part.strip() for part in cleaned.split(",") if part.strip()]


def _normalize_categories(value: Sequence[str]) -> Optional[str]:
    unique_items: List[str] = []
    seen: set[str] = set()
    for item in value:
        cleaned = (item or "").strip()
        if not cleaned:
            continue
        canonical = cleaned.lower()
        if canonical in seen:
            continue
        seen.add(canonical)
        unique_items.append(cleaned)

    if not unique_items:
        return None
    return ", ".join(unique_items)


def _propagate_supplier_category_update(
    session: Session, *, previous_label: Optional[str], new_label: str
) -> None:
    if not previous_label:
        return

    lower_previous = previous_label.lower()
    suppliers = session.exec(
        select(Supplier).where(Supplier.categories.is_not(None))
    ).all()

    updated = False
    for supplier in suppliers:
        categories = _split_categories(supplier.categories)
        if not categories:
            continue

        updated_categories = []
        changed = False
        for item in categories:
            if item.lower() == lower_previous:
                updated_categories.append(new_label)
                changed = True
            else:
                updated_categories.append(item)

        normalized = _normalize_categories(updated_categories)
        if changed and normalized != supplier.categories:
            supplier.categories = normalized
            session.add(supplier)
            updated = True

    if updated:
        session.commit()


def delete_supplier_category(session: Session, category_id: int) -> bool:
    category = session.get(SupplierCategory, category_id)
    if not category:
        return False

    session.delete(category)
    session.commit()
    return True


def ensure_supplier_categories(session: Session, labels: List[str]) -> None:
    normalized_labels = []
    for label in labels:
        cleaned = (label or "").strip()
        if cleaned and cleaned not in normalized_labels:
            normalized_labels.append(cleaned)

    if not normalized_labels:
        return

    lowered = [label.lower() for label in normalized_labels]
    existing_labels = {
        category.label.lower()
        for category in session.exec(
            select(SupplierCategory).where(func.lower(SupplierCategory.label).in_(lowered))
        )
    }

    new_categories = [
        SupplierCategory(label=label)
        for label in normalized_labels
        if label.lower() not in existing_labels
    ]

    if not new_categories:
        return

    session.add_all(new_categories)
    session.commit()


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
    records = session.exec(stmt.limit(limit)).all()

    if effective_filters.get("order_status") == "overdue":
        current_week = datetime.utcnow().isocalendar().week
        filtered = []
        for record in records:
            if record.status == "fait":
                continue
            week_number = _parse_week_number(record.order_week)
            if week_number is not None and week_number < current_week:
                filtered.append(record)
        records = filtered

    return records


def get_subcontracted_service(
    session: Session, service_id: int
) -> Optional[SubcontractedService]:
    stmt = (
        select(SubcontractedService)
        .where(SubcontractedService.id == service_id)
        .options(selectinload(SubcontractedService.client))
    )
    return session.exec(stmt).one_or_none()


def list_subcontracted_service_comments(
    session: Session, service_id: int
) -> List[SubcontractedServiceComment]:
    stmt = (
        select(SubcontractedServiceComment)
        .where(SubcontractedServiceComment.service_id == service_id)
        .order_by(SubcontractedServiceComment.created_at.desc())
    )
    return session.exec(stmt).all()


def list_subcontracted_service_comments_by_service(
    session: Session, service_ids: Sequence[int]
) -> Dict[int, List[SubcontractedServiceComment]]:
    if not service_ids:
        return {}

    stmt = (
        select(SubcontractedServiceComment)
        .where(SubcontractedServiceComment.service_id.in_(service_ids))
        .order_by(
            SubcontractedServiceComment.service_id.asc(),
            SubcontractedServiceComment.created_at.desc(),
        )
    )
    comments = session.exec(stmt).all()

    grouped: Dict[int, List[SubcontractedServiceComment]] = {sid: [] for sid in service_ids}
    for comment in comments:
        grouped.setdefault(comment.service_id, []).append(comment)
    return grouped


def create_subcontracted_service_comment(
    session: Session,
    *,
    service_id: int,
    author_name: str,
    author_initials: str,
    content: str,
) -> Optional[SubcontractedServiceComment]:
    service = session.get(SubcontractedService, service_id)
    if not service:
        return None

    comment = SubcontractedServiceComment(
        service_id=service_id,
        author_name=author_name,
        author_initials=author_initials,
        content=content,
    )
    session.add(comment)
    session.commit()
    session.refresh(comment)
    return comment


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


def _parse_week_number(value: Optional[str]) -> Optional[int]:
    if not value:
        return None
    match = re.search(r"(\d{1,2})", value)
    if not match:
        return None
    week_number = int(match.group(1))
    if 1 <= week_number <= 53:
        return week_number
    return None


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


def list_filter_lines(
    session: Session, q: Optional[str] = None
) -> List[FilterLine]:
    stmt = select(FilterLine).order_by(FilterLine.created_at.desc())
    if q:
        like = f"%{q}%"
        stmt = stmt.where(
            FilterLine.site.ilike(like)
            | FilterLine.equipment.ilike(like)
            | FilterLine.efficiency.ilike(like)
            | FilterLine.dimensions.ilike(like)
            | FilterLine.order_week.ilike(like)
            | FilterLine.format_type.ilike(like)
        )
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


def list_belt_lines(session: Session, q: Optional[str] = None) -> List[BeltLine]:
    stmt = select(BeltLine).order_by(BeltLine.created_at.desc())
    if q:
        like = f"%{q}%"
        stmt = stmt.where(
            BeltLine.site.ilike(like)
            | BeltLine.equipment.ilike(like)
            | BeltLine.reference.ilike(like)
            | BeltLine.order_week.ilike(like)
        )
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


def get_user_by_username(session: Session, username: str) -> Optional[User]:
    stmt = select(User).where(User.username == username)
    return session.exec(stmt).first()


def create_user(session: Session, data: UserCreate) -> User:
    user = User.model_validate(data)
    session.add(user)
    session.commit()
    session.refresh(user)
    return user


def list_users(session: Session) -> List[User]:
    stmt = select(User).order_by(User.username.asc())
    return session.exec(stmt).all()


def update_user_password(
    session: Session, user: User, hashed_password: str
) -> User:
    user.hashed_password = hashed_password
    session.add(user)
    session.commit()
    session.refresh(user)
    return user


def record_user_login(session: Session, user: User) -> User:
    if not user.id:
        return user
    now = datetime.utcnow()
    user.last_login_at = now
    user.last_active_at = now
    user.is_online = True
    event = UserLoginEvent(user_id=user.id, event_type="login", occurred_at=now)
    session.add_all([user, event])
    session.commit()
    session.refresh(user)
    return user


def record_user_logout(session: Session, user: User) -> User:
    if not user.id:
        return user
    now = datetime.utcnow()
    user.last_logout_at = now
    user.is_online = False
    event = UserLoginEvent(user_id=user.id, event_type="logout", occurred_at=now)
    session.add_all([user, event])
    session.commit()
    session.refresh(user)
    return user


def touch_user_activity(session: Session, user: User) -> User:
    if not user.id:
        return user
    user.last_active_at = datetime.utcnow()
    user.is_online = True
    session.add(user)
    session.commit()
    session.refresh(user)
    return user


def get_login_history_for_users(
    session: Session, user_ids: List[int], *, limit: int = 5
) -> Dict[int, List[UserLoginEvent]]:
    history: Dict[int, List[UserLoginEvent]] = {}
    if not user_ids:
        return history
    for user_id in user_ids:
        stmt = (
            select(UserLoginEvent)
            .where(UserLoginEvent.user_id == user_id)
            .order_by(UserLoginEvent.occurred_at.desc())
            .limit(limit)
        )
        history[user_id] = session.exec(stmt).all()
    return history
