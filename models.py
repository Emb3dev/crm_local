from typing import List, Optional
from datetime import datetime
from sqlalchemy import (
    Column,
    String,
    Integer,
    UniqueConstraint,
    Boolean,
    DateTime,
)
from sqlmodel import SQLModel, Field, Relationship


# =======================
# TABLE ENTREPRISE
# =======================
class EntrepriseBase(SQLModel):
    nom: str = Field(index=True, description="Nom de l'entreprise")
    adresse_facturation: Optional[str] = None
    tag: Optional[str] = None
    statut: bool = Field(default=True, description="Actif ou non")


class Entreprise(EntrepriseBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    created_at: datetime = Field(default_factory=datetime.utcnow)
    clients: List["Client"] = Relationship(back_populates="entreprise")


class EntrepriseCreate(EntrepriseBase):
    pass


class EntrepriseUpdate(SQLModel):
    nom: Optional[str] = None
    adresse_facturation: Optional[str] = None
    tag: Optional[str] = None
    statut: Optional[bool] = None




# =======================
# TABLE CLIENT (rattachée à une entreprise)
# =======================
class ClientBase(SQLModel):
    name: str = Field(index=True, description="Nom du client")
    email: Optional[str] = None
    phone: Optional[str] = None
    technician_name: Optional[str] = Field(
        default=None,
        description="Nom du technicien référent",
    )
    billing_address: Optional[str] = Field(
        default=None,
        sa_column=Column("address", String, nullable=True),
        description="Adresse de facturation",
    )
    depannage: str = Field(
        default="non_refacturable",
        sa_column=Column("siret", String, nullable=True),
        description="Type de dépannage",
    )
    astreinte: str = Field(default="pas_d_astreinte", description="Type d'astreinte")
    tags: Optional[str] = None
    status: Optional[str] = Field(default="actif")


class Client(ClientBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    created_at: datetime = Field(default_factory=datetime.utcnow)
    entreprise_id: Optional[int] = Field(
        default=None,
        foreign_key="entreprise.id",
        description="Entreprise rattachée au client",
    )
    entreprise: Optional[Entreprise] = Relationship(back_populates="clients")
    company_name: Optional[str] = Field(
        default=None,
        description="Nom d'entreprise dénormalisé (compatibilité historique)",
    )
    contacts: List["Contact"] = Relationship(
        back_populates="client",
        sa_relationship_kwargs={"cascade": "all, delete-orphan"},
    )
    subcontractings: List["SubcontractedService"] = Relationship(
        back_populates="client",
        sa_relationship_kwargs={"cascade": "all, delete-orphan"},
    )


class ClientCreate(ClientBase):
    entreprise_id: int


class ClientUpdate(SQLModel):
    entreprise_id: Optional[int] = None
    name: Optional[str] = None
    email: Optional[str] = None
    phone: Optional[str] = None
    technician_name: Optional[str] = None
    billing_address: Optional[str] = None
    depannage: Optional[str] = None
    astreinte: Optional[str] = None
    tags: Optional[str] = None
    status: Optional[str] = None


# =======================
# TABLE CONTACT (rattachée à un client)
# =======================
class ContactBase(SQLModel):
    name: str = Field(index=True, description="Nom du contact")
    email: Optional[str] = None
    phone: Optional[str] = None


class Contact(ContactBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    client_id: int = Field(foreign_key="client.id")
    created_at: datetime = Field(default_factory=datetime.utcnow)
    client: Optional[Client] = Relationship(back_populates="contacts")


class ContactCreate(ContactBase):
    pass


class ContactUpdate(SQLModel):
    name: Optional[str] = None
    email: Optional[str] = None
    phone: Optional[str] = None


# =======================
# TABLE FOURNISSEUR
# =======================


class SupplierBase(SQLModel):
    name: str = Field(index=True, description="Nom du fournisseur")
    our_code: Optional[str] = Field(
        default=None, description="Code interne utilisé pour le fournisseur"
    )
    supplier_type: str = Field(default="fournisseur", description="Type de partenaire")
    categories: Optional[str] = Field(
        default=None,
        description="Liste de catégories associées au fournisseur",
    )


class Supplier(SupplierBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    created_at: datetime = Field(default_factory=datetime.utcnow)
    contacts: List["SupplierContact"] = Relationship(
        back_populates="supplier",
        sa_relationship_kwargs={"cascade": "all, delete-orphan"},
    )


class SupplierCreate(SupplierBase):
    pass


class SupplierUpdate(SQLModel):
    name: Optional[str] = None
    our_code: Optional[str] = None
    supplier_type: Optional[str] = None
    categories: Optional[str] = None


class SupplierContactBase(SQLModel):
    name: str = Field(index=True, description="Nom du contact fournisseur")
    email: Optional[str] = None
    phone: Optional[str] = None
    description: Optional[str] = Field(
        default=None, description="Informations complémentaires sur le contact"
    )


class SupplierContact(SupplierContactBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    supplier_id: int = Field(foreign_key="supplier.id")
    created_at: datetime = Field(default_factory=datetime.utcnow)
    supplier: Optional[Supplier] = Relationship(back_populates="contacts")


class SupplierContactCreate(SupplierContactBase):
    pass


class SupplierContactUpdate(SQLModel):
    name: Optional[str] = None
    email: Optional[str] = None
    phone: Optional[str] = None
    description: Optional[str] = None


class SupplierCategory(SQLModel, table=True):
    __table_args__ = (UniqueConstraint("label"),)

    id: Optional[int] = Field(default=None, primary_key=True)
    label: str = Field(index=True, description="Nom de la catégorie fournisseur")
    created_at: datetime = Field(default_factory=datetime.utcnow)


# =======================
# TABLE PRESTATIONS SOUS-TRAITÉES / LOCATIONS (rattachée à un client)
# =======================


class PrestationDefinitionBase(SQLModel):
    key: str = Field(index=True, description="Identifiant interne de la prestation")
    label: str = Field(description="Libellé affiché dans les menus")
    budget_code: str = Field(description="Code budget associé à la prestation")
    category: str = Field(description="Famille de prestation (sous-traitance, location…)")
    position: int = Field(
        default=0,
        description="Ordre d'affichage dans la catégorie",
        sa_column=Column(Integer, nullable=False, server_default="0"),
    )


class PrestationDefinition(PrestationDefinitionBase, table=True):
    __table_args__ = (
        UniqueConstraint("key"),
        UniqueConstraint("category", "label"),
    )

    id: Optional[int] = Field(default=None, primary_key=True)
    created_at: datetime = Field(default_factory=datetime.utcnow)


class PrestationDefinitionCreate(PrestationDefinitionBase):
    pass


class PrestationDefinitionUpdate(SQLModel):
    label: Optional[str] = None
    budget_code: Optional[str] = None
    category: Optional[str] = None
    position: Optional[int] = None


class SubcontractedServiceBase(SQLModel):
    prestation_key: str = Field(index=True, description="Identifiant de la prestation")
    prestation_label: str = Field(description="Libellé de la prestation")
    category: str = Field(description="Famille de prestation (sous-traitance, location…)")
    budget_code: str = Field(description="Code budget associé")
    budget: Optional[float] = Field(default=None, description="Montant budgété")
    frequency: str = Field(description="Fréquence de la prestation")
    frequency_interval: Optional[int] = Field(
        default=None,
        description="Intervalle de récurrence (par exemple tous les X mois/années)",
        sa_column=Column("frequency_interval", Integer, nullable=True),
    )
    frequency_unit: Optional[str] = Field(
        default=None,
        description="Unité associée à l'intervalle de fréquence",
        sa_column=Column("frequency_unit", String, nullable=True),
    )
    status: str = Field(
        default="non_commence",
        description="Statut d'avancement de la prestation",
        sa_column=Column(
            "status",
            String,
            nullable=False,
            server_default="non_commence",
        ),
    )
    realization_week: Optional[str] = Field(
        default=None, description="Semaine de réalisation (format S01, S02…)"
    )
    order_week: Optional[str] = Field(
        default=None, description="Semaine de commande (format S01, S02…)"
    )


class SubcontractedService(SubcontractedServiceBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    client_id: int = Field(foreign_key="client.id")
    created_at: datetime = Field(default_factory=datetime.utcnow)
    client: Optional[Client] = Relationship(back_populates="subcontractings")
    comments: List["SubcontractedServiceComment"] = Relationship(
        back_populates="service",
        sa_relationship_kwargs={"cascade": "all, delete-orphan"},
    )


class SubcontractedServiceCreate(SubcontractedServiceBase):
    pass


class SubcontractedServiceUpdate(SQLModel):
    prestation_key: Optional[str] = None
    prestation_label: Optional[str] = None
    category: Optional[str] = None
    budget_code: Optional[str] = None
    budget: Optional[float] = None
    frequency: Optional[str] = None
    frequency_interval: Optional[int] = None
    frequency_unit: Optional[str] = None
    status: Optional[str] = None
    realization_week: Optional[str] = None
    order_week: Optional[str] = None
    client_id: Optional[int] = None


class SubcontractedServiceCommentBase(SQLModel):
    service_id: int = Field(foreign_key="subcontractedservice.id")
    author_name: str = Field(description="Nom ou prénom de l'auteur du commentaire")
    author_initials: str = Field(description="Initiales de l'auteur")
    content: str = Field(sa_column=Column("content", String, nullable=False))
    created_at: datetime = Field(default_factory=datetime.utcnow)


class SubcontractedServiceComment(SubcontractedServiceCommentBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    service: Optional[SubcontractedService] = Relationship(back_populates="comments")


# =======================
# TABLE PLAN DE CHARGE
# =======================


class WorkloadSiteBase(SQLModel):
    name: str = Field(
        description="Nom du site",
        sa_column=Column("name", String, unique=True, nullable=False),
    )
    position: int = Field(default=0, description="Position d'affichage du site")


class WorkloadSite(WorkloadSiteBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    cells: List["WorkloadCell"] = Relationship(
        back_populates="site",
        sa_relationship_kwargs={"cascade": "all, delete-orphan"},
    )


class WorkloadSiteCreate(SQLModel):
    name: str


class WorkloadSiteUpdate(SQLModel):
    name: Optional[str] = None
    position: Optional[int] = None


class WorkloadCellBase(SQLModel):
    day_index: int = Field(index=True, description="Indice du jour (0-363)")
    value: Optional[str] = Field(default=None, description="Valeur stockée pour le jour")


class WorkloadCell(WorkloadCellBase, table=True):
    __table_args__ = (
        UniqueConstraint("site_id", "day_index", name="uq_workloadcell_site_day"),
    )

    id: Optional[int] = Field(default=None, primary_key=True)
    site_id: int = Field(foreign_key="workloadsite.id")
    site: Optional[WorkloadSite] = Relationship(back_populates="cells")


class WorkloadCellUpdate(SQLModel):
    site_id: int
    day_index: int
    value: Optional[str] = None


# =======================
# TABLE LIGNES FILTRES
# =======================


class FilterLineBase(SQLModel):
    site: str = Field(index=True, description="Nom du site")
    equipment: str = Field(index=True, description="Équipement concerné")
    efficiency: Optional[str] = Field(default=None, description="Classe d'efficacité du filtre")
    format_type: str = Field(
        description="Format du filtre (cousus sur fil, cadre…)",
        sa_column=Column("filter_type", String, nullable=False),
    )
    dimensions: Optional[str] = Field(default=None, description="Dimensions du filtre")
    quantity: int = Field(default=1, description="Quantité requise pour ce filtre")
    order_week: Optional[str] = Field(
        default=None, description="Semaine de commande (format S01, S02…)"
    )


class FilterLine(FilterLineBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    created_at: datetime = Field(default_factory=datetime.utcnow)


class FilterLineCreate(FilterLineBase):
    pass


class FilterLineUpdate(SQLModel):
    site: Optional[str] = None
    equipment: Optional[str] = None
    efficiency: Optional[str] = None
    format_type: Optional[str] = None
    dimensions: Optional[str] = None
    quantity: Optional[int] = None
    order_week: Optional[str] = None


# =======================
# TABLE LIGNES COURROIES
# =======================


class BeltLineBase(SQLModel):
    site: str = Field(index=True, description="Nom du site")
    equipment: str = Field(index=True, description="Équipement concerné")
    reference: str = Field(description="Référence de la courroie")
    quantity: int = Field(default=1, description="Quantité requise")
    order_week: Optional[str] = Field(default=None, description="Semaine de commande (S01…)")


class BeltLine(BeltLineBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    created_at: datetime = Field(default_factory=datetime.utcnow)


class BeltLineCreate(BeltLineBase):
    pass


class BeltLineUpdate(SQLModel):
    site: Optional[str] = None
    equipment: Optional[str] = None
    reference: Optional[str] = None
    quantity: Optional[int] = None
    order_week: Optional[str] = None


# =======================
# TABLE UTILISATEURS
# =======================


class UserBase(SQLModel):
    username: str = Field(index=True, unique=True, description="Identifiant de connexion")


class User(UserBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    hashed_password: str = Field(description="Mot de passe haché")
    created_at: datetime = Field(default_factory=datetime.utcnow)
    last_login_at: Optional[datetime] = Field(
        default=None,
        sa_column=Column(DateTime, nullable=True),
        description="Dernière connexion réussie",
    )
    last_logout_at: Optional[datetime] = Field(
        default=None,
        sa_column=Column(DateTime, nullable=True),
        description="Dernière déconnexion",
    )
    last_active_at: Optional[datetime] = Field(
        default=None,
        sa_column=Column(DateTime, nullable=True),
        description="Dernière activité détectée",
    )
    is_online: bool = Field(
        default=False,
        sa_column=Column(Boolean, nullable=False, server_default="0"),
        description="Indique si l'utilisateur est actuellement en ligne",
    )
    login_events: List["UserLoginEvent"] = Relationship(back_populates="user")


class UserCreate(UserBase):
    hashed_password: str


class UserLoginEvent(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    user_id: int = Field(foreign_key="user.id")
    event_type: str = Field(description="Type d'événement (login/logout)")
    occurred_at: datetime = Field(
        default_factory=datetime.utcnow,
        sa_column=Column(DateTime, nullable=False),
        description="Horodatage de l'événement",
    )
    user: Optional[User] = Relationship(back_populates="login_events")
