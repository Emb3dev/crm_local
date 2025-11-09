from typing import List, Optional
from datetime import datetime
from sqlalchemy import Column, String
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
    company_name: str = Field(index=True, description="Nom de l'entreprise")
    name: str = Field(index=True, description="Nom du client")
    email: Optional[str] = None
    phone: Optional[str] = None
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
    contacts: List["Contact"] = Relationship(
        back_populates="client",
        sa_relationship_kwargs={"cascade": "all, delete-orphan"},
    )
    subcontractings: List["SubcontractedService"] = Relationship(
        back_populates="client",
        sa_relationship_kwargs={"cascade": "all, delete-orphan"},
    )


class ClientCreate(ClientBase):
    pass


class ClientUpdate(SQLModel):
    company_name: Optional[str] = None
    name: Optional[str] = None
    email: Optional[str] = None
    phone: Optional[str] = None
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
# TABLE PRESTATIONS SOUS-TRAITÉES / LOCATIONS (rattachée à un client)
# =======================


class SubcontractedServiceBase(SQLModel):
    prestation_key: str = Field(index=True, description="Identifiant de la prestation")
    prestation_label: str = Field(description="Libellé de la prestation")
    category: str = Field(description="Famille de prestation (sous-traitance, location…)")
    budget_code: str = Field(description="Code budget associé")
    budget: Optional[float] = Field(default=None, description="Montant budgété")
    frequency: str = Field(description="Fréquence de la prestation")
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


class SubcontractedServiceCreate(SubcontractedServiceBase):
    pass


class SubcontractedServiceUpdate(SQLModel):
    prestation_key: Optional[str] = None
    prestation_label: Optional[str] = None
    category: Optional[str] = None
    budget_code: Optional[str] = None
    budget: Optional[float] = None
    frequency: Optional[str] = None
    status: Optional[str] = None
    realization_week: Optional[str] = None
    order_week: Optional[str] = None
