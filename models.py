from typing import Optional
from datetime import datetime
from sqlmodel import SQLModel, Field


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
    name: str = Field(index=True)
    email: Optional[str] = None
    phone: Optional[str] = None
    siret: Optional[str] = None
    address: Optional[str] = None
    tags: Optional[str] = None
    status: Optional[str] = Field(default="prospect")


class Client(ClientBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    created_at: datetime = Field(default_factory=datetime.utcnow)


class ClientCreate(ClientBase):
    pass


class ClientUpdate(SQLModel):
    name: Optional[str] = None
    email: Optional[str] = None
    phone: Optional[str] = None
    siret: Optional[str] = None
    address: Optional[str] = None
    tags: Optional[str] = None
    status: Optional[str] = None
