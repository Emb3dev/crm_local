from typing import Optional, List
from datetime import datetime
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


    # une entreprise a plusieurs clients
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
    nom: str = Field(index=True)
    email: Optional[str] = None
    telephone: Optional[str] = None
    depannage_refacturable: bool = Field(default=True)
    astreinte: str = Field(
    default="pas_d_astreinte",
    description="incluse_non_refacturable | incluse_refacturable | pas_d_astreinte"
)


class Client(ClientBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    entreprise_id: Optional[int] = Field(default=None, foreign_key="entreprise.id")
    entreprise: Optional[Entreprise] = Relationship(back_populates="clients")
    created_at: datetime = Field(default_factory=datetime.utcnow)


class ClientCreate(ClientBase):
    entreprise_id: int


class ClientUpdate(SQLModel):
    nom: Optional[str] = None
    email: Optional[str] = None
    telephone: Optional[str] = None
    depannage_refacturable: Optional[bool] = None
    astreinte: Optional[str] = None
    entreprise_id: Optional[int] = None