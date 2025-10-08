
from __future__ import annotations

from datetime import datetime
from typing import List, Optional

from sqlalchemy import Boolean, Column, String, Text
from sqlmodel import Field, Relationship, SQLModel


class ContactBase(SQLModel):
    """Contact rattaché à un client entreprise."""

    name: str = Field(description="Nom du contact")
    email: Optional[str] = Field(default=None, description="Email du contact")
    phone: Optional[str] = Field(default=None, description="Téléphone du contact")


class Contact(ContactBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    client_id: Optional[int] = Field(default=None, foreign_key="client.id", index=True)
    client: Optional["Client"] = Relationship(back_populates="contacts")


class ContactInput(ContactBase):
    pass


class ClientBase(SQLModel):
    company_name: str = Field(
        sa_column=Column("name", String, index=True, nullable=False),
        description="Nom de l'entreprise",
    )
    billing_address: Optional[str] = Field(
        default=None,
        sa_column=Column("address", Text, nullable=True),
        description="Adresse de facturation",
    )
    depannage: str = Field(
        default="non_refacturable",
        sa_column=Column("depannage", String, nullable=False, default="non_refacturable"),
        description="Dépannage refacturable ou non",
    )
    astreinte: str = Field(
        default="pas_d_astreinte",
        sa_column=Column("astreinte", String, nullable=False, default="pas_d_astreinte"),
        description="Type d'astreinte pour le client",
    )
    tag: Optional[str] = Field(
        default=None,
        sa_column=Column("tags", String, nullable=True),
        description="Tag libre",
    )
    is_active: bool = Field(
        default=True,
        sa_column=Column("is_active", Boolean, nullable=False, server_default="1"),
        description="Client actif ou non",
    )


class Client(ClientBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    created_at: datetime = Field(default_factory=datetime.utcnow, index=True)
    contacts: List[Contact] = Relationship(
        back_populates="client",
        sa_relationship_kwargs={"cascade": "all, delete-orphan"},
    )


class ClientCreate(ClientBase):
    contacts: List[ContactInput] = Field(default_factory=list)


class ClientUpdate(SQLModel):
    company_name: Optional[str] = None
    billing_address: Optional[str] = None
    depannage: Optional[str] = None
    astreinte: Optional[str] = None
    tag: Optional[str] = None
    is_active: Optional[bool] = None
    contacts: Optional[List[ContactInput]] = None
