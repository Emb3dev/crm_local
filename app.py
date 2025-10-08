from itertools import zip_longest
from typing import List

from fastapi import Depends, FastAPI, Form, HTTPException, Request
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlmodel import Session

import crud
from database import get_session, init_db
from models import ClientCreate, ClientUpdate, ContactInput

app = FastAPI(title="CRM Local")
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")


@app.on_event("startup")
def on_startup() -> None:
    init_db()


def _as_bool(value: str | None, default: bool = True) -> bool:
    if value is None:
        return default
    return value.lower() in {"1", "true", "on", "oui", "yes"}


def _build_contacts(
    names: List[str], emails: List[str], phones: List[str]
) -> List[ContactInput]:
    contacts: List[ContactInput] = []
    for name, email, phone in zip_longest(names, emails, phones, fillvalue=""):
        name = (name or "").strip()
        email = (email or "").strip()
        phone = (phone or "").strip()
        if not any((name, email, phone)):
            continue
        contact_name = name or email or phone
        contacts.append(
            ContactInput(
                name=contact_name,
                email=email or None,
                phone=phone or None,
            )
        )
    return contacts


@app.get("/", response_class=HTMLResponse)
def clients_page(
    request: Request,
    q: str | None = None,
    session: Session = Depends(get_session),
) -> HTMLResponse:
    clients = crud.list_clients(session, q=q)
    return templates.TemplateResponse(
        "clients_list.html",
        {"request": request, "clients": clients, "q": q or ""},
    )


@app.get("/_clients", response_class=HTMLResponse)
def clients_fragment(
    request: Request,
    q: str | None = None,
    session: Session = Depends(get_session),
) -> HTMLResponse:
    clients = crud.list_clients(session, q=q)
    return templates.TemplateResponse(
        "clients_list.html",
        {"request": request, "clients": clients, "q": q or ""},
    )


@app.post("/clients/new")
def create_client(
    company_name: str = Form(...),
    billing_address: str | None = Form(None),
    depannage: str = Form("non_refacturable"),
    astreinte: str = Form("pas_d_astreinte"),
    tag: str | None = Form(None),
    is_active: str = Form("true"),
    contact_name: List[str] = Form([]),
    contact_email: List[str] = Form([]),
    contact_phone: List[str] = Form([]),
    session: Session = Depends(get_session),
) -> RedirectResponse:
    contacts = _build_contacts(contact_name, contact_email, contact_phone)
    crud.create_client(
        session,
        ClientCreate(
            company_name=company_name,
            billing_address=billing_address,
            depannage=depannage,
            astreinte=astreinte,
            tag=tag,
            is_active=_as_bool(is_active),
            contacts=contacts,
        ),
    )
    return RedirectResponse(url="/", status_code=303)


@app.post("/clients/{client_id}/edit")
def edit_client(
    client_id: int,
    company_name: str = Form(...),
    billing_address: str | None = Form(None),
    depannage: str = Form("non_refacturable"),
    astreinte: str = Form("pas_d_astreinte"),
    tag: str | None = Form(None),
    is_active: str | None = Form(None),
    contact_name: List[str] = Form([]),
    contact_email: List[str] = Form([]),
    contact_phone: List[str] = Form([]),
    session: Session = Depends(get_session),
) -> RedirectResponse:
    contacts = _build_contacts(contact_name, contact_email, contact_phone)
    updated = crud.update_client(
        session,
        client_id,
        ClientUpdate(
            company_name=company_name,
            billing_address=billing_address,
            depannage=depannage,
            astreinte=astreinte,
            tag=tag,
            is_active=_as_bool(is_active) if is_active is not None else None,
            contacts=contacts,
        ),
    )
    if not updated:
        raise HTTPException(404, "Client introuvable")
    return RedirectResponse(url="/", status_code=303)


@app.post("/clients/{client_id}/delete")
def remove_client(client_id: int, session: Session = Depends(get_session)) -> RedirectResponse:
    ok = crud.delete_client(session, client_id)
    if not ok:
        raise HTTPException(404, "Client introuvable")
    return RedirectResponse(url="/", status_code=303)
