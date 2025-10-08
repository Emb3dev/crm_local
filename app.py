from fastapi import FastAPI, Depends, Request, Form, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlmodel import Session
from database import init_db, get_session
from models import ClientCreate, ClientUpdate
import crud
#uvicorn app:app --reload

app = FastAPI(title="CRM Local")
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

DEPANNAGE_OPTIONS = {
    "refacturable": "Refacturable",
    "non_refacturable": "Non refacturable",
}

ASTREINTE_OPTIONS = {
    "incluse_non_refacturable": "Incluse mais non refacturable",
    "incluse_refacturable": "Incluse mais refacturable",
    "pas_d_astreinte": "Pas d'astreinte",
}

STATUS_OPTIONS = {
    "actif": "Actif",
    "inactif": "Inactif",
}


def _clients_context(request: Request, clients, q: str | None):
    return {
        "request": request,
        "clients": clients,
        "q": q or "",
        "depannage_options": DEPANNAGE_OPTIONS,
        "astreinte_options": ASTREINTE_OPTIONS,
        "status_options": STATUS_OPTIONS,
    }

@app.on_event("startup")
def on_startup():
    init_db()

# Page liste
@app.get("/", response_class=HTMLResponse)
def clients_page(request: Request, q: str | None = None, session: Session = Depends(get_session)):
    clients = crud.list_clients(session, q=q)
    return templates.TemplateResponse("clients_list.html", _clients_context(request, clients, q))

# Fragment liste (HTMX)
@app.get("/_clients", response_class=HTMLResponse)
def clients_fragment(request: Request, q: str | None = None, session: Session = Depends(get_session)):
    clients = crud.list_clients(session, q=q)
    return templates.TemplateResponse("clients_list.html", _clients_context(request, clients, q))

# Form création
@app.post("/clients/new")
def create_client(
    company_name: str = Form(...),
    name: str = Form(...),
    email: str | None = Form(None),
    phone: str | None = Form(None),
    billing_address: str | None = Form(None),
    depannage: str = Form("non_refacturable"),
    astreinte: str = Form("pas_d_astreinte"),
    tags: str | None = Form(None),
    status: str | None = Form("actif"),
    session: Session = Depends(get_session)
):
    crud.create_client(
        session,
        ClientCreate(
            company_name=company_name,
            name=name,
            email=email,
            phone=phone,
            billing_address=billing_address,
            depannage=depannage,
            astreinte=astreinte,
            tags=tags,
            status=status,
        ),
    )
    return RedirectResponse(url="/", status_code=303)

# Maj
@app.post("/clients/{client_id}/edit")
def edit_client(
    client_id: int,
    company_name: str = Form(...),
    name: str = Form(...),
    email: str | None = Form(None),
    phone: str | None = Form(None),
    billing_address: str | None = Form(None),
    depannage: str = Form("non_refacturable"),
    astreinte: str = Form("pas_d_astreinte"),
    tags: str | None = Form(None),
    status: str | None = Form("actif"),
    session: Session = Depends(get_session)
):
    updated = crud.update_client(
        session,
        client_id,
        ClientUpdate(
            company_name=company_name,
            name=name,
            email=email,
            phone=phone,
            billing_address=billing_address,
            depannage=depannage,
            astreinte=astreinte,
            tags=tags,
            status=status,
        ),
    )
    if not updated: raise HTTPException(404, "Client introuvable")
    return RedirectResponse(url="/", status_code=303)

@app.post("/clients/{client_id}/delete")
def remove_client(client_id: int, session: Session = Depends(get_session)):
    ok = crud.delete_client(session, client_id)
    if not ok: raise HTTPException(404, "Client introuvable")
    return RedirectResponse(url="/", status_code=303)
