from typing import Optional

from fastapi import FastAPI, Depends, Request, Form, HTTPException, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlmodel import Session
from database import init_db, get_session
from models import ClientCreate, ClientUpdate
import crud
from importers import parse_clients_excel
from uuid import uuid4
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


def _clients_context(request: Request, clients, q: Optional[str]):
    report_token = request.query_params.get("report")
    report = None
    if report_token:
        reports = getattr(app.state, "import_reports", None)
        if reports is None:
            reports = {}
            app.state.import_reports = reports
        report = reports.pop(report_token, None)
    return {
        "request": request,
        "clients": clients,
        "q": q or "",
        "depannage_options": DEPANNAGE_OPTIONS,
        "astreinte_options": ASTREINTE_OPTIONS,
        "status_options": STATUS_OPTIONS,
        "import_report": report,
    }

@app.on_event("startup")
def on_startup():
    init_db()
    app.state.import_reports = {}

# Page liste
@app.get("/", response_class=HTMLResponse)
def clients_page(request: Request, q: Optional[str] = None, session: Session = Depends(get_session)):
    clients = crud.list_clients(session, q=q)
    return templates.TemplateResponse("clients_list.html", _clients_context(request, clients, q))

# Fragment liste (HTMX)
@app.get("/_clients", response_class=HTMLResponse)
def clients_fragment(request: Request, q: Optional[str] = None, session: Session = Depends(get_session)):
    clients = crud.list_clients(session, q=q)
    return templates.TemplateResponse("clients_list.html", _clients_context(request, clients, q))

# Form création
@app.post("/clients/new")
def create_client(
    company_name: str = Form(...),
    name: str = Form(...),
    email: Optional[str] = Form(None),
    phone: Optional[str] = Form(None),
    billing_address: Optional[str] = Form(None),
    depannage: str = Form("non_refacturable"),
    astreinte: str = Form("pas_d_astreinte"),
    tags: Optional[str] = Form(None),
    status: Optional[str] = Form("actif"),
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
    email: Optional[str] = Form(None),
    phone: Optional[str] = Form(None),
    billing_address: Optional[str] = Form(None),
    depannage: str = Form("non_refacturable"),
    astreinte: str = Form("pas_d_astreinte"),
    tags: Optional[str] = Form(None),
    status: Optional[str] = Form("actif"),
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


@app.post("/clients/import")
async def import_clients(
    file: UploadFile = File(...),
    session: Session = Depends(get_session),
):
    def store_report(created: int, total: int, errors: list[str], filename: str) -> RedirectResponse:
        report_id = uuid4().hex
        reports = getattr(app.state, "import_reports", None)
        if reports is None:
            reports = {}
            app.state.import_reports = reports
        reports[report_id] = {
            "created": created,
            "errors": errors,
            "total": total,
            "filename": filename,
        }
        return RedirectResponse(url=f"/?report={report_id}", status_code=303)

    if not file.filename:
        return store_report(0, 0, ["Aucun fichier sélectionné."], "Import")

    allowed_extensions = (".xlsx", ".xlsm", ".xltx", ".xltm")
    if not file.filename.lower().endswith(allowed_extensions):
        return store_report(0, 0, ["Format de fichier non supporté. Merci d'utiliser un fichier Excel (.xlsx)."], file.filename)

    content = await file.read()
    try:
        rows = parse_clients_excel(content)
    except ValueError as exc:
        return store_report(0, 0, [str(exc)], file.filename)

    created = 0
    errors: list[str] = []
    for idx, row in enumerate(rows, start=1):
        payload = {k: v for k, v in row.items() if not k.startswith("__")}
        row_number = row.get("__row__", idx)
        try:
            crud.create_client(session, ClientCreate(**payload))
            created += 1
        except Exception as exc:
            session.rollback()
            errors.append(f"Ligne {row_number} : {exc}")

    return store_report(created, len(rows), errors, file.filename)
