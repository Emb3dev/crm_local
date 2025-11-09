from typing import Optional, Dict, List

from io import BytesIO

from fastapi import FastAPI, Depends, Request, Form, HTTPException, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlmodel import Session
from database import init_db, get_session
from models import ClientCreate, ClientUpdate, ContactCreate, SubcontractedServiceCreate
import crud
from importers import parse_clients_excel
from openpyxl import Workbook
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

SUBCONTRACTED_GROUPS = [
    {
        "key": "sous_traitance",
        "title": "Sous-traitance",
        "options": [
            {"value": "analyse_eau", "label": "Analyse d'eau", "budget_code": "S1000"},
            {"value": "analyse_huile", "label": "Analyse d'huile", "budget_code": "S1010"},
            {
                "value": "analyse_eau_nappe",
                "label": "Analyse eau de nappe",
                "budget_code": "S1020",
            },
            {
                "value": "analyse_legionnelle",
                "label": "Analyse légionnelle",
                "budget_code": "S1030",
            },
            {
                "value": "analyse_potabilite",
                "label": "Analyse potabilité",
                "budget_code": "S1040",
            },
            {"value": "colonnes_seches", "label": "Colonnes sèches", "budget_code": "S1050"},
            {"value": "controle_acces", "label": "Contrôle d'accès", "budget_code": "S1060"},
            {"value": "controle_ssi", "label": "Contrôle SSI", "budget_code": "S1070"},
            {"value": "detection_co", "label": "Détection CO", "budget_code": "S1080"},
            {"value": "detection_freon", "label": "Détection fréon", "budget_code": "S1090"},
            {"value": "detection_incendie", "label": "Détection incendie", "budget_code": "S2000"},
            {"value": "extincteurs", "label": "Extincteurs", "budget_code": "S2010"},
            {"value": "exutoires", "label": "Exutoires", "budget_code": "S2020"},
            {"value": "gtc", "label": "GTC", "budget_code": "S2030"},
            {
                "value": "inspection_video_puits",
                "label": "Inspection vidéo puits",
                "budget_code": "S2040",
            },
            {
                "value": "maintenance_cellule_hta",
                "label": "Maintenance cellule HTA",
                "budget_code": "S2050",
            },
            {
                "value": "maintenance_constructeur",
                "label": "Maintenance constructeur",
                "budget_code": "S2060",
            },
            {
                "value": "maintenance_groupe_electrogene",
                "label": "Maintenance groupe électrogène",
                "budget_code": "S2070",
            },
            {
                "value": "maintenance_groupe_froid",
                "label": "Maintenance groupe froid",
                "budget_code": "S2080",
            },
            {
                "value": "nettoyage_gaines",
                "label": "Nettoyage de gaines",
                "budget_code": "S3000",
            },
            {"value": "onduleurs", "label": "Onduleurs", "budget_code": "S3010"},
            {
                "value": "pompe_relevage",
                "label": "Pompe de relevage",
                "budget_code": "S3020",
            },
            {
                "value": "portes_automatiques",
                "label": "Portes automatiques",
                "budget_code": "S3030",
            },
            {
                "value": "portes_coupe_feu",
                "label": "Portes coupe-feu",
                "budget_code": "S3040",
            },
            {"value": "ramonage", "label": "Ramonage", "budget_code": "S3050"},
            {"value": "relamping", "label": "Relamping", "budget_code": "S3060"},
            {
                "value": "separateur_hydrocarbures",
                "label": "Séparateur hydrocarbures",
                "budget_code": "S3070",
            },
            {"value": "sorbonnes", "label": "Sorbonnes", "budget_code": "S4000"},
            {
                "value": "table_elevatrice",
                "label": "Table élévatrice",
                "budget_code": "S4010",
            },
            {
                "value": "telesurveillance",
                "label": "Télésurveillance",
                "budget_code": "S4020",
            },
            {"value": "thermographie", "label": "Thermographie", "budget_code": "S4030"},
            {"value": "traitement_eau", "label": "Traitement d'eau", "budget_code": "S4040"},
            {
                "value": "video_interphonie",
                "label": "Vidéo et interphonie",
                "budget_code": "S4050",
            },
        ],
    },
    {
        "key": "locations",
        "title": "Locations",
        "options": [
            {
                "value": "location_echafaudage",
                "label": "Location échafaudage",
                "budget_code": "L1010",
            },
            {
                "value": "location_groupe_electrogene",
                "label": "Location groupe électrogène",
                "budget_code": "L1020",
            },
            {
                "value": "location_nacelle",
                "label": "Location nacelle",
                "budget_code": "L1030",
            },
        ],
    },
]

SUBCONTRACTED_LOOKUP = {
    option["value"]: {
        "label": option["label"],
        "budget_code": option["budget_code"],
        "category": group["title"],
    }
    for group in SUBCONTRACTED_GROUPS
    for option in group["options"]
}

FREQUENCY_OPTIONS = {
    "contrat_annuel": "Contrat de maintenance annuel",
    "prestation_ponctuelle": "Prestation ponctuelle",
}


def _parse_budget(value: Optional[str]) -> Optional[float]:
    if value is None:
        return None
    normalized = value.strip().replace(" ", "").replace(",", ".")
    if not normalized:
        return None
    try:
        return float(normalized)
    except ValueError as exc:
        raise HTTPException(400, "Budget invalide") from exc


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
        "subcontracted_groups": SUBCONTRACTED_GROUPS,
        "frequency_options": FREQUENCY_OPTIONS,
        "import_report": report,
        "focus_id": request.query_params.get("focus"),
    }

def _subcontractings_context(request: Request, services, q: Optional[str]):
    category_totals: Dict[str, int] = {}
    frequency_totals: Dict[str, int] = {}
    total_budget = 0.0
    client_ids = set()

    for service in services:
        category_totals[service.category] = category_totals.get(service.category, 0) + 1
        frequency_totals[service.frequency] = frequency_totals.get(service.frequency, 0) + 1
        if service.budget is not None:
            total_budget += float(service.budget)
        if service.client_id:
            client_ids.add(service.client_id)

    return {
        "request": request,
        "services": services,
        "q": q or "",
        "frequency_options": FREQUENCY_OPTIONS,
        "category_totals": category_totals,
        "frequency_totals": frequency_totals,
        "total_budget": total_budget,
        "total_services": len(services),
        "distinct_clients": len(client_ids),
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


@app.get("/prestations", response_class=HTMLResponse)
def subcontracted_services_page(
    request: Request,
    q: Optional[str] = None,
    session: Session = Depends(get_session),
):
    services = crud.list_subcontracted_services(session, q=q)
    return templates.TemplateResponse(
        "subcontractings_list.html", _subcontractings_context(request, services, q)
    )

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
    contacts_payload: List[ContactCreate] = []
    if name or email or phone:
        contacts_payload.append(ContactCreate(name=name, email=email, phone=phone))

    client = crud.create_client(
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
        contacts=contacts_payload,
    )
    return RedirectResponse(
        url=f"/?focus={client.id}#client-{client.id}", status_code=303
    )

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


@app.post("/clients/{client_id}/contacts")
def add_contact(
    client_id: int,
    name: str = Form(..., alias="contact_name"),
    email: Optional[str] = Form(None, alias="contact_email"),
    phone: Optional[str] = Form(None, alias="contact_phone"),
    session: Session = Depends(get_session),
):
    created = crud.create_contact(
        session,
        client_id,
        ContactCreate(name=name, email=email, phone=phone),
    )
    if not created:
        raise HTTPException(404, "Client introuvable")
    return RedirectResponse(
        url=f"/?focus={client_id}#client-{client_id}", status_code=303
    )


@app.post("/clients/{client_id}/contacts/{contact_id}/delete")
def remove_contact(
    client_id: int,
    contact_id: int,
    session: Session = Depends(get_session),
):
    ok = crud.delete_contact(session, client_id, contact_id)
    if not ok:
        raise HTTPException(404, "Contact introuvable")
    return RedirectResponse(
        url=f"/?focus={client_id}#client-{client_id}", status_code=303
    )


@app.post("/clients/{client_id}/subcontractings")
def add_subcontracted_service(
    client_id: int,
    prestation: str = Form(...),
    budget: Optional[str] = Form(None),
    frequency: str = Form(...),
    realization_week: Optional[str] = Form(None),
    order_week: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    details = SUBCONTRACTED_LOOKUP.get(prestation)
    if not details:
        raise HTTPException(400, "Prestation inconnue")
    if frequency not in FREQUENCY_OPTIONS:
        raise HTTPException(400, "Fréquence inconnue")

    parsed_budget = _parse_budget(budget)
    realization_value = (
        realization_week.strip().upper() if realization_week and frequency == "prestation_ponctuelle" else None
    )
    order_value = order_week.strip().upper() if order_week else None

    created = crud.create_subcontracted_service(
        session,
        client_id,
        SubcontractedServiceCreate(
            prestation_key=prestation,
            prestation_label=details["label"],
            category=details["category"],
            budget_code=details["budget_code"],
            budget=parsed_budget,
            frequency=frequency,
            realization_week=realization_value,
            order_week=order_value,
        ),
    )
    if not created:
        raise HTTPException(404, "Client introuvable")
    return RedirectResponse(
        url=f"/?focus={client_id}#client-{client_id}", status_code=303
    )


@app.post("/clients/{client_id}/subcontractings/{service_id}/delete")
def remove_subcontracted_service(
    client_id: int,
    service_id: int,
    session: Session = Depends(get_session),
):
    ok = crud.delete_subcontracted_service(session, client_id, service_id)
    if not ok:
        raise HTTPException(404, "Prestation introuvable")
    return RedirectResponse(
        url=f"/?focus={client_id}#client-{client_id}", status_code=303
    )


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
        contacts_raw = row.get("contacts", [])
        contacts_payload = [ContactCreate(**contact) for contact in contacts_raw]
        payload = {
            k: v
            for k, v in row.items()
            if not k.startswith("__") and k != "contacts"
        }
        row_number = row.get("__row__", idx)
        try:
            crud.create_client(session, ClientCreate(**payload), contacts=contacts_payload)
            created += 1
        except Exception as exc:
            session.rollback()
            errors.append(f"Ligne {row_number} : {exc}")

    return store_report(created, len(rows), errors, file.filename)


def _build_client_import_template() -> BytesIO:
    workbook = Workbook()

    sheet = workbook.active
    sheet.title = "Clients"

    headers = [
        "company_name",
        "name",
        "email",
        "phone",
        "billing_address",
        "depannage",
        "astreinte",
        "tags",
        "status",
        "contact_2_name",
        "contact_2_email",
        "contact_2_phone",
        "contact_3_name",
        "contact_3_email",
        "contact_3_phone",
    ]
    sheet.append(headers)

    sample_rows = [
        {
            "company_name": "Exemple SARL",
            "name": "Alice Martin",
            "email": "alice@example.fr",
            "phone": "0102030405",
            "billing_address": "10 rue des Fleurs\n75000 Paris",
            "depannage": "refacturable",
            "astreinte": "incluse_refacturable",
            "tags": "premium, 2024",
            "status": "actif",
            "contact_2_name": "Paul Martin",
            "contact_2_email": "paul@example.fr",
            "contact_2_phone": "0188776655",
        },
        {
            "company_name": "Solutions BTP",
            "name": "Bruno Carrel",
            "email": "bruno@solutionsbtp.fr",
            "phone": "0611223344",
            "billing_address": "5 avenue du Port\n44000 Nantes",
            "depannage": "non_refacturable",
            "astreinte": "incluse_non_refacturable",
            "tags": "chantier",
            "status": "actif",
            "contact_2_name": "Sophie Lemaitre",
            "contact_2_email": "sophie@solutionsbtp.fr",
            "contact_3_name": "Service Achat",
            "contact_3_phone": "0245789650",
        },
        {
            "company_name": "Collectif Horizon",
            "name": "Chloé Bernard",
            "email": "chloe@collectif-horizon.fr",
            "phone": "0499887766",
            "billing_address": "42 boulevard National\n13001 Marseille",
            "depannage": "refacturable",
            "astreinte": "pas_d_astreinte",
            "tags": "association",
            "status": "inactif",
        },
    ]

    for row in sample_rows:
        sheet.append([row.get(column, "") for column in headers])

    sheet.freeze_panes = "A2"

    for column_cells in sheet.columns:
        max_length = 0
        column = column_cells[0].column_letter
        for cell in column_cells:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        sheet.column_dimensions[column].width = min(max_length + 2, 50)

    options_sheet = workbook.create_sheet("Options")
    options_sheet.append(["Champ", "Valeurs autorisées", "Description"])
    options_sheet.freeze_panes = "A2"

    def _format_options(options: Dict[str, str]) -> str:
        return "\n".join(f"{key} — {label}" for key, label in options.items())

    options_sheet.append([
        "depannage",
        _format_options(DEPANNAGE_OPTIONS),
        "Détermine si les interventions sont refacturées.",
    ])
    options_sheet.append([
        "astreinte",
        _format_options(ASTREINTE_OPTIONS),
        "Choisissez le type d'astreinte applicable.",
    ])
    options_sheet.append([
        "status",
        _format_options(STATUS_OPTIONS),
        "Etat du client dans votre CRM.",
    ])
    options_sheet.append([
        "contacts supplémentaires",
        "contact_2_name, contact_2_email, contact_2_phone, contact_3_name…",
        "Dupliquez le numéro (contact_4_*, contact_5_* …) pour ajouter des contacts additionnels. Seul le nom est obligatoire.",
    ])

    for column_cells in options_sheet.columns:
        max_length = 0
        column = column_cells[0].column_letter
        for cell in column_cells:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        options_sheet.column_dimensions[column].width = min(max_length + 4, 70)

    buffer = BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer


def _template_response() -> Response:
    buffer = _build_client_import_template()
    content = buffer.getvalue()
    headers_dict = {
        "Content-Disposition": (
            "attachment; filename=modele_import_clients.xlsx; "
            "filename*=UTF-8''modele_import_clients.xlsx"
        )
    }
    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers_dict,
    )


@app.get("/clients/import/template")
@app.get("/clients/import/template/")
def download_import_template():
    return _template_response()
