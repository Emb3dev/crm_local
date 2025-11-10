from typing import Optional, Dict, List, Iterable

from io import BytesIO

from fastapi import FastAPI, Depends, Request, Form, HTTPException, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlmodel import Session
from database import init_db, get_session
from models import (
    BeltLineCreate,
    ClientCreate,
    ClientUpdate,
    ContactCreate,
    FilterLineCreate,
    SubcontractedServiceCreate,
    SubcontractedServiceUpdate,
)
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

SUBCONTRACT_STATUS_DEFAULT = "non_commence"

SUBCONTRACT_STATUS_OPTIONS = {
    "non_commence": "Non commencé",
    "en_cours": "En cours",
    "fait": "Fait",
}

SUBCONTRACT_STATUS_STYLES = {
    "fait": "border-emerald-200 bg-emerald-50 text-emerald-700 dark:border-emerald-400/40 dark:bg-emerald-500/10 dark:text-emerald-100",
    "en_cours": "border-amber-200 bg-amber-50 text-amber-700 dark:border-amber-400/40 dark:bg-amber-500/10 dark:text-amber-100",
    "non_commence": "border-rose-200 bg-rose-50 text-rose-700 dark:border-rose-400/40 dark:bg-rose-500/10 dark:text-rose-100",
}

FILTER_FORMAT_OPTIONS = [
    ("cousus_sur_fil", "Cousus sur fil"),
    ("cadre", "Format cadre"),
    ("multi_diedre", "Format multi dièdre"),
    ("poche", "Format poche"),
]
FILTER_FORMAT_LABELS = {value: label for value, label in FILTER_FORMAT_OPTIONS}

CLIENT_FILTER_DEFINITIONS = [
    {
        "name": "status",
        "label": "Statut",
        "placeholder": "Tous les statuts",
        "options": list(STATUS_OPTIONS.items()),
    },
    {
        "name": "depannage",
        "label": "Dépannage",
        "placeholder": "Tous les dépannages",
        "options": list(DEPANNAGE_OPTIONS.items()),
    },
    {
        "name": "astreinte",
        "label": "Astreinte",
        "placeholder": "Toutes les astreintes",
        "options": [
            (key, label)
            for key, label in ASTREINTE_OPTIONS.items()
            if key != "pas_d_astreinte"
        ]
        + [("pas_d_astreinte", ASTREINTE_OPTIONS["pas_d_astreinte"])],
    },
]

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

FREQUENCY_UNITS = {
    "months": {
        "select_label": "mois",
        "singular_label": "mois",
        "plural_label": "mois",
        "use_plural_for_one": False,
    },
    "years": {
        "select_label": "années",
        "singular_label": "an",
        "plural_label": "ans",
        "use_plural_for_one": True,
    },
}

FREQUENCY_UNIT_OPTIONS = [(key, data["select_label"]) for key, data in FREQUENCY_UNITS.items()]

PREDEFINED_FREQUENCIES = {
    "contrat_annuel": {
        "label": "Contrat de maintenance annuel",
        "interval": 1,
        "unit": "years",
    },
    "contrat_semestriel": {
        "label": "Contrat de maintenance semestriel",
        "interval": 6,
        "unit": "months",
    },
    "contrat_trimestriel": {
        "label": "Contrat de maintenance trimestriel",
        "interval": 3,
        "unit": "months",
    },
    "contrat_bimestriel": {
        "label": "Contrat de maintenance bimestriel",
        "interval": 2,
        "unit": "months",
    },
    "contrat_mensuel": {
        "label": "Contrat de maintenance mensuel",
        "interval": 1,
        "unit": "months",
    },
    "prestation_ponctuelle": {
        "label": "Prestation ponctuelle",
        "interval": None,
        "unit": None,
    },
}

CUSTOM_INTERVAL_VALUE = "custom_interval"
CUSTOM_INTERVAL_LABEL = "Fréquence personnalisée"
INTERVAL_PREFIX = "interval:"

FREQUENCY_SELECT_OPTIONS = {
    key: data["label"] for key, data in PREDEFINED_FREQUENCIES.items()
}
FREQUENCY_SELECT_OPTIONS[CUSTOM_INTERVAL_VALUE] = CUSTOM_INTERVAL_LABEL

PREDEFINED_FREQUENCY_KEYS = tuple(PREDEFINED_FREQUENCIES.keys())

SUBCONTRACTING_FILTER_DEFINITIONS = [
    {
        "name": "category",
        "label": "Catégorie",
        "placeholder": "Toutes les catégories",
        "options": [(group["title"], group["title"]) for group in SUBCONTRACTED_GROUPS],
    },
    {
        "name": "frequency",
        "label": "Fréquence",
        "placeholder": "Toutes les fréquences",
        "options": [(key, data["label"]) for key, data in PREDEFINED_FREQUENCIES.items()],
    },
]


def _parse_interval_frequency(value: str) -> tuple[Optional[int], Optional[str]]:
    if not value.startswith(INTERVAL_PREFIX):
        return None, None
    parts = value.split(":")
    if len(parts) != 3:
        return None, None
    _, unit, interval_str = parts
    try:
        return int(interval_str), unit
    except ValueError:
        return None, None


def _format_interval_label(interval: int, unit: str) -> str:
    info = FREQUENCY_UNITS.get(unit)
    if not info:
        return f"Tous les {interval} {unit}"
    if interval == 1 and info.get("use_plural_for_one"):
        unit_label = info["plural_label"]
    elif interval == 1:
        unit_label = info["singular_label"]
    else:
        unit_label = info["plural_label"]
    if interval == 1:
        return f"Tous les {unit_label}"
    return f"Tous les {interval} {unit_label}"


def _frequency_label_from_details(
    frequency_value: str,
    interval: Optional[int],
    unit: Optional[str],
) -> str:
    predefined = PREDEFINED_FREQUENCIES.get(frequency_value)
    if predefined:
        return predefined["label"]
    resolved_interval = interval
    resolved_unit = unit
    if resolved_interval is None or not resolved_unit:
        parsed_interval, parsed_unit = _parse_interval_frequency(frequency_value)
        if resolved_interval is None:
            resolved_interval = parsed_interval
        if not resolved_unit:
            resolved_unit = parsed_unit
    if resolved_interval and resolved_unit in FREQUENCY_UNITS:
        return _format_interval_label(resolved_interval, resolved_unit)
    return frequency_value


def _resolve_frequency(
    selection: str,
    custom_interval: Optional[str],
    custom_unit: Optional[str],
) -> tuple[str, Optional[int], Optional[str]]:
    if selection == CUSTOM_INTERVAL_VALUE:
        if not custom_interval:
            raise HTTPException(400, "Veuillez renseigner l'intervalle de la fréquence")
        if custom_unit not in FREQUENCY_UNITS:
            raise HTTPException(400, "Unité de fréquence inconnue")
        try:
            interval_value = int(custom_interval)
        except ValueError as exc:
            raise HTTPException(400, "L'intervalle doit être un nombre entier") from exc
        if interval_value < 1:
            raise HTTPException(400, "L'intervalle doit être supérieur ou égal à 1")
        if interval_value > 120:
            raise HTTPException(400, "L'intervalle ne peut pas dépasser 120")
        frequency_value = f"{INTERVAL_PREFIX}{custom_unit}:{interval_value}"
        return frequency_value, interval_value, custom_unit
    if selection not in PREDEFINED_FREQUENCIES:
        raise HTTPException(400, "Fréquence inconnue")
    predefined = PREDEFINED_FREQUENCIES[selection]
    return selection, predefined["interval"], predefined["unit"]


def _build_frequency_labels(
    services: Iterable = (), extra_values: Iterable[str] = ()
) -> Dict[str, str]:
    labels = {key: data["label"] for key, data in PREDEFINED_FREQUENCIES.items()}
    for service in services:
        value = getattr(service, "frequency", None)
        if not value:
            continue
        interval = getattr(service, "frequency_interval", None)
        unit = getattr(service, "frequency_unit", None)
        labels[value] = _frequency_label_from_details(value, interval, unit)
    for value in extra_values:
        if not value or value == CUSTOM_INTERVAL_VALUE:
            continue
        if value not in labels:
            labels[value] = _frequency_label_from_details(value, None, None)
    return labels


def _build_frequency_filter_options(
    services: Iterable = (), extra_values: Iterable[str] = ()
) -> List[tuple[str, str]]:
    labels = _build_frequency_labels(services, extra_values)
    options: List[tuple[str, str]] = [
        (key, labels[key]) for key in PREDEFINED_FREQUENCY_KEYS if key in labels
    ]
    seen = {key for key, _ in options}
    for value in sorted(labels.keys()):
        if value in seen or value == CUSTOM_INTERVAL_VALUE:
            continue
        options.append((value, labels[value]))
        seen.add(value)
    return options


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


def _clients_context(request: Request, clients, q: Optional[str], filters: Dict[str, str]):
    report_token = request.query_params.get("report")
    report = None
    if report_token:
        reports = getattr(app.state, "import_reports", None)
        if reports is None:
            reports = {}
            app.state.import_reports = reports
        report = reports.pop(report_token, None)
    all_services = [
        service
        for client in clients
        for service in getattr(client, "subcontractings", [])
    ]
    frequency_labels = _build_frequency_labels(all_services)
    return {
        "request": request,
        "clients": clients,
        "q": q or "",
        "depannage_options": DEPANNAGE_OPTIONS,
        "astreinte_options": ASTREINTE_OPTIONS,
        "status_options": STATUS_OPTIONS,
        "subcontract_status_options": SUBCONTRACT_STATUS_OPTIONS,
        "subcontract_status_styles": SUBCONTRACT_STATUS_STYLES,
        "subcontract_status_default": SUBCONTRACT_STATUS_DEFAULT,
        "subcontracted_groups": SUBCONTRACTED_GROUPS,
        "frequency_options": frequency_labels,
        "frequency_select_options": FREQUENCY_SELECT_OPTIONS,
        "frequency_unit_options": FREQUENCY_UNIT_OPTIONS,
        "custom_interval_value": CUSTOM_INTERVAL_VALUE,
        "predefined_frequency_keys": PREDEFINED_FREQUENCY_KEYS,
        "import_report": report,
        "focus_id": request.query_params.get("focus"),
        "active_filters": filters,
        "filters_definition": CLIENT_FILTER_DEFINITIONS,
    }

def _subcontractings_context(
    request: Request, services, q: Optional[str], filters: Dict[str, str]
):
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

    active_frequency = filters.get("frequency")
    frequency_labels = _build_frequency_labels(services, [active_frequency] if active_frequency else [])
    frequency_filter_options = _build_frequency_filter_options(services, [active_frequency] if active_frequency else [])
    filters_definition = [
        {**SUBCONTRACTING_FILTER_DEFINITIONS[0]},
        {**SUBCONTRACTING_FILTER_DEFINITIONS[1], "options": frequency_filter_options},
    ]

    return {
        "request": request,
        "services": services,
        "q": q or "",
        "frequency_options": frequency_labels,
        "frequency_select_options": FREQUENCY_SELECT_OPTIONS,
        "frequency_unit_options": FREQUENCY_UNIT_OPTIONS,
        "custom_interval_value": CUSTOM_INTERVAL_VALUE,
        "predefined_frequency_keys": PREDEFINED_FREQUENCY_KEYS,
        "category_totals": category_totals,
        "frequency_totals": frequency_totals,
        "total_budget": total_budget,
        "total_services": len(services),
        "distinct_clients": len(client_ids),
        "active_filters": filters,
        "filters_definition": filters_definition,
        "focus_id": request.query_params.get("focus"),
        "subcontract_status_options": SUBCONTRACT_STATUS_OPTIONS,
        "subcontract_status_styles": SUBCONTRACT_STATUS_STYLES,
        "subcontract_status_default": SUBCONTRACT_STATUS_DEFAULT,
    }


def _extract_client_filters(
    status: Optional[str], depannage: Optional[str], astreinte: Optional[str]
) -> Dict[str, str]:
    filters: Dict[str, str] = {}
    if status in STATUS_OPTIONS:
        filters["status"] = status
    if depannage in DEPANNAGE_OPTIONS:
        filters["depannage"] = depannage
    if astreinte in ASTREINTE_OPTIONS:
        filters["astreinte"] = astreinte
    return filters


def _extract_subcontracting_filters(
    category: Optional[str], frequency: Optional[str]
) -> Dict[str, str]:
    filters: Dict[str, str] = {}
    valid_categories = {group["title"] for group in SUBCONTRACTED_GROUPS}
    if category in valid_categories:
        filters["category"] = category
    if frequency and frequency != CUSTOM_INTERVAL_VALUE:
        filters["frequency"] = frequency
    return filters


@app.on_event("startup")
def on_startup():
    init_db()
    app.state.import_reports = {}

# Page liste
@app.get("/", response_class=HTMLResponse)
def clients_page(
    request: Request,
    q: Optional[str] = None,
    status: Optional[str] = None,
    depannage: Optional[str] = None,
    astreinte: Optional[str] = None,
    session: Session = Depends(get_session),
):
    filters = _extract_client_filters(status, depannage, astreinte)
    clients = crud.list_clients(session, q=q, filters=filters)
    return templates.TemplateResponse(
        "clients_list.html", _clients_context(request, clients, q, filters)
    )

# Fragment liste (HTMX)
@app.get("/_clients", response_class=HTMLResponse)
def clients_fragment(
    request: Request,
    q: Optional[str] = None,
    status: Optional[str] = None,
    depannage: Optional[str] = None,
    astreinte: Optional[str] = None,
    session: Session = Depends(get_session),
):
    filters = _extract_client_filters(status, depannage, astreinte)
    clients = crud.list_clients(session, q=q, filters=filters)
    return templates.TemplateResponse(
        "clients_list.html", _clients_context(request, clients, q, filters)
    )


@app.get("/prestations", response_class=HTMLResponse)
def subcontracted_services_page(
    request: Request,
    q: Optional[str] = None,
    category: Optional[str] = None,
    frequency: Optional[str] = None,
    session: Session = Depends(get_session),
):
    filters = _extract_subcontracting_filters(category, frequency)
    services = crud.list_subcontracted_services(session, q=q, filters=filters)
    return templates.TemplateResponse(
        "subcontractings_list.html",
        _subcontractings_context(request, services, q, filters),
    )


@app.get("/prestations/{service_id}/edit", response_class=HTMLResponse)
def subcontracted_service_edit_page(
    request: Request,
    service_id: int,
    session: Session = Depends(get_session),
):
    service = crud.get_subcontracted_service(session, service_id)
    if not service:
        raise HTTPException(404, "Prestation introuvable")

    available_keys = set(SUBCONTRACTED_LOOKUP.keys())
    clients = crud.list_client_choices(session)
    return templates.TemplateResponse(
        "subcontracting_edit.html",
        {
            "request": request,
            "service": service,
            "subcontracted_groups": SUBCONTRACTED_GROUPS,
            "frequency_options": _build_frequency_labels([service]),
            "frequency_select_options": FREQUENCY_SELECT_OPTIONS,
            "frequency_unit_options": FREQUENCY_UNIT_OPTIONS,
            "custom_interval_value": CUSTOM_INTERVAL_VALUE,
            "predefined_frequency_keys": PREDEFINED_FREQUENCY_KEYS,
            "subcontract_status_options": SUBCONTRACT_STATUS_OPTIONS,
            "subcontract_status_default": SUBCONTRACT_STATUS_DEFAULT,
            "available_prestations": available_keys,
            "clients": clients,
            "return_url": f"/prestations?focus={service.id}#service-{service.id}",
        },
    )


@app.post("/prestations/{service_id}/edit")
def update_subcontracted_service(
    service_id: int,
    prestation: str = Form(...),
    client_id: int = Form(...),
    budget: Optional[str] = Form(None),
    frequency: str = Form(...),
    custom_frequency_interval: Optional[str] = Form(None),
    custom_frequency_unit: Optional[str] = Form(None),
    status: str = Form(SUBCONTRACT_STATUS_DEFAULT),
    realization_week: Optional[str] = Form(None),
    order_week: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    service = crud.get_subcontracted_service(session, service_id)
    if not service:
        raise HTTPException(404, "Prestation introuvable")

    if not crud.get_client(session, client_id):
        raise HTTPException(400, "Client inconnu")

    (
        resolved_frequency,
        resolved_interval,
        resolved_unit,
    ) = _resolve_frequency(frequency, custom_frequency_interval, custom_frequency_unit)

    details = SUBCONTRACTED_LOOKUP.get(prestation)
    if not details and prestation != service.prestation_key:
        raise HTTPException(400, "Prestation inconnue")

    if status not in SUBCONTRACT_STATUS_OPTIONS:
        raise HTTPException(400, "Statut inconnu")

    parsed_budget = _parse_budget(budget)

    realization_value: Optional[str] = None
    if resolved_frequency == "prestation_ponctuelle" and realization_week:
        realization_value = realization_week.strip().upper()

    order_value = order_week.strip().upper() if order_week else None

    update_payload = SubcontractedServiceUpdate(
        prestation_key=prestation,
        prestation_label=(
            details["label"] if details else service.prestation_label
        ),
        category=(details["category"] if details else service.category),
        budget_code=(details["budget_code"] if details else service.budget_code),
        budget=parsed_budget,
        frequency=resolved_frequency,
        frequency_interval=resolved_interval,
        frequency_unit=resolved_unit,
        status=status,
        realization_week=realization_value,
        order_week=order_value,
        client_id=client_id,
    )

    updated = crud.update_subcontracted_service(session, service_id, update_payload)
    if not updated:
        raise HTTPException(404, "Prestation introuvable")

    return RedirectResponse(
        url=f"/prestations?focus={service_id}#service-{service_id}",
        status_code=303,
    )

# Form création
@app.post("/clients/new")
def create_client(
    company_name: str = Form(...),
    name: str = Form(...),
    email: Optional[str] = Form(None),
    phone: Optional[str] = Form(None),
    technician_name: Optional[str] = Form(None),
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
            technician_name=technician_name,
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
    technician_name: Optional[str] = Form(None),
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
            technician_name=technician_name,
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
    custom_frequency_interval: Optional[str] = Form(None),
    custom_frequency_unit: Optional[str] = Form(None),
    status: str = Form(SUBCONTRACT_STATUS_DEFAULT),
    realization_week: Optional[str] = Form(None),
    order_week: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    details = SUBCONTRACTED_LOOKUP.get(prestation)
    if not details:
        raise HTTPException(400, "Prestation inconnue")
    (
        resolved_frequency,
        resolved_interval,
        resolved_unit,
    ) = _resolve_frequency(frequency, custom_frequency_interval, custom_frequency_unit)

    if status not in SUBCONTRACT_STATUS_OPTIONS:
        raise HTTPException(400, "Statut inconnu")

    parsed_budget = _parse_budget(budget)
    realization_value = (
        realization_week.strip().upper()
        if realization_week and resolved_frequency == "prestation_ponctuelle"
        else None
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
            frequency=resolved_frequency,
            frequency_interval=resolved_interval,
            frequency_unit=resolved_unit,
            status=status,
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


@app.get("/filtres-courroies", response_class=HTMLResponse)
def list_filters_and_belts(
    request: Request,
    session: Session = Depends(get_session),
):
    filters = crud.list_filter_lines(session)
    belts = crud.list_belt_lines(session)
    return templates.TemplateResponse(
        "filters_belts.html",
        {
            "request": request,
            "filters": filters,
            "belts": belts,
            "filter_format_options": FILTER_FORMAT_OPTIONS,
            "filter_format_labels": FILTER_FORMAT_LABELS,
        },
    )


@app.post("/filtres-courroies/filtres")
async def create_filter_line(
    site: str = Form(...),
    equipment: str = Form(...),
    filter_type: str = Form(...),
    efficiency: Optional[str] = Form(None),
    format_type: str = Form(...),
    dimensions: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    if format_type not in FILTER_FORMAT_LABELS:
        raise HTTPException(status_code=400, detail="Format de filtre invalide")

    payload = FilterLineCreate(
        site=site.strip(),
        equipment=equipment.strip(),
        filter_type=filter_type.strip(),
        efficiency=(efficiency.strip() if efficiency else None),
        format_type=format_type,
        dimensions=(dimensions.strip() if dimensions else None),
    )
    crud.create_filter_line(session, payload)
    return RedirectResponse("/filtres-courroies", status_code=303)


@app.post("/filtres-courroies/filtres/{line_id}/delete")
def delete_filter_line(line_id: int, session: Session = Depends(get_session)):
    if not crud.delete_filter_line(session, line_id):
        raise HTTPException(status_code=404, detail="Ligne filtre introuvable")
    return RedirectResponse("/filtres-courroies", status_code=303)


@app.post("/filtres-courroies/courroies")
async def create_belt_line(
    site: str = Form(...),
    equipment: str = Form(...),
    reference: str = Form(...),
    quantity: int = Form(...),
    order_week: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    payload = BeltLineCreate(
        site=site.strip(),
        equipment=equipment.strip(),
        reference=reference.strip(),
        quantity=quantity,
        order_week=(order_week.strip().upper() if order_week else None),
    )
    crud.create_belt_line(session, payload)
    return RedirectResponse("/filtres-courroies", status_code=303)


@app.post("/filtres-courroies/courroies/{line_id}/delete")
def delete_belt_line(line_id: int, session: Session = Depends(get_session)):
    if not crud.delete_belt_line(session, line_id):
        raise HTTPException(status_code=404, detail="Ligne courroie introuvable")
    return RedirectResponse("/filtres-courroies", status_code=303)


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
        "technician_name",
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
            "technician_name": "Léo Fournier",
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
            "technician_name": "Nina Perrot",
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
            "technician_name": "",
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
        "technician_name",
        "Nom complet",
        "Identifie le technicien référent pour ce client.",
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
