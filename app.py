from typing import Optional, Dict, List, Iterable, Union

from io import BytesIO

from fastapi import FastAPI, Depends, Request, Form, HTTPException, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlmodel import Session
from database import init_db, get_session
from models import (
    BeltLineCreate,
    BeltLineUpdate,
    ClientCreate,
    ClientUpdate,
    ContactCreate,
    FilterLineCreate,
    FilterLineUpdate,
    SubcontractedServiceCreate,
    SubcontractedServiceUpdate,
    WorkloadCellUpdate,
    WorkloadSiteCreate,
    WorkloadSiteUpdate,
)
import crud
from importers import (
    parse_belt_lines_excel,
    parse_clients_excel,
    parse_filter_lines_excel,
    parse_workload_plan_excel,
)
from openpyxl import Workbook
from uuid import uuid4
from pydantic import BaseModel, Field
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

ALLOWED_IMPORT_EXTENSIONS = (".xlsx", ".xlsm", ".xltx", ".xltm")


def _status_to_bool(value: Optional[str]) -> Optional[bool]:
    if value is None:
        return None
    if value not in STATUS_OPTIONS:
        raise HTTPException(400, "Statut inconnu")
    return value == "actif"


def _status_key_from_bool(value: Optional[bool]) -> str:
    if value is None:
        return ""
    return "actif" if value else "inactif"


def _consume_import_report(request: Request):
    report_token = request.query_params.get("report")
    if not report_token:
        return None
    reports = getattr(app.state, "import_reports", None)
    if reports is None:
        reports = {}
        app.state.import_reports = reports
    return reports.pop(report_token, None)


def _store_import_report(
    redirect_url: str,
    *,
    created: int,
    total: int,
    errors: List[str],
    filename: str,
    singular_label: str,
    plural_label: str,
) -> RedirectResponse:
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
        "entity_label": singular_label,
        "entity_label_plural": plural_label,
    }
    return RedirectResponse(url=f"{redirect_url}?report={report_id}", status_code=303)

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


def _clients_context(
    request: Request,
    clients,
    q: Optional[str],
    filters: Dict[str, str],
    entreprises,
):
    report = _consume_import_report(request)
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
        "entreprises": entreprises,
        "entreprise_name_options": sorted({e.nom for e in entreprises}),
        "status_key_from_bool": _status_key_from_bool,
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
    entreprises = crud.list_entreprises(session)
    return templates.TemplateResponse(
        "clients_list.html",
        _clients_context(request, clients, q, filters, entreprises),
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
    entreprises = crud.list_entreprises(session)
    return templates.TemplateResponse(
        "clients_list.html",
        _clients_context(request, clients, q, filters, entreprises),
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

    statut_bool = _status_to_bool(status)
    entreprise = crud.ensure_entreprise(
        session,
        name=company_name,
        adresse_facturation=billing_address,
        tag=tags,
        statut=statut_bool,
    )
    client = crud.create_client(
        session,
        ClientCreate(
            entreprise_id=entreprise.id,
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
    statut_bool = _status_to_bool(status)
    entreprise = crud.ensure_entreprise(
        session,
        name=company_name,
        adresse_facturation=billing_address,
        tag=tags,
        statut=statut_bool,
    )
    updated = crud.update_client(
        session,
        client_id,
        ClientUpdate(
            entreprise_id=entreprise.id,
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
    editing_filter = None
    editing_belt = None

    edit_filter_id = request.query_params.get("edit_filter")
    if edit_filter_id:
        try:
            editing_filter = crud.get_filter_line(session, int(edit_filter_id))
        except ValueError:
            raise HTTPException(status_code=404, detail="Ligne filtre introuvable")
        if not editing_filter:
            raise HTTPException(status_code=404, detail="Ligne filtre introuvable")

    edit_belt_id = request.query_params.get("edit_belt")
    if edit_belt_id:
        try:
            editing_belt = crud.get_belt_line(session, int(edit_belt_id))
        except ValueError:
            raise HTTPException(status_code=404, detail="Ligne courroie introuvable")
        if not editing_belt:
            raise HTTPException(status_code=404, detail="Ligne courroie introuvable")

    return templates.TemplateResponse(
        "filters_belts.html",
        {
            "request": request,
            "filters": filters,
            "belts": belts,
            "filter_format_options": FILTER_FORMAT_OPTIONS,
            "filter_format_labels": FILTER_FORMAT_LABELS,
            "import_report": _consume_import_report(request),
            "editing_filter": editing_filter,
            "editing_belt": editing_belt,
        },
    )


@app.get("/plan-de-charge", response_class=HTMLResponse)
def workload_plan(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(
        "plan_de_charge.html",
        {
            "request": request,
        },
    )


class WorkloadPlanSitePayload(BaseModel):
    name: str = Field(..., min_length=1, max_length=255)


class WorkloadPlanCellPayload(BaseModel):
    site_id: int
    day_index: int
    value: Optional[str] = None


class WorkloadPlanCellsPayload(BaseModel):
    updates: List[WorkloadPlanCellPayload] = Field(default_factory=list)


class WorkloadPlanImportPayload(BaseModel):
    version: Optional[int] = None
    sites: List[str] = Field(default_factory=list)
    cells: Dict[str, List[Optional[str]]] = Field(default_factory=dict)


class WorkloadPlanSiteResponse(BaseModel):
    id: int
    name: str
    position: int
    cells: List[str]


class WorkloadPlanResponse(BaseModel):
    version: int
    sites: List[WorkloadPlanSiteResponse]


@app.get("/api/workload-plan", response_model=WorkloadPlanResponse)
def get_workload_plan(session: Session = Depends(get_session)) -> WorkloadPlanResponse:
    sites = crud.list_workload_sites(session)
    payload_sites: List[WorkloadPlanSiteResponse] = []
    for site in sites:
        cells = ["" for _ in range(364)]
        for cell in site.cells:
            if 0 <= cell.day_index < 364 and cell.value:
                cells[cell.day_index] = cell.value
        payload_sites.append(
            WorkloadPlanSiteResponse(
                id=site.id,
                name=site.name,
                position=site.position,
                cells=cells,
            )
        )
    return WorkloadPlanResponse(version=1, sites=payload_sites)


@app.post(
    "/api/workload-plan/sites",
    response_model=WorkloadPlanSiteResponse,
    status_code=201,
)
def create_workload_plan_site(
    payload: WorkloadPlanSitePayload, session: Session = Depends(get_session)
) -> WorkloadPlanSiteResponse:
    try:
        site = crud.create_workload_site(
            session, WorkloadSiteCreate(name=payload.name)
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return WorkloadPlanSiteResponse(
        id=site.id,
        name=site.name,
        position=site.position,
        cells=["" for _ in range(364)],
    )


@app.patch("/api/workload-plan/sites/{site_id}", response_model=WorkloadPlanSiteResponse)
def rename_workload_plan_site(
    site_id: int,
    payload: WorkloadPlanSitePayload,
    session: Session = Depends(get_session),
) -> WorkloadPlanSiteResponse:
    try:
        site = crud.rename_workload_site(
            session, site_id, WorkloadSiteUpdate(name=payload.name)
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    if not site:
        raise HTTPException(status_code=404, detail="Site introuvable")
    cells = ["" for _ in range(364)]
    for cell in site.cells:
        if 0 <= cell.day_index < 364 and cell.value:
            cells[cell.day_index] = cell.value
    return WorkloadPlanSiteResponse(
        id=site.id,
        name=site.name,
        position=site.position,
        cells=cells,
    )


@app.delete("/api/workload-plan/sites/{site_id}", status_code=204)
def delete_workload_plan_site(
    site_id: int, session: Session = Depends(get_session)
) -> Response:
    deleted = crud.delete_workload_site(session, site_id)
    if not deleted:
        raise HTTPException(status_code=404, detail="Site introuvable")
    return Response(status_code=204)


@app.post("/api/workload-plan/cells")
def update_workload_plan_cells(
    payload: WorkloadPlanCellsPayload, session: Session = Depends(get_session)
) -> Dict[str, int]:
    updates = [
        WorkloadCellUpdate(**item.model_dump()) for item in payload.updates or []
    ]
    try:
        count = crud.bulk_update_workload_cells(session, updates)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {"updated": count}


@app.post("/api/workload-plan/import")
def import_workload_plan(
    payload: WorkloadPlanImportPayload, session: Session = Depends(get_session)
) -> Dict[str, int]:
    try:
        count = crud.replace_workload_plan(session, payload.sites, payload.cells)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {"sites": count}


@app.get("/api/workload-plan/export/excel")
def export_workload_plan_excel(session: Session = Depends(get_session)) -> Response:
    sites = crud.list_workload_sites(session)
    buffer = _build_workload_plan_workbook(sites)
    return _template_response(buffer, "plan_de_charge.xlsx")


@app.post("/api/workload-plan/import/excel")
async def import_workload_plan_excel(
    file: UploadFile = File(...),
    session: Session = Depends(get_session),
) -> Dict[str, int]:
    if not file.filename:
        raise HTTPException(status_code=400, detail="Aucun fichier sélectionné.")
    if not file.filename.lower().endswith(ALLOWED_IMPORT_EXTENSIONS):
        raise HTTPException(
            status_code=400,
            detail="Format de fichier non supporté. Merci d'utiliser un fichier Excel (.xlsx).",
        )
    content = await file.read()
    try:
        sites, cells = parse_workload_plan_excel(content)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    count = crud.replace_workload_plan(session, sites, cells)
    return {"sites": count}


@app.post("/filtres-courroies/filtres")
async def create_filter_line(
    site: str = Form(...),
    equipment: str = Form(...),
    efficiency: Optional[str] = Form(None),
    format_type: str = Form(...),
    dimensions: Optional[str] = Form(None),
    quantity: int = Form(...),
    order_week: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    if format_type not in FILTER_FORMAT_LABELS:
        raise HTTPException(status_code=400, detail="Format de filtre invalide")

    if quantity < 1:
        raise HTTPException(status_code=400, detail="Quantité invalide")

    payload = FilterLineCreate(
        site=site.strip(),
        equipment=equipment.strip(),
        efficiency=(efficiency.strip() if efficiency else None),
        format_type=format_type,
        dimensions=(dimensions.strip() if dimensions else None),
        quantity=quantity,
        order_week=(order_week.strip().upper() if order_week else None),
    )
    crud.create_filter_line(session, payload)
    return RedirectResponse("/filtres-courroies", status_code=303)


@app.post("/filtres-courroies/filtres/{line_id}/update")
async def update_filter_line(
    line_id: int,
    site: str = Form(...),
    equipment: str = Form(...),
    efficiency: Optional[str] = Form(None),
    format_type: str = Form(...),
    dimensions: Optional[str] = Form(None),
    quantity: int = Form(...),
    order_week: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    if format_type not in FILTER_FORMAT_LABELS:
        raise HTTPException(status_code=400, detail="Format de filtre invalide")

    if quantity < 1:
        raise HTTPException(status_code=400, detail="Quantité invalide")

    payload = FilterLineUpdate(
        site=site.strip(),
        equipment=equipment.strip(),
        efficiency=(efficiency.strip() if efficiency else None),
        format_type=format_type,
        dimensions=(dimensions.strip() if dimensions else None),
        quantity=quantity,
        order_week=(order_week.strip() if order_week else None),
    )

    updated = crud.update_filter_line(session, line_id, payload)
    if not updated:
        raise HTTPException(status_code=404, detail="Ligne filtre introuvable")

    return RedirectResponse("/filtres-courroies", status_code=303)


@app.post("/filtres-courroies/filtres/import")
async def import_filter_lines(
    file: UploadFile = File(...),
    session: Session = Depends(get_session),
):
    if not file.filename:
        return _store_import_report(
            "/filtres-courroies",
            created=0,
            total=0,
            errors=["Aucun fichier sélectionné."],
            filename="Import filtres",
            singular_label="ligne filtre",
            plural_label="lignes filtre",
        )

    if not file.filename.lower().endswith(ALLOWED_IMPORT_EXTENSIONS):
        return _store_import_report(
            "/filtres-courroies",
            created=0,
            total=0,
            errors=[
                "Format de fichier non supporté. Merci d'utiliser un fichier Excel (.xlsx)."
            ],
            filename=file.filename,
            singular_label="ligne filtre",
            plural_label="lignes filtre",
        )

    content = await file.read()
    try:
        rows = parse_filter_lines_excel(content)
    except ValueError as exc:
        return _store_import_report(
            "/filtres-courroies",
            created=0,
            total=0,
            errors=[str(exc)],
            filename=file.filename,
            singular_label="ligne filtre",
            plural_label="lignes filtre",
        )

    created = 0
    errors: List[str] = []
    for idx, row in enumerate(rows, start=1):
        payload = {
            key: value
            for key, value in row.items()
            if not key.startswith("__")
        }
        row_number = row.get("__row__", idx)
        try:
            crud.create_filter_line(session, FilterLineCreate(**payload))
            created += 1
        except Exception as exc:
            session.rollback()
            errors.append(f"Ligne {row_number} : {exc}")

    return _store_import_report(
        "/filtres-courroies",
        created=created,
        total=len(rows),
        errors=errors,
        filename=file.filename,
        singular_label="ligne filtre",
        plural_label="lignes filtre",
    )


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


@app.post("/filtres-courroies/courroies/{line_id}/update")
async def update_belt_line(
    line_id: int,
    site: str = Form(...),
    equipment: str = Form(...),
    reference: str = Form(...),
    quantity: int = Form(...),
    order_week: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    if quantity < 1:
        raise HTTPException(status_code=400, detail="Quantité invalide")

    payload = BeltLineUpdate(
        site=site.strip(),
        equipment=equipment.strip(),
        reference=reference.strip(),
        quantity=quantity,
        order_week=(order_week.strip() if order_week else None),
    )

    updated = crud.update_belt_line(session, line_id, payload)
    if not updated:
        raise HTTPException(status_code=404, detail="Ligne courroie introuvable")

    return RedirectResponse("/filtres-courroies", status_code=303)


@app.post("/filtres-courroies/courroies/import")
async def import_belt_lines(
    file: UploadFile = File(...),
    session: Session = Depends(get_session),
):
    if not file.filename:
        return _store_import_report(
            "/filtres-courroies",
            created=0,
            total=0,
            errors=["Aucun fichier sélectionné."],
            filename="Import courroies",
            singular_label="ligne courroie",
            plural_label="lignes courroie",
        )

    if not file.filename.lower().endswith(ALLOWED_IMPORT_EXTENSIONS):
        return _store_import_report(
            "/filtres-courroies",
            created=0,
            total=0,
            errors=[
                "Format de fichier non supporté. Merci d'utiliser un fichier Excel (.xlsx)."
            ],
            filename=file.filename,
            singular_label="ligne courroie",
            plural_label="lignes courroie",
        )

    content = await file.read()
    try:
        rows = parse_belt_lines_excel(content)
    except ValueError as exc:
        return _store_import_report(
            "/filtres-courroies",
            created=0,
            total=0,
            errors=[str(exc)],
            filename=file.filename,
            singular_label="ligne courroie",
            plural_label="lignes courroie",
        )

    created = 0
    errors: List[str] = []
    for idx, row in enumerate(rows, start=1):
        payload = {
            key: value
            for key, value in row.items()
            if not key.startswith("__")
        }
        row_number = row.get("__row__", idx)
        try:
            crud.create_belt_line(session, BeltLineCreate(**payload))
            created += 1
        except Exception as exc:
            session.rollback()
            errors.append(f"Ligne {row_number} : {exc}")

    return _store_import_report(
        "/filtres-courroies",
        created=created,
        total=len(rows),
        errors=errors,
        filename=file.filename,
        singular_label="ligne courroie",
        plural_label="lignes courroie",
    )


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
    if not file.filename:
        return _store_import_report(
            "/",
            created=0,
            total=0,
            errors=["Aucun fichier sélectionné."],
            filename="Import",
            singular_label="client",
            plural_label="clients",
        )

    if not file.filename.lower().endswith(ALLOWED_IMPORT_EXTENSIONS):
        return _store_import_report(
            "/",
            created=0,
            total=0,
            errors=[
                "Format de fichier non supporté. Merci d'utiliser un fichier Excel (.xlsx)."
            ],
            filename=file.filename,
            singular_label="client",
            plural_label="clients",
        )

    content = await file.read()
    try:
        rows = parse_clients_excel(content)
    except ValueError as exc:
        return _store_import_report(
            "/",
            created=0,
            total=0,
            errors=[str(exc)],
            filename=file.filename,
            singular_label="client",
            plural_label="clients",
        )

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
            company_name = payload.pop("company_name", None)
            if not company_name:
                raise ValueError("Nom d'entreprise manquant")
            status_value = payload.get("status")
            try:
                statut_bool = _status_to_bool(status_value) if status_value else None
            except HTTPException as exc:
                raise ValueError(str(exc.detail))
            entreprise = crud.ensure_entreprise(
                session,
                name=company_name,
                adresse_facturation=payload.get("billing_address"),
                tag=payload.get("tags"),
                statut=statut_bool,
            )
            payload["entreprise_id"] = entreprise.id
            crud.create_client(
                session,
                ClientCreate(**payload),
                contacts=contacts_payload,
            )
            created += 1
        except Exception as exc:
            session.rollback()
            errors.append(f"Ligne {row_number} : {exc}")

    return _store_import_report(
        "/",
        created=created,
        total=len(rows),
        errors=errors,
        filename=file.filename,
        singular_label="client",
        plural_label="clients",
    )


def _autofit_sheet(sheet, padding: int = 2, max_width: int = 60) -> None:
    for column_cells in sheet.columns:
        max_length = 0
        column = column_cells[0].column_letter
        for cell in column_cells:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        sheet.column_dimensions[column].width = min(max_length + padding, max_width)


def _build_workload_plan_workbook(sites: Iterable) -> BytesIO:
    workbook = Workbook()

    sheet = workbook.active
    sheet.title = "Plan de charge"

    headers = ["Site"] + [f"Jour {index + 1}" for index in range(364)]
    sheet.append(headers)

    def _export_value(value: Optional[str]) -> Optional[Union[int, str]]:
        if value == "bad":
            return 8
        if value == "warn":
            return 4
        return value

    for site in sites:
        values: List[Optional[Union[int, str]]] = [""] * 364
        for cell in getattr(site, "cells", []) or []:
            if cell and 0 <= getattr(cell, "day_index", -1) < 364 and cell.value:
                values[cell.day_index] = _export_value(cell.value)
        sheet.append([getattr(site, "name", "")] + values)

    sheet.freeze_panes = "B2"
    sheet.column_dimensions["A"].width = 32

    legend = workbook.create_sheet("Légende")
    legend.append(["Valeur", "Signification"])
    legend.append(["4", "Orange — intervention à confirmer (4 h)"])
    legend.append(["8", "Rouge — charge pleine (8 h)"])
    legend.append(["ok:4", "Vert 4 h — retour au vert depuis orange"])
    legend.append(["ok:8", "Vert 8 h — retour au vert depuis rouge"])
    legend.append(["ok", "Vert — intervention validée"])
    legend.append(["(vide)", "Aucune information planifiée pour ce jour."])
    legend.freeze_panes = "A2"
    _autofit_sheet(legend, padding=4, max_width=55)

    buffer = BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer


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
    _autofit_sheet(sheet, padding=2, max_width=50)

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

    _autofit_sheet(options_sheet, padding=4, max_width=70)

    buffer = BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer


def _build_filter_import_template() -> BytesIO:
    workbook = Workbook()

    sheet = workbook.active
    sheet.title = "Filtres"

    headers = [
        "site",
        "equipment",
        "format_type",
        "efficiency",
        "dimensions",
        "quantity",
        "order_week",
    ]
    sheet.append(headers)

    sample_rows = [
        {
            "site": "Data Center Nord",
            "equipment": "CTA N°1",
            "format_type": "cadre",
            "efficiency": "ISO ePM1 80%",
            "dimensions": "592 x 592 x 47",
            "quantity": 3,
            "order_week": "S12",
        },
        {
            "site": "Siège social",
            "equipment": "Unité Rooftop",
            "format_type": "poche",
            "efficiency": "F7",
            "dimensions": "287 x 592 x 635",
            "quantity": 2,
            "order_week": "S22",
        },
    ]

    for row in sample_rows:
        sheet.append([row.get(column, "") for column in headers])

    sheet.freeze_panes = "A2"
    _autofit_sheet(sheet, padding=2, max_width=55)

    options_sheet = workbook.create_sheet("Options")
    options_sheet.append(["format_type", "Libellé"])
    options_sheet.freeze_panes = "A2"

    for value, label in FILTER_FORMAT_OPTIONS:
        options_sheet.append([value, label])

    _autofit_sheet(options_sheet, padding=4, max_width=50)

    buffer = BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer


def _build_belt_import_template() -> BytesIO:
    workbook = Workbook()

    sheet = workbook.active
    sheet.title = "Courroies"

    headers = ["site", "equipment", "reference", "quantity", "order_week"]
    sheet.append(headers)

    sample_rows = [
        {
            "site": "Usine Est",
            "equipment": "Compresseur A",
            "reference": "SPB-2000",
            "quantity": 2,
            "order_week": "S08",
        },
        {
            "site": "Usine Est",
            "equipment": "Ventilateur extraction",
            "reference": "XPZ-1250",
            "quantity": 3,
            "order_week": "S30",
        },
        {
            "site": "Entrepôt Sud",
            "equipment": "CTA zone picking",
            "reference": "SPC-1780",
            "quantity": 1,
            "order_week": "",
        },
    ]

    for row in sample_rows:
        sheet.append([row.get(column, "") for column in headers])

    sheet.freeze_panes = "A2"
    _autofit_sheet(sheet, padding=2, max_width=55)

    buffer = BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer


def _template_response(buffer: BytesIO, filename: str) -> Response:
    content = buffer.getvalue()
    headers_dict = {
        "Content-Disposition": (
            f"attachment; filename={filename}; "
            f"filename*=UTF-8''{filename}"
        )
    }
    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers_dict,
    )


@app.get("/clients/import/template")
@app.get("/clients/import/template/")
def download_client_import_template():
    return _template_response(
        _build_client_import_template(), "modele_import_clients.xlsx"
    )


@app.get("/filtres-courroies/filtres/import/template")
@app.get("/filtres-courroies/filtres/import/template/")
def download_filter_import_template():
    return _template_response(
        _build_filter_import_template(), "modele_import_filtres.xlsx"
    )


@app.get("/filtres-courroies/courroies/import/template")
@app.get("/filtres-courroies/courroies/import/template/")
def download_belt_import_template():
    return _template_response(
        _build_belt_import_template(), "modele_import_courroies.xlsx"
    )
