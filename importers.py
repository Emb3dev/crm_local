from __future__ import annotations

from io import BytesIO
from typing import Dict, List, Optional, Tuple, Union
import unicodedata
import re

from openpyxl import load_workbook


EXPECTED_FIELDS = {
    "company_name",
    "name",
}

COLUMN_ALIASES: Dict[str, str] = {
    "company_name": "company_name",
    "entreprise": "company_name",
    "nom_entreprise": "company_name",
    "societe": "company_name",
    "raison_sociale": "company_name",
    "name": "name",
    "client": "name",
    "nom_client": "name",
    "contact": "name",
    "email": "email",
    "mail": "email",
    "courriel": "email",
    "phone": "phone",
    "telephone": "phone",
    "tel": "phone",
    "technician_name": "technician_name",
    "technicien": "technician_name",
    "technicien_referent": "technician_name",
    "referent": "technician_name",
    "billing_address": "billing_address",
    "adresse": "billing_address",
    "adresse_facturation": "billing_address",
    "depannage": "depannage",
    "type_depannage": "depannage",
    "astreinte": "astreinte",
    "tags": "tags",
    "tag": "tags",
    "status": "status",
    "statut": "status",
}

CONTACT_FIELD_ALIASES = {
    "name": "name",
    "nom": "name",
    "prenom": "name",
    "nom_complet": "name",
    "email": "email",
    "mail": "email",
    "courriel": "email",
    "phone": "phone",
    "telephone": "phone",
    "tel": "phone",
    "mobile": "phone",
}

HeaderType = Union[str, Tuple[str, int, str]]

STATUS_ALIASES = {
    "actif": "actif",
    "active": "actif",
    "oui": "actif",
    "true": "actif",
    "1": "actif",
    "inactif": "inactif",
    "inactive": "inactif",
    "non": "inactif",
    "false": "inactif",
    "0": "inactif",
}

DEPANNAGE_CHOICES = {
    "refacturable",
    "non_refacturable",
}

ASTREINTE_CHOICES = {
    "incluse_non_refacturable",
    "incluse_refacturable",
    "pas_d_astreinte",
}

FILTER_EXPECTED_FIELDS = {"site", "equipment", "format_type"}

FILTER_COLUMN_ALIASES: Dict[str, str] = {
    "site": "site",
    "lieu": "site",
    "equipement": "equipment",
    "equipment": "equipment",
    "machine": "equipment",
    "format": "format_type",
    "format_type": "format_type",
    "type": "format_type",
    "efficacite": "efficiency",
    "efficiency": "efficiency",
    "classe": "efficiency",
    "dimensions": "dimensions",
    "dimension": "dimensions",
    "taille": "dimensions",
    "quantite": "quantity",
    "quantity": "quantity",
    "qte": "quantity",
    "commande": "order_week",
    "semaine_commande": "order_week",
    "order_week": "order_week",
    "semaine": "order_week",
}

FILTER_FORMAT_LOOKUP = {
    "cousus_sur_fil": "cousus_sur_fil",
    "format_cousus_sur_fil": "cousus_sur_fil",
    "cadre": "cadre",
    "format_cadre": "cadre",
    "multi_diedre": "multi_diedre",
    "format_multi_diedre": "multi_diedre",
    "poche": "poche",
    "format_poche": "poche",
}

BELT_EXPECTED_FIELDS = {"site", "equipment", "reference"}

BELT_COLUMN_ALIASES: Dict[str, str] = {
    "site": "site",
    "lieu": "site",
    "equipement": "equipment",
    "equipment": "equipment",
    "machine": "equipment",
    "reference": "reference",
    "ref": "reference",
    "code": "reference",
    "quantite": "quantity",
    "quantity": "quantity",
    "qte": "quantity",
    "commande": "order_week",
    "semaine_commande": "order_week",
    "order_week": "order_week",
    "semaine": "order_week",
}


def _normalize_header(header: Optional[str]) -> str:
    if header is None:
        return ""
    value = str(header).strip().lower()
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    for char in ("-", ".", "'", "\u00a0"):
        value = value.replace(char, " ")
    return "_".join(part for part in value.split() if part)


def _coerce_value(value):
    if value is None:
        return None
    if isinstance(value, str):
        cleaned = value.strip()
        return cleaned or None
    if isinstance(value, (int, float)):
        if isinstance(value, float):
            if value.is_integer():
                return str(int(value))
            return str(value).rstrip("0").rstrip(".")
        return str(value)
    return str(value).strip() or None


CONTACT_HEADER_RE = re.compile(r"contact_?(\d+)_([a-z0-9_]+)")


def _resolve_header(header: str) -> HeaderType:
    if header in COLUMN_ALIASES:
        return COLUMN_ALIASES[header]
    match = CONTACT_HEADER_RE.match(header)
    if match:
        index = int(match.group(1))
        field_key = match.group(2)
        field = CONTACT_FIELD_ALIASES.get(field_key)
        if field:
            return ("contact", index, field)
    return ""


def _resolve_filter_header(header: str) -> str:
    return FILTER_COLUMN_ALIASES.get(header, "")


def _resolve_belt_header(header: str) -> str:
    return BELT_COLUMN_ALIASES.get(header, "")


def _parse_positive_int(value: str, row_index: int, field_name: str) -> int:
    try:
        if "." in value:
            parsed = float(value)
            if not parsed.is_integer():
                raise ValueError
            quantity = int(parsed)
        else:
            quantity = int(value)
    except ValueError:
        raise ValueError(
            f"Ligne {row_index}: valeur de {field_name} invalide '{value}'."
        ) from None

    if quantity < 1:
        raise ValueError(
            f"Ligne {row_index}: valeur de {field_name} invalide '{value}'."
        )
    return quantity


def parse_clients_excel(content: bytes) -> List[Dict[str, str]]:
    try:
        workbook = load_workbook(BytesIO(content), data_only=True)
    except Exception as exc:  # pragma: no cover - delegated to openpyxl
        raise ValueError(f"Impossible de lire le fichier Excel: {exc}")

    sheet = workbook.active
    try:
        header_row = next(sheet.iter_rows(max_row=1))
    except StopIteration:
        raise ValueError("Le fichier ne contient aucune donnée.")

    headers: List[HeaderType] = [
        _resolve_header(_normalize_header(cell.value)) for cell in header_row
    ]

    if not any(headers):
        raise ValueError("Le fichier ne contient pas d'en-têtes valides.")

    missing = EXPECTED_FIELDS - {
        header for header in headers if isinstance(header, str) and header
    }
    if missing:
        raise ValueError(
            "Colonnes obligatoires manquantes: " + ", ".join(sorted(missing))
        )

    rows: List[Dict[str, str]] = []
    for row_index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        record: Dict[str, str] = {"__row__": row_index}
        empty = True
        for idx, raw_value in enumerate(row):
            header = headers[idx] if idx < len(headers) else ""
            if not header:
                continue
            value = _coerce_value(raw_value)
            if value is None:
                continue
            empty = False
            if isinstance(header, tuple):
                _, contact_index, contact_field = header
                contact_bucket = record.setdefault("__contacts__", {})
                contact_data = contact_bucket.setdefault(contact_index, {})
                contact_data[contact_field] = value
                continue

            key = header
            if key == "status":
                normalized = _normalize_header(value)
                record[key] = STATUS_ALIASES.get(normalized, normalized or None)
            elif key == "depannage":
                normalized = _normalize_header(value)
                if normalized and normalized not in DEPANNAGE_CHOICES:
                    raise ValueError(
                        f"Ligne {row_index}: valeur de dépannage invalide '{value}'."
                    )
                if normalized:
                    record[key] = normalized
            elif key == "astreinte":
                normalized = _normalize_header(value)
                if normalized and normalized not in ASTREINTE_CHOICES:
                    raise ValueError(
                        f"Ligne {row_index}: valeur d'astreinte invalide '{value}'."
                    )
                if normalized:
                    record[key] = normalized
            else:
                record[key] = value

        if empty:
            continue

        missing_fields = [field for field in EXPECTED_FIELDS if not record.get(field)]
        if missing_fields:
            raise ValueError(
                f"Ligne {row_index}: valeurs manquantes pour {', '.join(missing_fields)}"
            )

        if record.get("status") and record["status"] not in STATUS_ALIASES.values():
            raise ValueError(
                f"Ligne {row_index}: statut inconnu '{record['status']}'. Valeurs acceptées: actif, inactif."
            )

        contacts_map = record.pop("__contacts__", {})
        contacts: List[Dict[str, str]] = []
        for order, data in sorted(contacts_map.items()):
            if not data.get("name"):
                if any(data.get(field) for field in ("email", "phone")):
                    raise ValueError(
                        f"Ligne {row_index}: le contact {order} doit avoir un nom."
                    )
                continue
            contacts.append(data)

        if contacts:
            record["contacts"] = contacts

        rows.append(record)

    return rows


def parse_filter_lines_excel(content: bytes) -> List[Dict[str, Union[str, int]]]:
    try:
        workbook = load_workbook(BytesIO(content), data_only=True)
    except Exception as exc:
        raise ValueError(f"Impossible de lire le fichier Excel: {exc}")

    sheet = workbook.active
    try:
        header_row = next(sheet.iter_rows(max_row=1))
    except StopIteration:
        raise ValueError("Le fichier ne contient aucune donnée.")

    headers = [
        _resolve_filter_header(_normalize_header(cell.value)) for cell in header_row
    ]

    if not any(headers):
        raise ValueError("Le fichier ne contient pas d'en-têtes valides.")

    missing = FILTER_EXPECTED_FIELDS - {header for header in headers if header}
    if missing:
        raise ValueError(
            "Colonnes obligatoires manquantes: " + ", ".join(sorted(missing))
        )

    rows: List[Dict[str, Union[str, int]]] = []
    for row_index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        record: Dict[str, Union[str, int]] = {"__row__": row_index}
        empty = True
        for idx, raw_value in enumerate(row):
            header = headers[idx] if idx < len(headers) else ""
            if not header:
                continue
            value = _coerce_value(raw_value)
            if value is None:
                continue
            empty = False
            if header == "format_type":
                normalized = _normalize_header(value)
                resolved = FILTER_FORMAT_LOOKUP.get(normalized)
                if not resolved:
                    raise ValueError(
                        f"Ligne {row_index}: format de filtre inconnu '{value}'."
                    )
                record[header] = resolved
            elif header == "quantity":
                record[header] = _parse_positive_int(value, row_index, "quantité")
            elif header == "order_week":
                record[header] = value.strip().upper()
            else:
                record[header] = value

        if empty:
            continue

        missing_fields = [
            field for field in FILTER_EXPECTED_FIELDS if not record.get(field)
        ]
        if missing_fields:
            raise ValueError(
                "Ligne {row_index}: champ(s) obligatoire(s) manquant(s): "
                + ", ".join(sorted(missing_fields))
            )

        if "quantity" not in record:
            record["quantity"] = 1

        rows.append(record)

    return rows


def parse_belt_lines_excel(content: bytes) -> List[Dict[str, Union[str, int]]]:
    try:
        workbook = load_workbook(BytesIO(content), data_only=True)
    except Exception as exc:
        raise ValueError(f"Impossible de lire le fichier Excel: {exc}")

    sheet = workbook.active
    try:
        header_row = next(sheet.iter_rows(max_row=1))
    except StopIteration:
        raise ValueError("Le fichier ne contient aucune donnée.")

    headers = [
        _resolve_belt_header(_normalize_header(cell.value)) for cell in header_row
    ]

    if not any(headers):
        raise ValueError("Le fichier ne contient pas d'en-têtes valides.")

    missing = BELT_EXPECTED_FIELDS - {header for header in headers if header}
    if missing:
        raise ValueError(
            "Colonnes obligatoires manquantes: " + ", ".join(sorted(missing))
        )

    rows: List[Dict[str, Union[str, int]]] = []
    for row_index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        record: Dict[str, Union[str, int]] = {"__row__": row_index}
        empty = True
        for idx, raw_value in enumerate(row):
            header = headers[idx] if idx < len(headers) else ""
            if not header:
                continue
            value = _coerce_value(raw_value)
            if value is None:
                continue
            empty = False
            if header == "quantity":
                record[header] = _parse_positive_int(value, row_index, "quantité")
            elif header == "order_week":
                record[header] = value.strip().upper()
            else:
                record[header] = value

        if empty:
            continue

        missing_fields = [
            field for field in BELT_EXPECTED_FIELDS if not record.get(field)
        ]
        if missing_fields:
            raise ValueError(
                "Ligne {row_index}: champ(s) obligatoire(s) manquant(s): "
                + ", ".join(sorted(missing_fields))
            )

        if "quantity" not in record:
            record["quantity"] = 1

        rows.append(record)

    return rows
