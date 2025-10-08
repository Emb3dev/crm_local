from __future__ import annotations

from io import BytesIO
from typing import Dict, List
import unicodedata

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


def _normalize_header(header: str | None) -> str:
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

    headers = [COLUMN_ALIASES.get(_normalize_header(cell.value), "") for cell in header_row]

    if not any(headers):
        raise ValueError("Le fichier ne contient pas d'en-têtes valides.")

    missing = EXPECTED_FIELDS - set(filter(None, headers))
    if missing:
        raise ValueError(
            "Colonnes obligatoires manquantes: " + ", ".join(sorted(missing))
        )

    rows: List[Dict[str, str]] = []
    for row_index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        record: Dict[str, str] = {"__row__": row_index}
        empty = True
        for idx, raw_value in enumerate(row):
            key = headers[idx] if idx < len(headers) else ""
            if not key:
                continue
            value = _coerce_value(raw_value)
            if value is None:
                continue
            empty = False
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

        rows.append(record)

    return rows
