"""Default configuration values for CRM Local."""

from __future__ import annotations

from typing import List, Dict, Any


DEFAULT_PRESTATION_GROUPS: List[Dict[str, Any]] = [
    {
        "category": "Sous-traitance",
        "options": [
            {"value": "analyse_eau", "label": "Analyse d'eau", "budget_code": "S1000", "position": 1},
            {"value": "analyse_huile", "label": "Analyse d'huile", "budget_code": "S1010", "position": 2},
            {
                "value": "analyse_eau_nappe",
                "label": "Analyse eau de nappe",
                "budget_code": "S1020",
                "position": 3,
            },
            {
                "value": "analyse_legionnelle",
                "label": "Analyse légionnelle",
                "budget_code": "S1030",
                "position": 4,
            },
            {
                "value": "analyse_potabilite",
                "label": "Analyse potabilité",
                "budget_code": "S1040",
                "position": 5,
            },
            {"value": "colonnes_seches", "label": "Colonnes sèches", "budget_code": "S1050", "position": 6},
            {"value": "controle_acces", "label": "Contrôle d'accès", "budget_code": "S1060", "position": 7},
            {"value": "controle_ssi", "label": "Contrôle SSI", "budget_code": "S1070", "position": 8},
            {"value": "detection_co", "label": "Détection CO", "budget_code": "S1080", "position": 9},
            {"value": "detection_freon", "label": "Détection fréon", "budget_code": "S1090", "position": 10},
            {"value": "detection_incendie", "label": "Détection incendie", "budget_code": "S2000", "position": 11},
            {"value": "extincteurs", "label": "Extincteurs", "budget_code": "S2010", "position": 12},
            {"value": "exutoires", "label": "Exutoires", "budget_code": "S2020", "position": 13},
            {"value": "gtc", "label": "GTC", "budget_code": "S2030", "position": 14},
            {
                "value": "inspection_video_puits",
                "label": "Inspection vidéo puits",
                "budget_code": "S2040",
                "position": 15,
            },
            {
                "value": "maintenance_cellule_hta",
                "label": "Maintenance cellule HTA",
                "budget_code": "S2050",
                "position": 16,
            },
            {
                "value": "maintenance_constructeur",
                "label": "Maintenance constructeur",
                "budget_code": "S2060",
                "position": 17,
            },
            {
                "value": "maintenance_groupe_electrogene",
                "label": "Maintenance groupe électrogène",
                "budget_code": "S2070",
                "position": 18,
            },
            {
                "value": "maintenance_groupe_froid",
                "label": "Maintenance groupe froid",
                "budget_code": "S2080",
                "position": 19,
            },
            {
                "value": "nettoyage_gaines",
                "label": "Nettoyage de gaines",
                "budget_code": "S3000",
                "position": 20,
            },
            {"value": "onduleurs", "label": "Onduleurs", "budget_code": "S3010", "position": 21},
            {
                "value": "pompe_relevage",
                "label": "Pompe de relevage",
                "budget_code": "S3020",
                "position": 22,
            },
            {
                "value": "portes_automatiques",
                "label": "Portes automatiques",
                "budget_code": "S3030",
                "position": 23,
            },
            {
                "value": "portes_coupe_feu",
                "label": "Portes coupe-feu",
                "budget_code": "S3040",
                "position": 24,
            },
            {"value": "ramonage", "label": "Ramonage", "budget_code": "S3050", "position": 25},
            {"value": "relamping", "label": "Relamping", "budget_code": "S3060", "position": 26},
            {
                "value": "separateur_hydrocarbures",
                "label": "Séparateur hydrocarbures",
                "budget_code": "S3070",
                "position": 27,
            },
            {"value": "sorbonnes", "label": "Sorbonnes", "budget_code": "S4000", "position": 28},
            {
                "value": "table_elevatrice",
                "label": "Table élévatrice",
                "budget_code": "S4010",
                "position": 29,
            },
            {
                "value": "telesurveillance",
                "label": "Télésurveillance",
                "budget_code": "S4020",
                "position": 30,
            },
            {"value": "thermographie", "label": "Thermographie", "budget_code": "S4030", "position": 31},
            {"value": "traitement_eau", "label": "Traitement d'eau", "budget_code": "S4040", "position": 32},
            {
                "value": "video_interphonie",
                "label": "Vidéo et interphonie",
                "budget_code": "S4050",
                "position": 33,
            },
        ],
    },
    {
        "category": "Locations",
        "options": [
            {
                "value": "location_echafaudage",
                "label": "Location échafaudage",
                "budget_code": "L1010",
                "position": 1,
            },
            {
                "value": "location_groupe_electrogene",
                "label": "Location groupe électrogène",
                "budget_code": "L1020",
                "position": 2,
            },
            {
                "value": "location_nacelle",
                "label": "Location nacelle",
                "budget_code": "L1030",
                "position": 3,
            },
        ],
    },
]

