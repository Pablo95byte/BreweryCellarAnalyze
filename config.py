#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Configurazioni centrali per Tank Analysis Tool
"""

# ============= INFO APPLICAZIONE =============
APP_TITLE = "Analisi Tank - PA "
APP_VERSION = "1.0"
APP_AUTHOR = "Paolo Aru"
APP_EMAIL = "paolo_aru@heinekenitalia.it"
APP_DEPT = "ASS_ST"

# ============= MAPPING MATERIALI =============
MATERIAL_MAPPING = {
    '7': 'ichnusa',
    '8': 'non filtrata',
    '9': 'cruda',
    '28': 'ambra limpida',
    '0': 'vuoto',
    '10': 'ich(prop)',
    '32': 'Recovered Beer',
    '36': 'Recovered Beer',
    '3': 'NF Bottle',
    '1': 'Ich Bottle',
    '2': 'Ich Fusti',
}

MATERIAL_DEFAULT_EMPTY = '(vuoto)'

# ============= FORMULE =============
# Coefficienti per f(A) = ((a*G + b)*G + c)*G + d
FA_COEFFICIENTS = {
    'a': 0.0000188792,
    'b': 0.003646886,
    'c': 1.001077,
    'd': -0.01223565
}

# Formula completa per documentazione
FA_FORMULA = "f(A) = ((0.0000188792 × G + 0.003646886) × G + 1.001077) × G - 0.01223565"

# ============= FORMATI DATE =============
DATE_FORMATS = [
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%Y/%m/%d %H:%M:%S"
]

# ============= REGEX PATTERNS =============
import re

REGEX_PATTERNS = {
    # Accetta spazi opzionali tra tipo e numero, e spazi finali
    'avg': re.compile(r"^(FST|BBT|RBT)\s*([0-9]+)\s*Average\s*(Plato|Gravity)\s*$", re.I),
    'level': re.compile(r"^(FST|BBT|RBT)\s*([0-9]+)\s*Level\s*$", re.I),
    'material': re.compile(r"^(FST|BBT|RBT)\s*([0-9]+)\s*Material\s*$", re.I),
}

# ============= UI SETTINGS =============
WINDOW_SIZE = "1280x760"
WINDOW_MIN_SIZE = (1180, 700)
SPLASH_DURATION = 2500  # millisecondi

# ============= COLORI =============
COLORS = {
    'decrease': '#ffcccc',  # Rosso chiaro per cali
    'increase': '#ccffcc',  # Verde chiaro per aumenti
    'info': '#555555',
    'link': '#0066cc',
}

# ============= SOGLIE =============
THRESHOLDS = {
    'significant_level_change': 10.0,  # hl - per evidenziare variazioni
}

# ============= EXPORT DEFAULTS =============
EXPORT_FILENAMES = {
    'tank_csv': 'per_tank_gravity_volume_material_fa_kg.csv',
    'material_csv': 'per_material_somma_kg_fa.csv',
    'debug_csv': 'debug_dettaglio_calcoli.csv',
    'variations_csv': 'variazioni_level_kg.csv',
    'raw_csv': 'dati_raw_completi.csv',
    'xlsx': 'report_tank_material_fa_kg.xlsx',
    'pdf': 'grafici_analisi_tank.pdf',
}