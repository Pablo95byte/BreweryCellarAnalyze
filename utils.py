#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Utilità, formattazione e calcoli per Tank Analysis Tool
"""

import math
from datetime import datetime
from config import FA_COEFFICIENTS, DATE_FORMATS, MATERIAL_MAPPING, MATERIAL_DEFAULT_EMPTY


# ============= FORMATTAZIONE =============

def to_float(s):
    """Converte una stringa in float, gestendo formati europei"""
    if s is None:
        return None
    s = str(s).strip()
    if not s:
        return None
    # Gestione formato europeo (1.234,56)
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def fmt_it(x, nd=2):
    """Formatta un numero in stile italiano (1.234,56)"""
    if x is None:
        return ""
    s = f"{x:,.{nd}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def parse_time(s):
    """Parse una stringa di data/ora in datetime"""
    if s is None:
        return None
    s = s.strip()
    if not s:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return None


# ============= CALCOLI =============

def calculate_fA(gravity):
    """
    Calcola f(A) dalla gravity usando la formula:
    f(A) = ((a*G + b)*G + c)*G + d
    """
    if gravity is None:
        return None
    
    coef = FA_COEFFICIENTS
    result = ((coef['a'] * gravity + coef['b']) * gravity + coef['c']) * gravity + coef['d']
    return result


def calculate_kg_extracted(gravity, level):
    """
    Calcola i Kg estratti:
    Kg = f(A) × Level
    """
    if gravity is None or level is None:
        return None
    
    fa_value = calculate_fA(gravity)
    if fa_value is None:
        return None
    
    # Gestione valori negativi o NaN
    if isinstance(level, float) and math.isnan(level):
        level = 0.0
    if level < 0:
        level = 0.0
    
    return fa_value * level


def is_valid_value(value):
    """Verifica se un valore è valido (non None, non NaN)"""
    if value is None:
        return False
    if isinstance(value, float) and math.isnan(value):
        return False
    return True


# ============= NORMALIZZAZIONE MATERIALI =============

def normalize_material(val):
    """
    Normalizza il valore del materiale:
    - Se è un codice numerico, usa il mapping
    - Se è testo, cerca pattern conosciuti
    - Altrimenti mantiene il valore originale
    """
    if val is None:
        return MATERIAL_DEFAULT_EMPTY
    
    s = str(val).strip()
    if not s:
        return MATERIAL_DEFAULT_EMPTY
    
    # 1. Prova come codice diretto
    if s in MATERIAL_MAPPING:
        return MATERIAL_MAPPING[s]
    
    # 2. Prova come numero (include codice 0)
    try:
        k = str(int(float(s)))
        if k in MATERIAL_MAPPING:
            return MATERIAL_MAPPING[k]
        return k
    except Exception:
        pass
    
    # 3. Prova come testo (normalizza accentazione)
    low = s.lower().strip()
    low = _remove_accents(low)
    
    # Pattern matching per testi
    if 'ichnusa' in low:
        return 'ichnusa'
    if 'non filtrata' in low or 'nonfiltrata' in low:
        return 'non filtrata'
    if 'cruda' in low:
        return 'cruda'
    if 'ambra' in low and 'limpida' in low:
        return 'ambra limpida'
    
    return s


def _remove_accents(text):
    """Rimuove accenti dalle lettere"""
    accent_map = {
        'ì': 'i', 'í': 'i', 'ï': 'i', 'î': 'i',
        'à': 'a', 'á': 'a', 'ä': 'a', 'â': 'a',
        'è': 'e', 'é': 'e', 'ë': 'e', 'ê': 'e',
        'ò': 'o', 'ó': 'o', 'ö': 'o', 'ô': 'o',
        'ù': 'u', 'ú': 'u', 'ü': 'u', 'û': 'u',
    }
    return ''.join(accent_map.get(ch, ch) for ch in text)


# ============= VALIDAZIONE DATI =============

def sanitize_level(level):
    """
    Sanifica il valore di Level:
    - None o NaN → 0.0
    - Negativo → 0.0
    - Altrimenti mantiene il valore
    """
    if level is None or (isinstance(level, float) and math.isnan(level)):
        return 0.0
    if level < 0:
        return 0.0
    return level
