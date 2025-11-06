#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script di debug per verificare colonne duplicate
"""

import csv
import sys
from collections import Counter
from config import REGEX_PATTERNS

def debug_csv_columns(csv_path):
    """Analizza le colonne del CSV per trovare duplicati"""

    with open(csv_path, 'r', encoding='utf-8-sig', newline='') as f:
        reader = csv.reader(f)
        rows = list(reader)

    if not rows:
        print("CSV vuoto!")
        return

    header = [h.strip() for h in rows[0]]

    print("=" * 80)
    print("ANALISI COLONNE CSV")
    print("=" * 80)
    print(f"\nTotale colonne: {len(header)}")
    print(f"Totale righe dati: {len(rows)-1}")

    # Verifica duplicati nel header
    print("\n" + "=" * 80)
    print("VERIFICA COLONNE DUPLICATE")
    print("=" * 80)
    counter = Counter(header)
    duplicates = {col: count for col, count in counter.items() if count > 1}

    if duplicates:
        print("\n⚠️  TROVATE COLONNE DUPLICATE:")
        for col, count in duplicates.items():
            print(f"  - '{col}' appare {count} volte")
    else:
        print("\n✓ Nessuna colonna duplicata trovata")

    # Identifica colonne Average
    print("\n" + "=" * 80)
    print("COLONNE AVERAGE GRAVITY/PLATO IDENTIFICATE")
    print("=" * 80)

    patterns = REGEX_PATTERNS
    avg_cols = []

    for idx, col in enumerate(header):
        col_normalized = ' '.join(col.split())
        m = patterns['avg'].match(col_normalized)
        if m:
            family = m.group(1).upper()
            num = m.group(2)
            tank_key = f"{family}{num}"
            avg_cols.append((idx, tank_key, family, col))
            print(f"  [{idx:2d}] {tank_key:8s} <- '{col}'")

    # Verifica se ci sono tank_key duplicati
    print("\n" + "=" * 80)
    print("VERIFICA TANK DUPLICATI")
    print("=" * 80)
    tank_counter = Counter([tk for _, tk, _, _ in avg_cols])
    tank_duplicates = {tank: count for tank, count in tank_counter.items() if count > 1}

    if tank_duplicates:
        print("\n⚠️  TROVATI TANK PROCESSATI PIÙ VOLTE:")
        for tank, count in tank_duplicates.items():
            print(f"  - {tank} appare {count} volte")
            matching_cols = [col for _, tk, _, col in avg_cols if tk == tank]
            for col in matching_cols:
                print(f"      <- '{col}'")
        print("\n⚠️  QUESTO CAUSA IL RADDOPPIO DEI VALORI!")
    else:
        print("\n✓ Nessun tank duplicato")

    # Analizza un giorno specifico per vedere quante righe ci sono
    print("\n" + "=" * 80)
    print("ANALISI RIGHE PER GIORNO")
    print("=" * 80)

    from utils import parse_time

    days_count = {}
    time_idx = None

    for i, h in enumerate(header):
        if h.strip().lower() == 'time':
            time_idx = i
            break

    if time_idx is not None:
        for row in rows[1:]:
            if time_idx < len(row):
                dt = parse_time(row[time_idx])
                if dt:
                    day_key = dt.date().strftime("%Y-%m-%d")
                    days_count[day_key] = days_count.get(day_key, 0) + 1

        print(f"\nGiorni trovati: {len(days_count)}")
        for day in sorted(days_count.keys())[:10]:  # Mostra primi 10 giorni
            print(f"  {day}: {days_count[day]} righe")

        if len(days_count) > 10:
            print(f"  ... (altri {len(days_count)-10} giorni)")

    print("\n" + "=" * 80)

if __name__ == '__main__':
    if len(sys.argv) > 1:
        csv_path = sys.argv[1]
    else:
        csv_path = 'data_example/tank.csv'

    print(f"\nAnalizzando: {csv_path}\n")
    debug_csv_columns(csv_path)
