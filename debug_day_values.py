#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script di debug per analizzare i valori f(A) giorno per giorno
"""

import sys
from analyzer import TankAnalyzer
from utils import fmt_it
from collections import defaultdict

def debug_day_values(csv_path, target_day=None):
    """Analizza i valori f(A) e kg estratti per giorno"""

    analyzer = TankAnalyzer(csv_path)

    print("=" * 80)
    print("ANALISI VALORI f(A) PER GIORNO E TANK")
    print("=" * 80)

    # Raggruppa per giorno e tank
    daily_by_tank = defaultdict(lambda: defaultdict(list))

    for row in analyzer.rows:
        dt = analyzer._get_row_timestamp(row)
        if dt is None:
            continue

        day_key = dt.date().strftime("%Y-%m-%d")

        for idx, tank_key, family in analyzer.avg_cols:
            data = analyzer._extract_tank_data(row, idx, tank_key)
            if data is None:
                continue

            gravity, level, material, fa_value, kg_extracted = data
            daily_by_tank[day_key][tank_key].append({
                'timestamp': dt.strftime("%Y-%m-%d %H:%M:%S"),
                'gravity': gravity,
                'level': level,
                'material': material,
                'fa_value': fa_value,
                'kg_extracted': kg_extracted
            })

    # Analizza ogni giorno
    for day in sorted(daily_by_tank.keys()):
        if target_day and day != target_day:
            continue

        print(f"\n{'=' * 80}")
        print(f"GIORNO: {day}")
        print('=' * 80)

        for tank in sorted(daily_by_tank[day].keys()):
            measurements = daily_by_tank[day][tank]
            print(f"\n  Tank: {tank}")
            print(f"  Numero misurazioni: {len(measurements)}")

            # Calcola somma e media di f(A)
            sum_fa = sum(m['fa_value'] for m in measurements)
            avg_fa = sum_fa / len(measurements)
            sum_kg = sum(m['kg_extracted'] for m in measurements)
            avg_kg = sum_kg / len(measurements)

            print(f"  Somma f(A): {fmt_it(sum_fa, 6)}")
            print(f"  Media f(A): {fmt_it(avg_fa, 6)}")
            print(f"  Somma Kg estratto: {fmt_it(sum_kg, 3)}")
            print(f"  Media Kg estratto: {fmt_it(avg_kg, 3)}")

            # Mostra dettagli di ogni misurazione
            print(f"\n  Dettaglio misurazioni:")
            for i, m in enumerate(measurements, 1):
                print(f"    [{i}] {m['timestamp']} - Gravity: {fmt_it(m['gravity'], 2)}, "
                      f"Level: {fmt_it(m['level'], 2)}, f(A): {fmt_it(m['fa_value'], 6)}, "
                      f"Kg: {fmt_it(m['kg_extracted'], 3)}")

            # Verifica se ci sono valori identici (possibili duplicati)
            fa_values = [m['fa_value'] for m in measurements]
            if len(fa_values) != len(set(fa_values)):
                print(f"\n  ⚠️  ATTENZIONE: Ci sono valori f(A) duplicati per questo tank!")

    print("\n" + "=" * 80)

if __name__ == '__main__':
    csv_path = 'data_example/tank.csv'
    target_day = None

    if len(sys.argv) > 1:
        csv_path = sys.argv[1]
    if len(sys.argv) > 2:
        target_day = sys.argv[2]

    print(f"\nAnalizzando: {csv_path}")
    if target_day:
        print(f"Giorno specifico: {target_day}")
    print()

    debug_day_values(csv_path, target_day)
