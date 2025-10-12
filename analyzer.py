#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Analizzatore dati CSV per Tank Analysis Tool
"""

import csv
from datetime import datetime
from collections import defaultdict

from config import REGEX_PATTERNS
from utils import parse_time, to_float, calculate_fA, sanitize_level, normalize_material, is_valid_value


class TankAnalyzer:
    """
    Classe per analizzare i dati dei tank dal CSV
    """
    
    def __init__(self, csv_path):
        self.path = csv_path
        self.header = []
        self.rows = []
        self.time_idx = None
        self.avg_cols = []      # (idx, tank_key, family)
        self.level_idx = {}     # tank_key -> idx
        self.material_idx = {}  # tank_key -> idx
        self.min_time = None
        self.max_time = None
        self._load_csv()
    
    def _load_csv(self):
        """Carica e parsea il CSV"""
        with open(self.path, 'r', encoding='utf-8-sig', newline='') as f:
            reader = csv.reader(f)
            rows = list(reader)
        
        if not rows:
            raise ValueError("CSV vuoto")
        
        self.header = [h.strip() for h in rows[0]]
        self.rows = rows[1:]
        
        self._identify_columns()
        self._calculate_time_range()
    
    def _identify_columns(self):
        """Identifica le colonne rilevanti nel CSV"""
        # Trova colonna Time
        for i, h in enumerate(self.header):
            if h.strip().lower() == 'time':
                self.time_idx = i
                break
        
        # Identifica colonne Average, Level, Material
        patterns = REGEX_PATTERNS
        
        for idx, col in enumerate(self.header):
            # Average Gravity/Plato
            m = patterns['avg'].match(col)
            if m:
                family = m.group(1).upper()
                num = m.group(2)
                tank_key = f"{family}{num}"
                self.avg_cols.append((idx, tank_key, family))
            
            # Level
            m2 = patterns['level'].match(col)
            if m2:
                family = m2.group(1).upper()
                num = m2.group(2)
                self.level_idx[f"{family}{num}"] = idx
            
            # Material
            m3 = patterns['material'].match(col)
            if m3:
                family = m3.group(1).upper()
                num = m3.group(2)
                self.material_idx[f"{family}{num}"] = idx
    
    def _calculate_time_range(self):
        """Calcola il range temporale dei dati"""
        if self.time_idx is None:
            return
        
        for row in self.rows:
            if self.time_idx < len(row):
                dt = parse_time(row[self.time_idx])
                if dt:
                    if self.min_time is None or dt < self.min_time:
                        self.min_time = dt
                    if self.max_time is None or dt > self.max_time:
                        self.max_time = dt
    
    def analyze(self, t_from=None, t_to=None, include_fst=True, include_bbt=True):
        """
        Analizza i dati e restituisce aggregazioni per tank, materiale e debug
        
        Returns:
            tuple: (tank_rows, material_rows, debug_data)
        """
        by_tank = {}
        by_material = {}
        debug_data = []
        
        for row in self.rows:
            # Filtro temporale
            dt = self._get_row_timestamp(row)
            if not self._passes_time_filter(dt, t_from, t_to):
                continue
            
            # Processa ogni tank
            for idx, tank_key, family in self.avg_cols:
                if not self._passes_family_filter(family, include_fst, include_bbt):
                    continue
                
                # Estrai dati
                data = self._extract_tank_data(row, idx, tank_key)
                if data is None:
                    continue
                
                gravity, level, material, fa_value, kg_extracted = data
                
                # Aggiungi a debug
                debug_data.append((dt, tank_key, material, gravity, level, fa_value, kg_extracted))
                
                # Aggrega per tank
                self._aggregate_by_tank(by_tank, tank_key, gravity, level, material, fa_value, kg_extracted, dt)
                
                # Aggrega per materiale
                self._aggregate_by_material(by_material, material, fa_value, kg_extracted)
        
        # Ordina risultati
        tank_rows = self._sort_tank_results(by_tank)
        material_rows = self._sort_material_results(by_material)
        debug_data.sort(key=lambda x: x[0] if x[0] else datetime.min)
        
        return tank_rows, material_rows, debug_data
    
    def analyze_all_days(self, include_fst=True, include_bbt=True):
        """
        Analizza tutti i giorni disponibili per grafici temporali
        
        Returns:
            dict: {day_string: {'kg': total, 'by_material': {}, 'by_tank': {}}}
        """
        if self.time_idx is None:
            return {}
        
        daily_data = defaultdict(lambda: {
            'kg': 0.0,
            'by_material': defaultdict(float),
            'by_tank': defaultdict(float)
        })
        
        for row in self.rows:
            dt = self._get_row_timestamp(row)
            if dt is None:
                continue
            
            day_key = dt.date().strftime("%Y-%m-%d")
            
            for idx, tank_key, family in self.avg_cols:
                if not self._passes_family_filter(family, include_fst, include_bbt):
                    continue
                
                data = self._extract_tank_data(row, idx, tank_key)
                if data is None:
                    continue
                
                _, _, material, _, kg_extracted = data
                
                daily_data[day_key]['kg'] += kg_extracted
                daily_data[day_key]['by_material'][material] += kg_extracted
                daily_data[day_key]['by_tank'][tank_key] += kg_extracted
        
        return daily_data
    
    # ============= METODI HELPER PRIVATI =============
    
    def _get_row_timestamp(self, row):
        """Estrae il timestamp da una riga"""
        if self.time_idx is None or self.time_idx >= len(row):
            return None
        return parse_time(row[self.time_idx])
    
    def _passes_time_filter(self, dt, t_from, t_to):
        """Verifica se il timestamp passa il filtro temporale"""
        if self.time_idx is not None and (t_from or t_to):
            if dt is None:
                return False
            if t_from and dt < t_from:
                return False
            if t_to and dt > t_to:
                return False
        return True
    
    def _passes_family_filter(self, family, include_fst, include_bbt):
        """Verifica se la famiglia di tank passa il filtro"""
        if family == 'FST' and not include_fst:
            return False
        if family == 'BBT' and not include_bbt:
            return False
        return True
    
    def _extract_tank_data(self, row, idx, tank_key):
        """
        Estrae e valida i dati di un tank da una riga
        
        Returns:
            tuple or None: (gravity, level, material, fa_value, kg_extracted) or None se non valido
        """
        if idx >= len(row):
            return None
        
        # Gravity
        gravity = to_float(row[idx])
        if not is_valid_value(gravity):
            return None
        
        # f(A)
        fa_value = calculate_fA(gravity)
        if not is_valid_value(fa_value):
            return None
        
        # Level
        level_idx = self.level_idx.get(tank_key)
        level = to_float(row[level_idx]) if (level_idx is not None and level_idx < len(row)) else None
        level = sanitize_level(level)
        
        # Material
        mat_idx = self.material_idx.get(tank_key)
        material = row[mat_idx] if (mat_idx is not None and mat_idx < len(row)) else None
        material = normalize_material(material)
        
        # Kg estratto
        kg_extracted = fa_value * level
        
        return (gravity, level, material, fa_value, kg_extracted)
    
    def _aggregate_by_tank(self, by_tank, tank_key, gravity, level, material, fa_value, kg_extracted, timestamp):
        """Aggrega dati per tank"""
        if tank_key not in by_tank:
            by_tank[tank_key] = {
                'G_last': None,
                'V_last': None,
                'M_last': None,
                't_last': None,
                'sum_fA': 0.0,
                'sum_kg': 0.0,
                'count': 0
            }
        
        rec = by_tank[tank_key]
        
        # Aggiorna ultimi valori se timestamp più recente
        if timestamp is not None:
            if rec['t_last'] is None or timestamp > rec['t_last']:
                rec['t_last'] = timestamp
                rec['G_last'] = gravity
                rec['V_last'] = level
                rec['M_last'] = material
        else:
            rec['G_last'] = gravity
            rec['V_last'] = level
            rec['M_last'] = material
        
        rec['sum_fA'] += fa_value
        rec['sum_kg'] += kg_extracted
        rec['count'] += 1
    
    def _aggregate_by_material(self, by_material, material, fa_value, kg_extracted):
        """Aggrega dati per materiale"""
        if material not in by_material:
            by_material[material] = {
                'sum_kg': 0.0,
                'sum_fA': 0.0,
                'count': 0
            }
        
        rec = by_material[material]
        rec['sum_kg'] += kg_extracted
        rec['sum_fA'] += fa_value
        rec['count'] += 1
    
    def _sort_tank_results(self, by_tank):
        """Converte e ordina risultati per tank"""
        results = []
        for tank, stats in by_tank.items():
            results.append((
                tank,
                stats['M_last'],
                stats['G_last'],
                stats['V_last'],
                stats['sum_fA'],
                stats['sum_kg'],
                stats['count']
            ))
        # Ordina per kg discendente
        results.sort(key=lambda x: (-x[5], x[0]))
        return results
    
    def _sort_material_results(self, by_material):
        """Converte e ordina risultati per materiale"""
        results = []
        for material, stats in by_material.items():
            results.append((
                material,
                stats['sum_kg'],
                stats['sum_fA'],
                stats['count']
            ))
        # Ordina per kg discendente
        results.sort(key=lambda x: (-x[1], x[0]))
        return results
