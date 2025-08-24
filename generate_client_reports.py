#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
generate_client_reports.py - Skrypt do generowania raportu kosztów dla klienta w XLSX.

Instrukcje użycia:
- Uruchamiany z argumentem ścieżki do folderu Raporty.
- Odczytuje data.json, oblicza koszty z narzutami dynamicznymi.
- Dodaje gięcie pełne, dodatkowe.
- Wylicza koszty jednostkowe, całkowite, cenę sprzedaży z marżą.
- Zapisuje do "Raport dla klienta.xlsx" i dopisuje do logu.

Wymaga bibliotek: openpyxl, json, os
"""

import sys
import os
import json
from openpyxl import Workbook

def main(raporty_path):
    # Odczyt data.json
    json_path = os.path.join(raporty_path, "data.json")
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    all_parts = data['all_parts']
    marza = data['marza']

    # Log
    log_path = os.path.join(raporty_path, "cost_calculation_log.txt")
    with open(log_path, 'a', encoding='utf-8') as log:
        log.write("\n--- Raport kosztów dla klienta (z narzutami) ---\n")

    # Tworzenie XLSX
    wb = Workbook()
    ws = wb.active
    ws.title = "Koszty dla klienta"

    # Nagłówki
    ws.append(["ID", "Nazwa", "Materiał", "Grubość", "Ilość", "Koszt jednostkowy (z narzutami)", "Gięcie", "Dodatkowe", "Koszt całkowity", "Cena sprzedaży z marżą"])

    total_cost = 0.0
    for part in all_parts:
        unit_cost = part['cost_per_unit']  # Z narzutami dynamicznymi
        bending = part['bending_per_unit']
        additional = part['additional_per_unit']
        total_unit = unit_cost + bending + additional
        total_part = total_unit * part['qty']
        sales_price = total_unit * (1 + marza / 100)
        total_cost += total_part

        ws.append([part['id'], part['name'], part['material'], part['thickness'], part['qty'], unit_cost, bending, additional, total_part, sales_price])

        # Log
        log.write(f"{part['name']}: Jednostkowy {unit_cost}, Gięcie {bending}, Dodatkowe {additional}, Całkowity {total_part}, Sprzedaż {sales_price}\n")

    ws.append(["", "", "", "", "Suma", total_cost])

    wb.save(os.path.join(raporty_path, "Raport dla klienta.xlsx"))

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Użycie: python generate_client_reports.py <raporty_path>")
    else:
        main(sys.argv[1])
        