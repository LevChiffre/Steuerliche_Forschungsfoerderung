import pandas as pd
import holidays
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from calendar import monthrange, month_name




# Globale Einstellungen
MONTHS_GER = ["dummy","Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Detember"]
YEAR = 2023
EXCEL_NAME = "Stundennachweis_" + str(YEAR) + ".xlsx" #Excel wird in das gleiche Verzeichnis wie dieses Skript gespeichtert

# Feiertage für Hessen laden
hessian_holidays = holidays.DE(prov='HE', years=YEAR)

# Funktion, um alle Tage des Jahres zu generieren
def generate_all_days(year):
    days = []
    start_date = datetime(year, 1, 1)
    end_date = datetime(year, 12, 31)
    current_date = start_date
    while current_date <= end_date:
        days.append(current_date.date())
        current_date += timedelta(days=1)
    return days

# Excel-Datei mit openpyxl erstellen und stylen
wb = Workbook()
ws = wb.active
ws.title = "Arbeitstage_Kalender"

# Jahr-Überschrift
ws['A1'] = f"Kalender {YEAR}"
ws['A1'].font = Font(bold=True, size=14)
ws.merge_cells('A1:C1')

# Definiere Stile
red_font = Font(color="FF0000", bold=True)  # Rot für Feiertage
red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")  # Hellrot
header_font = Font(bold=True)
alignment_center = Alignment(horizontal="center", vertical="center")
alignment_left = Alignment(horizontal="left", vertical="center")

current_row = 3

# Für jeden Monat -> 1-basiert, da erster Januar = 1 
for month in range(1, 13):
    
    # Name des Monats
    month_name_str = MONTHS_GER[month]
    ws[f'A{current_row}'] = f"{month_name_str}"
    ws[f'A{current_row}'].font = header_font
    ws[f'A{current_row}'].alignment = alignment_left
    ws.merge_cells(f'A{current_row}:C{current_row}')
    current_row += 1
    
    # Header für Monat
    ws[f'A{current_row}'] = "Tag"
    ws[f'B{current_row}'] = "Wochentag"
    for col in ['A', 'B', 'C']:
        ws[f'{col}{current_row}'].font = header_font
        ws[f'{col}{current_row}'].alignment = alignment_center
    current_row += 1
    
    # Tage des Monats
    num_days = monthrange(YEAR, month)[1]
    for day in range(1, num_days + 1):
        current_date = datetime(YEAR, month, day).date()
        
        # Überspringe Wochenenden
        if current_date.weekday() >= 5:
            continue
        
        weekday = current_date.strftime('%A')
        is_holiday = current_date in hessian_holidays
        
        # Daten schreiben
        ws[f'A{current_row}'] = day
        ws[f'B{current_row}'] = weekday
        
        # Feiertage rot markieren
        if is_holiday:
            for col in ['A', 'B', 'C']:
                #ws[f'{col}{current_row}'].font = red_font
                ws[f'{col}{current_row}'].fill = red_fill
        
        for col in ['A', 'B', 'C']:
            ws[f'{col}{current_row}'].alignment = alignment_center
        
        current_row += 1
    
    current_row += 1  # Leerzeile zwischen Monaten

# Spaltenbreite anpassen
ws.column_dimensions['A'].width = 8
ws.column_dimensions['B'].width = 15
ws.column_dimensions['C'].width = 12

# Datei speichern
try:
    wb.save(EXCEL_NAME)
    print(f"Excel-Datei '{EXCEL_NAME}' wurde erstellt. Feiertage sind rot markiert.")
except:
    print(f"Excel-Datei '{EXCEL_NAME}' ist schon geöffnet.")