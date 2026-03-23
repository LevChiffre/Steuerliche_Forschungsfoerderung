import pandas as pd
import holidays
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from calendar import monthrange, month_name

class cEmployees:
    def __init__(self, lastName:str, name:str, salary_month:int, project:str, workload:int=100):
        self.lastName = lastName
        self.name = name
        self.salary_month = salary_month
        self.project = project


#Abteilungsleiter 50%, Bereichleiter 50%, Sven 60%, Lev 100%
employee_1 = cEmployees("Mustermann", "Max", 60000/12, "A", 50)
employee_2 = cEmployees("Zufall", "Reiner", 50000/12, "A")
employee_3 = cEmployees("Maduschen", "Isolde", 150000/12, "B",60)
employee_4 = cEmployees("Silie", "Peter", 50000/12, "B")



# Globale Einstellungen
MONTHS_GER = ["dummy","Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Detember"]
DAY_SHORT_GER = ["Mo", "Di", "Mi", "Do", "Fr"]
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
    current_row += 1
    
    # Sammle Arbeitstage für diesen Monat
    workdays = []
    num_days = monthrange(YEAR, month)[1]
    for day in range(1, num_days + 1):
        current_date = datetime(YEAR, month, day).date()
        
        day_no = current_date.weekday()
        # Überspringe Wochenenden
        if day_no >= 5:
            continue
        weekday = DAY_SHORT_GER[day_no]
        is_holiday = current_date in hessian_holidays
        workdays.append({'day': day, 'weekday': weekday, 'is_holiday': is_holiday})
    
    # Tag-Reihe (horizontal)
    ws[f'A{current_row}'] = "Tag"
    ws[f'A{current_row}'].font = header_font
    ws[f'A{current_row}'].alignment = alignment_center
    for idx, day_info in enumerate(workdays):
        col = 2 + idx  # Startkolumne B (2)
        day_cell = ws.cell(row=current_row, column=col, value=day_info['day'])
        day_cell.font = header_font
        day_cell.alignment = alignment_center
        if day_info['is_holiday']:
            day_cell.fill = red_fill
    current_row += 1
    
    # Wochentag-Reihe (horizontal)
    ws[f'A{current_row}'] = "Wochentag"
    ws[f'A{current_row}'].font = header_font
    ws[f'A{current_row}'].alignment = alignment_center
    for idx, day_info in enumerate(workdays):
        col = 2 + idx
        weekday_cell = ws.cell(row=current_row, column=col, value=day_info['weekday'])
        weekday_cell.alignment = alignment_center
        if day_info['is_holiday']:
            weekday_cell.fill = red_fill
    
    current_row += 5  # Leerzeile zwischen Monaten

# Spaltenbreite anpassen
ws.column_dimensions['A'].width = 12
for col in range(2, 35):  # Genug Spalten für alle Tage im Monat
    col_letter = get_column_letter(col)
    ws.column_dimensions[col_letter].width = 4

# Datei speichern
try:
    wb.save(EXCEL_NAME)
    print(f"Excel-Datei '{EXCEL_NAME}' wurde erstellt. Feiertage sind rot markiert.")
except:
    print(f"Excel-Datei '{EXCEL_NAME}' ist schon geöffnet.")