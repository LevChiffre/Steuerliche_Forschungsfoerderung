import pandas as pd
import holidays
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from calendar import monthrange, month_name

AVERAGE_WEEKS_PER_MONTH = 4.33


class cProject:
    def __init__(self, name, budget):
        self.name = name
        self.budget = budget

class cWeeklyWorkingTime:
    def __init__(self, monday:int=8.5, tuesday:int=8.5, wednesday:int=8.5, thursday:int=8.5, friday:int=6):
        self.monday = monday
        self.tuesday = tuesday
        self.wednesday =wednesday
        self.thursday = thursday
        self.friday = friday
        self.weekly_working_time = monday + tuesday + wednesday + thursday + friday

class cEmployees:
    def __init__(self, lastName:str, name:str, salary_year:int, project:cProject, workload:int=100, weekly_working_time:cWeeklyWorkingTime=cWeeklyWorkingTime()):
        self.lastName = lastName
        self.name = name
        self.salary_hour = salary_year / (12 * weekly_working_time.weekly_working_time * AVERAGE_WEEKS_PER_MONTH)
        self.project = project
        self.workload = workload
        self.weekly_working_time = weekly_working_time


project_1 = cProject("Project A", 200000)
project_2 = cProject("Project B", 300000)


#Abteilungsleiter 50%, Bereichleiter 50%, Sven 60%, Lev 100%
employee_1 = cEmployees("Mustermann", "Max", 60000, project_1, 50, cWeeklyWorkingTime())
employee_2 = cEmployees("Zufall", "Reiner", 50000, project_1,cWeeklyWorkingTime())
employee_3 = cEmployees("Maduschen", "Isolde", 150000, project_2,60, cWeeklyWorkingTime())
employee_4 = cEmployees("Silie", "Peter", 50000, project_2, cWeeklyWorkingTime())

EMPLOYEES = [employee_1, employee_2, employee_3, employee_4]

# Arbeitsstunden
#NORMAL_HOURS = 8.5
#FRIDAY_HOURS = 6.0

# Funktion zum Parsen der Fehltage
def parse_absences(file_path):
    df = pd.read_excel(file_path, header=None)
    absences = {}
    for emp in EMPLOYEES:
        emp_name = f"{emp.lastName}, {emp.name}"
        absences[emp_name] = {}
    
    current_employee = None
    current_month = None
    
    for idx, row in df.iterrows():
        cell_0 = str(row[0]).strip() if pd.notna(row[0]) else ""
        
        # Mitarbeiter finden
        for emp in EMPLOYEES:
            emp_name = f"{emp.lastName}, {emp.name}"
            if emp_name in cell_0:
                current_employee = emp_name
                current_month = None
                break
        
        if current_employee is None:
            continue
        
        # Monat finden
        if cell_0 in ['Jan', 'Feb', 'Mär', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez']:
            month_map = {'Jan':1, 'Feb':2, 'Mär':3, 'Apr':4, 'Mai':5, 'Jun':6, 'Jul':7, 'Aug':8, 'Sep':9, 'Okt':10, 'Nov':11, 'Dez':12}
            current_month = month_map[cell_0]
            continue
        
        # Datum mit Wert
        if isinstance(cell_0, str) and cell_0.startswith('202'):
            try:
                date = pd.to_datetime(cell_0).date()
                # Werte in Spalten 1-4 (UT, KT, UT2, KiKr)
                for col in range(1, 5):
                    if pd.notna(row[col]) and row[col] != 0:
                        reason = df.iloc[3, col]  # Header aus Zeile 3
                        absences[current_employee][date] = reason
            except:
                continue
    
    return absences



# Globale Einstellungen
MONTHS_GER = ["dummy","Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"]
DAY_SHORT_GER = ["Mo", "Di", "Mi", "Do", "Fr"]
YEAR = 2023
EXCEL_NAME = "Stundennachweis_" + str(YEAR) + ".xlsx" #Excel wird in das gleiche Verzeichnis wie dieses Skript gespeichert
FEHLTAGE_FILE = "Fehltage DIM 23-24.xlsx"

# Feiertage für Hessen laden
hessian_holidays = holidays.DE(prov='HE', years=YEAR)

# Fehltage laden
absences = parse_absences(FEHLTAGE_FILE)

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
        workdays.append({'day': day, 'weekday': weekday, 'date': current_date, 'is_holiday': is_holiday})
    
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
    current_row += 1
    
    # Für jeden Mitarbeiter
    for emp in EMPLOYEES:
        emp_name = f"{emp.lastName}, {emp.name}"
        ws[f'A{current_row}'] = emp_name
        ws[f'A{current_row}'].font = header_font
        ws[f'A{current_row}'].alignment = alignment_left
        
        for idx, day_info in enumerate(workdays):
            col = 2 + idx
            date = day_info['date']
            is_friday = date.weekday() == 4
            
            if date in absences.get(emp_name, {}):
                # Fehltag
                reason = absences[emp_name][date]
                cell = ws.cell(row=current_row, column=col, value=reason)
                cell.fill = red_fill
            else:
                # Arbeitsstunden mit workload
                base_hours = FRIDAY_HOURS if is_friday else NORMAL_HOURS
                hours = base_hours * (emp.workload / 100.0)
                # Auf halbe Stunden runden
                hours = round(hours * 2) / 2
                cell = ws.cell(row=current_row, column=col, value=hours)
            
            cell.alignment = alignment_center
        
        current_row += 1
    
    current_row += 2  # Leerzeile zwischen Monaten

# Spaltenbreite anpassen
ws.column_dimensions['A'].width = 25
for col in range(2, 35):  # Genug Spalten für alle Tage im Monat
    col_letter = get_column_letter(col)
    ws.column_dimensions[col_letter].width = 6

# Datei speichern
try:
    wb.save(EXCEL_NAME)
    print(f"Excel-Datei '{EXCEL_NAME}' wurde erstellt. Feiertage sind rot markiert.")
except:
    print(f"Excel-Datei '{EXCEL_NAME}' ist schon geöffnet.")