import random
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

class ShiftScheduler:
    def __init__(self, start_date, end_date, names):
        self.start_date = start_date
        self.end_date = end_date
        self.names = names
        self.schedule = {}
        self.assignments = {name: {'total': 0, 'weekends': 0, 'dates': [], 'last_assigned': None} for name in names}
        self.weekend_days = self.calculate_weekend_days()
        self.total_shifts = ((self.end_date - self.start_date).days + 1) * 2

    def calculate_weekend_days(self):
        return sum(1 for date in (self.start_date + timedelta(n) for n in range((self.end_date - self.start_date).days + 1))
                   if date.weekday() >= 5)

    def generate_schedule(self):
        total_days = (self.end_date - self.start_date).days + 1
        total_shifts = total_days * 2  # 2 people per day
        shifts_per_person = total_shifts // len(self.names)
        extra_shifts = total_shifts % len(self.names)
        weekend_shifts_per_person = self.weekend_days * 2 // len(self.names)
        extra_weekend_shifts = (self.weekend_days * 2) % len(self.names)
        
        current_date = self.start_date
        while current_date <= self.end_date:
            # Prioritize filling weekend shifts evenly
            is_weekend = current_date.weekday() >= 5
            if is_weekend:
                available_names = [name for name in self.names 
                                   if self.assignments[name]['last_assigned'] != current_date - timedelta(days=1)
                                   and self.assignments[name]['weekends'] < weekend_shifts_per_person + (1 if extra_weekend_shifts > 0 else 0)]
                random.shuffle(available_names)  # Shuffle to ensure randomness
                extra_weekend_shifts -= 2  # Decrement extra weekend shifts as they're assigned
            else:
                available_names = [name for name in self.names 
                                   if self.assignments[name]['last_assigned'] != current_date - timedelta(days=1)]
                random.shuffle(available_names)  # Shuffle to ensure randomness
            
            # Sort by total shifts to ensure balanced distribution
            available_names.sort(key=lambda x: self.assignments[x]['total'])
            shift = []
            for _ in range(2):  # Assign 2 people per day
                if not available_names:
                    available_names = [name for name in self.names 
                                       if self.assignments[name]['last_assigned'] != current_date - timedelta(days=1)]
                    random.shuffle(available_names)
                    available_names.sort(key=lambda x: self.assignments[x]['total'])
                
                # Assign shifts, ensuring minimal disparity
                name = available_names.pop(0)
                shift.append(name)
                self.assignments[name]['total'] += 1
                self.assignments[name]['dates'].append(current_date)  # Track assigned dates
                self.assignments[name]['last_assigned'] = current_date
                
                if is_weekend:  # Weekend shift
                    self.assignments[name]['weekends'] += 1

                # Check to ensure no one exceeds the minimum necessary extra shifts
                if self.assignments[name]['total'] > shifts_per_person + (1 if extra_shifts > 0 else 0):
                    if extra_shifts > 0:
                        extra_shifts -= 1
                    else:
                        # Move this person to the end to prevent further selection
                        available_names.append(name)

            self.schedule[current_date] = shift
            current_date += timedelta(days=1)

    def print_schedule(self):
        print("Schedule by Date:")
        for date, shift in sorted(self.schedule.items()):
            print(f"{date.strftime('%Y-%m-%d')}: {', '.join(shift)}")

    def print_personal_schedules(self):
        print("\nSchedule by Person:")
        for name, stats in self.assignments.items():
            dates = ', '.join(date.strftime('%Y-%m-%d') for date in stats['dates'])
            print(f"{name}: {dates}")

    def print_statistics(self):
        print("\nAssignment Statistics:")
        for name, stats in self.assignments.items():
            print(f"{name}: Total: {stats['total']}, Weekends: {stats['weekends']}")

    def export_to_excel(self, filename):
        wb = Workbook()

        # Shift Schedule sheet
        ws = wb.active
        ws.title = "Shift Schedule"
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        
        # Headers
        ws['A1'] = "Date"
        ws['B1'] = "Person 1"
        ws['C1'] = "Person 2"
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center")
        
        # Data for shift schedule
        for row, (date, shift) in enumerate(sorted(self.schedule.items()), start=2):
            ws.cell(row=row, column=1, value=date.strftime("%d/%m/%Y"))
            for col, name in enumerate(shift, start=2):
                cell = ws.cell(row=row, column=col, value=name)
                cell.border = border
                cell.alignment = Alignment(horizontal="center")
        
        # Adjust column widths for shift schedule
        for column in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column_letter].width = adjusted_width

        # Personal Schedule sheet
        ws_personal = wb.create_sheet(title="Personal Schedules")
        ws_personal['A1'] = "Person"
        ws_personal['B1'] = "Dates"
        ws_personal['A1'].fill = header_fill
        ws_personal['B1'].fill = header_fill
        ws_personal['A1'].font = Font(bold=True, color="FFFFFF")
        ws_personal['B1'].font = Font(bold=True, color="FFFFFF")
        ws_personal['A1'].alignment = Alignment(horizontal="center")
        ws_personal['B1'].alignment = Alignment(horizontal="center")

        # Data for personal schedules
        for row, (name, stats) in enumerate(self.assignments.items(), start=2):
            dates = ', '.join(date.strftime('%d/%m/%Y') for date in stats['dates'])
            ws_personal.cell(row=row, column=1, value=name)
            ws_personal.cell(row=row, column=2, value=dates)
            ws_personal.cell(row=row, column=1).border = border
            ws_personal.cell(row=row, column=2).border = border
            ws_personal.cell(row=row, column=1).alignment = Alignment(horizontal="center")
            ws_personal.cell(row=row, column=2).alignment = Alignment(horizontal="center")

        # Adjust column widths for personal schedules
        for column in ws_personal.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws_personal.column_dimensions[column_letter].width = adjusted_width

        wb.save(filename)
        print(f"Schedule exported to {filename}")

# Example usage
start_date = datetime(2024, 1, 1)
end_date = datetime(2024, 1, 31)
names = ["Shakir", "Fikhry", "Syed", "Munshi", "Aiman", "Luthfi", "Dalvin", "Hazim", "Jerry", "Yassin", "Donavan"]
scheduler = ShiftScheduler(start_date, end_date, names)
scheduler.generate_schedule()
scheduler.print_schedule()
scheduler.print_personal_schedules()
scheduler.print_statistics()
scheduler.export_to_excel("shift_schedule.xlsx")
