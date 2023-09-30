import pandas as pd
from datetime import datetime
from openpyxl import Workbook

class EmployeeXLSXGenerator:

    def __init__(self, csv_file_name, xlsx_file_name):
        self.csv_file_name = csv_file_name
        self.xlsx_file_name = xlsx_file_name

    def calculate_age(self, dob):
        """Calculate the age based on the date of birth."""
        today = datetime.today()
        age = today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))
        return age

    def generate_xlsx(self):
        try:
            # Load CSV data into a pandas DataFrame
            data = pd.read_csv(self.csv_file_name, encoding='utf-8-sig', sep=';')

            # Calculate age
            data['Дата народження'] = pd.to_datetime(data['Дата народження'], format='%Y-%m-%d')
            data['Вік'] = data['Дата народження'].apply(self.calculate_age)

            # Create a new Excel workbook and add sheets
            workbook = Workbook()

            sheets = ["all", "younger_18", "18-45", "45-70", "older_70"]

            for sheet_name in sheets:
                sheet = workbook.create_sheet(title=sheet_name)
                if sheet_name == "all":
                    header = data.columns.tolist()  # Отримати всі стовпці для аркушу "all"
                else:
                    header = ["Прізвище", "Ім’я", "По батькові", "Дата народження", "Вік"]

                sheet.append(header)

                filtered_data = data.copy()

                if sheet_name != "all":
                    if sheet_name == "younger_18":
                        filtered_data = filtered_data[filtered_data['Вік'] < 18]
                    elif sheet_name == "18-45":
                        filtered_data = filtered_data[(filtered_data['Вік'] >= 18) & (filtered_data['Вік'] <= 45)]
                    elif sheet_name == "45-70":
                        filtered_data = filtered_data[(filtered_data['Вік'] > 45) & (filtered_data['Вік'] <= 70)]
                    elif sheet_name == "older_70":
                        filtered_data = filtered_data[filtered_data['Вік'] > 70]


                    for column in filtered_data.columns:
                        if column not in header:
                            filtered_data = filtered_data.drop(columns=[column])

                for _, row in filtered_data.iterrows():
                    sheet.append(row.tolist())

            # Remove the default sheet created by openpyxl
            default_sheet = workbook["Sheet"]
            workbook.remove(default_sheet)

            # Save the Excel file
            workbook.save(self.xlsx_file_name)
            print("Ok")

        except pd.errors.ParserError:
            print(f"Error: Could not parse the CSV file: {self.csv_file_name}")
        except PermissionError:
            print(f"Error: Could not create the XLSX file: {self.xlsx_file_name} (Permission denied)")
        except Exception as e:
            print(f"Error: {e}")

# Usage:
xlsx_generator = EmployeeXLSXGenerator(
    csv_file_name="employees.csv",
    xlsx_file_name="employees.xlsx"
)
xlsx_generator.generate_xlsx()
