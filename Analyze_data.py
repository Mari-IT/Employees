import pandas as pd
import os
import matplotlib.pyplot as plt
from datetime import datetime

class EmployeesDataAnalyzer:
    def __init__(self, csv_file_path):
        self.csv_file_path = csv_file_path
        self.data = None

    def load_data(self):
        if not os.path.exists(self.csv_file_path):
            print(f"Error: Could not find or open the CSV file: {self.csv_file_path}")
            return False

        try:
            self.data = pd.read_csv(self.csv_file_path, encoding='utf-8-sig', sep=';', parse_dates=['Дата народження'])
            print("Data loaded successfully.")
            return True
        except pd.errors.ParserError:
            print(f"Error: Could not parse the CSV file: {self.csv_file_path}")
            return False

    def calculate_age(self, dob):
        today = datetime.today()
        age = today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))
        return age

    def analyze_gender(self):
        print("\nTask: Analyze Gender Distribution\n" + "-"*40)
        if self.data is not None:
            gender_counts = self.data['Стать'].value_counts()
            print(gender_counts)
            plt.figure()  # Створення нової фігури
            gender_counts.plot(kind='bar')
            plt.title('Distribution by Gender')
            plt.show()
        else:
            print("Error: Data not loaded.")

    def analyze_age_groups(self):
        print("\nTask: Analyze Age Group Distribution\n" + "-"*40)
        if self.data is not None:
            self.data['Вік'] = self.data['Дата народження'].apply(self.calculate_age)
            age_bins = [0, 18, 45, 70, 200]
            age_labels = ['younger_18', '18-45', '45-70', 'older_70']
            self.data['Вікова група'] = pd.cut(self.data['Вік'], bins=age_bins, labels=age_labels, right=False)
            age_group_counts = self.data['Вікова група'].value_counts()
            print(age_group_counts)
            plt.figure()  # Створення нової фігури
            age_group_counts.plot(kind='bar')
            plt.title('Distribution by Age Groups')
            plt.show()
        else:
            print("Error: Data not loaded.")

    def analyze_gender_age_groups(self):
        print("\nTask: Analyze Gender and Age Group Distribution\n" + "-"*40)
        if self.data is not None:
            gender_age_groups_counts = self.data.groupby(['Стать', 'Вікова група'], observed=True).size().unstack()
            print(gender_age_groups_counts)
            plt.figure()  # Створення нової фігури
            gender_age_groups_counts.plot(kind='bar', stacked=True)
            plt.title('Distribution by Gender and Age Groups')
            plt.show()
        else:
            print("Error: Data not loaded.")

    def count_gender(self):
        print("\nTask: Count Gender Distribution\n" + "-"*40)
        if self.data is not None:
            gender_counts = self.data['Стать'].value_counts()
            print(gender_counts)
            plt.figure()  # Створення нової фігури
            gender_counts.plot(kind='bar')
            plt.title('Gender Distribution')
            plt.show()
        else:
            print("Error: Data not loaded.")

    def count_age_groups(self):
        print("\nTask: Count Age Group Distribution\n" + "-"*40)
        if self.data is not None:
            self.data['Вік'] = self.data['Дата народження'].apply(self.calculate_age)
            age_bins = [0, 18, 45, 70, 200]
            age_labels = ['younger_18', '18-45', '45-70', 'older_70']
            self.data['Вікова група'] = pd.cut(self.data['Вік'], bins=age_bins, labels=age_labels, right=False)
            age_group_counts = self.data['Вікова група'].value_counts()
            print(age_group_counts)
            plt.figure()  # Створення нової фігури
            age_group_counts.plot(kind='bar')
            plt.title('Age Group Distribution')
            plt.show()
        else:
            print("Error: Data not loaded.")

    def count_gender_age_groups(self):
        print("\nTask: Count Gender and Age Group Distribution\n" + "-"*40)
        if self.data is not None:
            gender_age_groups_counts = self.data.groupby(['Стать', 'Вікова група'], observed=True).size().unstack()
            print(gender_age_groups_counts)
            plt.figure()  # Створення нової фігури
            gender_age_groups_counts.plot(kind='bar', stacked=True)
            plt.title('Gender and Age Group Distribution')
            plt.show()
        else:
            print("Error: Data not loaded.")

# Використання:
analyzer = EmployeesDataAnalyzer("employees.csv")
if analyzer.load_data():
    analyzer.analyze_gender()
    analyzer.analyze_age_groups()
    analyzer.analyze_gender_age_groups()
    analyzer.count_gender()
    analyzer.count_age_groups()
    analyzer.count_gender_age_groups()
