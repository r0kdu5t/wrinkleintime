import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, 
    QComboBox, QPushButton, QVBoxLayout, QHBoxLayout, QMessageBox
)
from PyQt6.QtCore import Qt
import openpyxl

class ExcelDataEntryApp(QWidget):
    def __init__(self):
        super().__init__()
        self.excel_filename = "Data_Entry_Records.xlsx"
        self.init_excel()
        self.init_ui()

    def init_excel(self):
        """Creates the Excel file with headers if it doesn't exist."""
        if not os.path.exists(self.excel_filename):
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Form Data"
            # Define your column headers
            sheet.append(["Name", "Age", "Role"])
            workbook.save(self.excel_filename)

    def init_ui(self):
        """Sets up the GUI layout and widgets."""
        self.setWindowTitle("Excel Data Entry Form")
        self.setMinimumWidth(350)

        # Layouts
        main_layout = QVBoxLayout()
        
        # Form Elements
        # 1. Name Input
        self.label_name = QLabel("Full Name:")
        self.input_name = QLineEdit()
        self.input_name.setPlaceholderText("Enter employee name")

        # 2. Age Input
        self.label_age = QLabel("Age:")
        self.input_age = QLineEdit()
        self.input_age.setPlaceholderText("Enter age")

        # 3. Role Dropdown
        self.label_role = QLabel("Job Role:")
        self.input_role = QComboBox()
        self.input_role.addItems(["Developer", "Designer", "Data Analyst", "Manager", "HR"])

        # 4. Submit Button
        self.btn_submit = QPushButton("Submit to Excel")
        self.btn_submit.setStyleSheet("background-color: #2ca02c; color: white; font-weight: bold; padding: 6px;")
        self.btn_submit.clicked.connect(self.submit_data)

        # Adding widgets to the main layout
        main_layout.addWidget(self.label_name)
        main_layout.addWidget(self.input_name)
        
        main_layout.addWidget(self.label_age)
        main_layout.addWidget(self.input_age)
        
        main_layout.addWidget(self.label_role)
        main_layout.addWidget(self.input_role)
        
        main_layout.addSpacing(15)
        main_layout.addWidget(self.btn_submit)

        self.setLayout(main_layout)

    def submit_data(self):
        """Extracts text inputs and appends them to the Excel file."""
        name = self.input_name.text().strip()
        age = self.input_age.text().strip()
        role = self.input_role.currentText()

        # Simple Validation
        if not name or not age:
            QMessageBox.warning(self, "Input Error", "Please fill out all fields before submitting!")
            return

        if not age.isdigit():
            QMessageBox.warning(self, "Input Error", "Age must be a valid number!")
            return

        try:
            # Open existing Excel file and append row
            workbook = openpyxl.load_workbook(self.excel_filename)
            sheet = workbook.active
            sheet.append([name, int(age), role])
            workbook.save(self.excel_filename)

            # Success Feedback
            QMessageBox.information(self, "Success", f"Data for '{name}' successfully saved!")
            
            # Clear fields for the next entry
            self.clear_fields()

        except PermissionError:
            QMessageBox.critical(self, "File Error", f"Could not write to file! Close '{self.excel_filename}' if it is open in Excel and try again.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An unexpected error occurred: {str(e)}")

    def clear_fields(self):
        """Resets the form fields to blank/defaults."""
        self.input_name.clear()
        self.input_age.clear()
        self.input_role.setCurrentIndex(0)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelDataEntryApp()
    window.show()
    sys.exit(app.exec())
