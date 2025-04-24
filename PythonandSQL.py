import sys
import psycopg2
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel,
    QLineEdit, QPushButton, QMessageBox
)

# Replace with your actual DB connection details
DB_HOST = "localhost"
DB_PORT = "5432"
DB_NAME = "pyex"
DB_USER = "postgres"
DB_PASSWORD = "kingsman"

class DataEntryApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Add Person to PostgreSQL")
        self.setGeometry(100, 100, 300, 150)

        layout = QVBoxLayout()

        self.name_input = QLineEdit(self)
        self.name_input.setPlaceholderText("Enter name")
        layout.addWidget(QLabel("Name:"))
        layout.addWidget(self.name_input)

        self.age_input = QLineEdit(self)
        self.age_input.setPlaceholderText("Enter age")
        layout.addWidget(QLabel("Age:"))
        layout.addWidget(self.age_input)

        self.submit_button = QPushButton("Add to Database", self)
        self.submit_button.clicked.connect(self.add_to_database)
        layout.addWidget(self.submit_button)

        self.setLayout(layout)

    def add_to_database(self):
        name = self.name_input.text()
        age_text = self.age_input.text()

        if not name or not age_text.isdigit():
            QMessageBox.warning(self, "Input Error", "Please enter a valid name and age.")
            return

        age = int(age_text)

        try:
            conn = psycopg2.connect(
                host=DB_HOST,
                port=DB_PORT,
                dbname=DB_NAME,
                user=DB_USER,
                password=DB_PASSWORD
            )
            cur = conn.cursor()
            cur.execute('INSERT INTO "PERSON" (Name, Age) VALUES (%s, %s)', (name, age))

            conn.commit()
            cur.close()
            conn.close()

            QMessageBox.information(self, "Success", "Data added successfully!")
            self.name_input.clear()
            self.age_input.clear()
        except Exception as e:
            QMessageBox.critical(self, "Database Error", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DataEntryApp()
    window.show()
    sys.exit(app.exec())
