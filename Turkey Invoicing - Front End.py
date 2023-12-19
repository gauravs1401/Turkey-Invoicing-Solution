import sys
import subprocess
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QLabel,
    QPushButton, QVBoxLayout, QWidget, QGridLayout
)
from PyQt5.QtGui import QPixmap, QResizeEvent, QFont
from PyQt5.QtCore import Qt


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Hewlett Packard Enterprise")
        self.setWindowState(Qt.WindowMaximized)

        # Set background image
        self.set_background_image("HPE_Wallpapers_2022_4K_3840x2160px_01.jpg")

        # Create a central widget and set the layout
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)

        # Add header label
        header_label = QLabel("TURKEY INVOICING AUTOMATION", self)
        header_label.setAlignment(Qt.AlignCenter)
        header_font = QFont("Arial", 40, QFont.Bold)
        header_label.setFont(header_font)
        header_label.setStyleSheet("color: white;")
        self.main_layout.addWidget(header_label)

        # Create a grid layout
        self.grid_layout = QGridLayout()
        self.grid_layout.setAlignment(Qt.AlignCenter)
        self.main_layout.addLayout(self.grid_layout)  # Add the grid layout to the main layout

        # Add button
        self.add_button("Execute Excel", "Turkey Invoicing - Excel.py")
        self.add_button("Execute S4", "Turkey Invoicing - S4.py")

        # Set column stretch to 0 to remove space between buttons
        self.grid_layout.setColumnStretch(0, 0)
        self.grid_layout.setColumnStretch(1, 0)

    def set_background_image(self, image_path):
        self.background_label = QLabel(self)
        self.background_label.setPixmap(QPixmap(image_path))
        self.background_label.setScaledContents(True)

    def add_button(self, button_text, python_file):
        button = QPushButton(button_text, self)
        button.setStyleSheet(
            "QPushButton {"
            "background-color: black;"
            "color: white;"
            "font-size: 40px;"
            "font-family: Arial;"
            "border-radius: 40px;"  # Change the border-radius to modify the shape
            "}"
            "QPushButton:hover {"
            "background-color: darkgray;"  # Change the background color on hover
            "}"
            "QPushButton:pressed {"
            "background-color: gray;"  # Change the background color when pressed
            "}"
        )

        # Connect the button's clicked signal to the run_python_file slot
        button.clicked.connect(lambda: self.run_python_file(python_file))

        # Add the button to the grid layout
        row = self.grid_layout.rowCount()
        self.grid_layout.addWidget(button, row, 0, alignment=Qt.AlignCenter)

    def resizeEvent(self, event: QResizeEvent):
        super().resizeEvent(event)

        # Resize the background label to fit the maximized window
        self.background_label.setGeometry(0, 0, self.width(), self.height())

    def run_python_file(self, python_file):
        command = ["python", python_file]
        subprocess.run(command)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
