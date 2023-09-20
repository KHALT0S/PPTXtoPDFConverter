# Import necessary modules
import sys
import os
import zipfile
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QPushButton,
    QLabel, QCheckBox, QVBoxLayout, QWidget, QTextEdit, QMessageBox, QProgressBar
)
from PyQt5.QtCore import Qt, QTimer
from comtypes import client

# Create a class for the PPTX to PDF converter application
class PPTXtoPDFConverter(QMainWindow):
    def __init__(self):
        super().__init__()

        # Initialize the user interface
        self.initUI()

    def initUI(self):
        # Set the main window's geometry and title
        self.setGeometry(100, 100, 600, 400)
        self.setWindowTitle('PPTX/PPT to PDF Converter')

        # Create a central widget for the main window
        self.centralWidget = QWidget()
        self.setCentralWidget(self.centralWidget)

        # Create a vertical layout for the central widget
        layout = QVBoxLayout()

        # Create a label for file selection
        self.label = QLabel('Select PowerPoint files (PPTX or PPT) to convert to PDF:')
        layout.addWidget(self.label, alignment=Qt.AlignCenter)

        # Create a button to initiate the conversion process
        self.convert_btn = QPushButton('Convert to PDF')
        layout.addWidget(self.convert_btn)

        # Create a progress bar to show the conversion progress
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        # Create a checkbox for toggling the dark theme
        self.theme_switch = QCheckBox('Dark Theme')
        layout.addWidget(self.theme_switch, alignment=Qt.AlignLeft)

        # Create a checkbox for compressing converted files into a zip archive
        self.zip_switch = QCheckBox('Compress to Zip')
        layout.addWidget(self.zip_switch, alignment=Qt.AlignLeft)

        # Create a text box for displaying the converted file names
        self.converted_files = QTextEdit()
        self.converted_files.setReadOnly(True)  # Make it read-only
        layout.addWidget(self.converted_files)

        # Create a bottom bar for version information
        bottom_bar = QWidget()
        bottom_layout = QVBoxLayout()
        version_label = QLabel('v1.0')
        version_label.setAlignment(Qt.AlignCenter)
        bottom_layout.addWidget(version_label)
        bottom_bar.setLayout(bottom_layout)

        # Add the bottom bar to the main layout
        layout.addWidget(bottom_bar)

        # Set the layout for the central widget
        self.centralWidget.setLayout(layout)

        # Set the dark theme as the default
        self.theme_switch.setChecked(True)
        self.theme_switch.stateChanged.connect(self.toggleTheme)

        # Set the zip compression checkbox to unchecked by default
        self.zip_switch.setChecked(False)
        self.zip_switch.stateChanged.connect(self.toggleZip)

        # Apply the dark theme stylesheet initially
        self.setStyleSheet(self.getStylesheet('dark'))

        # Create a timer for the conversion process
        self.conversion_timer = QTimer(self)
        self.conversion_timer.timeout.connect(self.convertNextFile)
        self.conversion_timer.setInterval(1000)

        # Connect the convert button click event to the conversion method
        self.convert_btn.clicked.connect(self.convertToPDF)

        # Initialize variables for zip file handling
        self.zip_filename = None
        self.zip_file = None

        # Initialize a list to store converted PDF file paths
        self.converted_pdf_paths = []

    # Define the dark and light theme stylesheets
    def getStylesheet(self, theme):
        if theme == 'dark':
            return """
                QMainWindow {
                    background-color: #121212;
                    color: #FFFFFF;
                }
                QLabel {
                    color: #FFFFFF;
                }
                QPushButton {
                    background-color: #007BFF;
                    color: #FFFFFF;
                    border: none;
                    border-radius: 5px;
                    padding: 8px 16px;
                }
                QPushButton:hover {
                    background-color: #0056b3;
                }
                QCheckBox {
                    color: #FFFFFF;
                }
                QProgressBar {
                    background-color: #333333;
                    color: #FFFFFF;
                }
                QTextEdit {
                    background-color: #1E1E1E;
                    color: #FFFFFF;
                }
            """
        else:
            return """
                QMainWindow {
                    background-color: #FFFFFF;
                    color: #333333;
                }
                QLabel {
                    color: #333333;
                }
                QPushButton {
                    background-color: #007BFF;
                    color: #FFFFFF;
                    border: none;
                    border-radius: 5px;
                    padding: 8px 16px;
                }
                QPushButton:hover {
                    background-color: #0056b3;
                }
                QCheckBox {
                    color: #333333;
                }
                QProgressBar {
                    background-color: #E1E1E1;
                    color: #333333;
                }
                QTextEdit {
                    background-color: #FFFFFF;
                    color: #333333;
                }
            """

    # Method for toggling the application's theme
    def toggleTheme(self):
        if self.theme_switch.isChecked():
            self.setStyleSheet(self.getStylesheet('dark'))
        else:
            self.setStyleSheet(self.getStylesheet('light'))

    def toggleZip(self):
        if self.zip_switch.isChecked():
            # Do not open the file dialog here
            # The user will choose the location when pressing the "Convert" button
            self.zip_filename = None
            self.zip_file = None
        else:
            if self.zip_file:
                self.zip_file.close()
                self.zip_file = None
            self.zip_filename = None

    def convertToPDF(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        options |= QFileDialog.ExistingFiles

        ppt_files, _ = QFileDialog.getOpenFileNames(
            self, 'Open PowerPoint Files', '', 'PowerPoint Files (*.pptx *.ppt)', options=options
        )

        if ppt_files:
            self.ppt_files = ppt_files
            self.progress_bar.setMaximum(len(self.ppt_files))
            self.progress_bar.setValue(0)

            # Check if the "Compress to Zip" option is enabled
            if self.zip_switch.isChecked():
                # Open the file dialog to choose the zip file location
                self.zip_filename, _ = QFileDialog.getSaveFileName(
                    self, 'Save Zip File', '', 'Zip Files (*.zip)'
                )
                if not self.zip_filename:
                    # User canceled, so uncheck the checkbox and return
                    self.zip_switch.setChecked(False)
                    return
                # Create the zip file
                self.zip_file = zipfile.ZipFile(self.zip_filename, 'w', zipfile.ZIP_DEFLATED)

            self.conversion_timer.start()

    # Method for confirming file overwrite during conversion
    def confirmOverwrite(self, pdf_file):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setText(f"A PDF file with the name '{pdf_file}' already exists. Do you want to overwrite it?")
        msg.setWindowTitle("Overwrite Confirmation")
        msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        result = msg.exec_()
        return result == QMessageBox.Yes

    # Method for converting the next PowerPoint file to PDF
    def convertNextFile(self):
        if self.ppt_files:
            ppt_file = self.ppt_files.pop(0)
            ppt_file = os.path.abspath(ppt_file)
            pdf_file = os.path.splitext(ppt_file)[0] + '.pdf'

            if os.path.exists(pdf_file):
                if not self.confirmOverwrite(pdf_file):
                    self.converted_files.append(f'Skipped: {os.path.basename(ppt_file)}')
                    self.progress_bar.setValue(self.progress_bar.value() + 1)
                    return

            ppt_app = client.CreateObject("PowerPoint.Application")
            ppt = ppt_app.Presentations.Open(ppt_file)

            # Minimize the PowerPoint window
            ppt_app.WindowState = 2  # 2 represents minimized

            ppt.SaveAs(pdf_file, 32)
            ppt_app.Quit()

            self.converted_files.append(f'Converted: {os.path.basename(ppt_file)}')
            self.progress_bar.setValue(self.progress_bar.value() + 1)

            self.converted_pdf_paths.append(pdf_file)

            if self.zip_file:
                self.zip_file.write(pdf_file, os.path.basename(pdf_file))

        if not self.ppt_files:
            self.conversion_timer.stop()
            self.converted_files.append('Conversion complete.')

            if self.zip_file:
                self.zip_file.close()
                self.converted_files.append(f'PDFs compressed to {self.zip_filename}')

if __name__ == '__main__':
    # Create and run the PyQt5 application
    app = QApplication(sys.argv)
    window = PPTXtoPDFConverter()
    window.show()
    sys.exit(app.exec_())