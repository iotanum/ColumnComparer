import sys
import os
from os.path import expanduser
from pathlib import Path
from excel_parse import Parsing
from PyQt5.QtWidgets import (QWidget, QToolTip, QPushButton, QApplication, QMessageBox, QDesktopWidget,
                             QFileDialog, QLabel, QVBoxLayout, QHBoxLayout, QPlainTextEdit, QLineEdit)
from PyQt5.QtGui import QFont
import traceback
from concurrent.futures import *
import time


class Example(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()
        self.file_path = None
        self.main_file = None
        self.folder_path = None
        self.results_f_path = None
        self.main_text_input = None
        self.secondary_text_input = None
        self.main_file_column_status = None
        self.if_column_found_checks = []

    def initUI(self):
        self.text = ""
        QToolTip.setFont(QFont('SansSerif', 10))
        self.setWindowTitle('Paieška')
        self.setGeometry(450, 450, 450, 350)
        self.setFixedSize(self.size())
        self.hbox = QHBoxLayout()
        self.main_file_text_input_box()
        self.secondary_files_text_input_box()
        self.hbox.addStretch(1)
        self.file_button()
        self.folder_button()
        self.where_to_put_button()
        self.hbox.addWidget(self.exit_button())
        self.hbox.addWidget(self.reset_button())
        self.hbox.addWidget(self.start_button())
        self.hbox.addWidget(self.export_button())

        vbox = QVBoxLayout()
        vbox.addStretch(1)
        vbox.addLayout(self.hbox)
        self.text_field()
        self.setLayout(vbox)
        self.show()

    def start_button(self):
        self.start_btn = QPushButton("Pradėti paiešką")
        self.start_btn.clicked.connect(self.start_read)
        self.start_btn.setEnabled(False)
        return self.start_btn

    def exit_button(self):
        cancel_button = QPushButton("Atšaukti/Išeiti")
        cancel_button.clicked.connect(QApplication.instance().closeAllWindows)
        return cancel_button

    def reset_button(self):
        reset_button = QPushButton("Iš naujo")
        reset_button.clicked.connect(self.reset_UI)
        return reset_button

    def export_button(self):
        self.export_btn = QPushButton("Ieškoti atitikimų ir eksportuoti")
        self.export_btn.hide()
        self.export_btn.clicked.connect(self.export_start)
        return self.export_btn

    def reset_UI(self):
        self.file_btn.setText("Pasirinkti")
        self.folder_btn.setText("Pasirinkti")
        self.where_to_btn.setText("Pasirinkti")
        self.file_btn.setEnabled(True)
        self.folder_btn.setEnabled(False)
        self.where_to_btn.setEnabled(False)
        self.box.setPlainText("")
        self.text_input_one.setText("")
        self.text_input_two.setText("")
        self.text_input_one.setEnabled(False)
        self.text_input_two.setEnabled(False)
        self.file_path = None
        self.main_file = None
        self.folder_path = None
        self.results_f_path = None
        self.main_text_input = None
        self.secondary_text_input = None
        self.main_file_column_status = None
        self.if_column_found_checks = []
        self.export_btn.hide()
        self.start_btn.show()
        self.start_btn.setEnabled(True if self.start_btn.setEnabled(False) else False)
        self.export_btn.setEnabled(True)
        self.export_btn.setText("Ieškoti atitikimų ir eksportuoti")
        Parsing.__init__()

    def export_start(self):
        try:
            Parsing.make_excel(self.results_f_path)
        except:
            print(traceback.format_exc())

        if Parsing.found_status:
            self.box.appendPlainText(f"Sėkmingai eksportuoti rezultatai, "
                                     f"dokumento pavadinimas -\n"
                                     f"'{Parsing.result_file_name}',"
                                     f" aplanke '{os.path.basename(self.results_f_path)}'.")
            self.export_btn.setText("Eksportuota")
        else:
            self.box.appendPlainText(f"Nei vieno įrašo nerasta.")
            self.export_btn.setText("Nerasta")
        self.export_btn.setEnabled(False)
        self.box.verticalScrollBar().setValue(self.box.verticalScrollBar().maximum())

    def start_read(self):
        if self.checks():
            file_paths = Parsing.get_excel_list(self.folder_path)
            if "~$" not in str(file_paths):
                if "" is not self.text_input_one.text() and "" is not self.text_input_two.text():
                    self.find_and_parse_main_excel_file()
                    self.find_and_parse_secondary_excel_files(file_paths)
                    self.start_btn.hide()
                    self.export_btn.show()
                    self.hbox.update()
                else:
                    QMessageBox.about(self, "Kažkas netaip", f"Pateikite stulpelio pavadinimą!")
            else:
                QMessageBox.about(self, "Kažkas netaip", f"Kažkuris dokumentas yra atidarytas!")

    def checks(self):
        checks = [self.file_path, self.folder_path, self.results_f_path]
        if not all(checks):
            QMessageBox.about(self, "Kažkas netaip", f"Visi laukai turi būti parinkti!")
        else:
            return True

    def executor_and_columns(self, directories):
        try:
            start = time.time()
            with ThreadPoolExecutor(max_workers=5) as p:
                a = {p.submit(self.find_secondary_file_columns, directory): directory for directory in directories}
                for future in as_completed(a):
                    self.column_name_list.append(future.result())
            end = time.time()
            print(f"Search and parse took: {round(end - start, 2)}s")

        except:
            print(traceback.format_exc())

        finally:
            self.print_found()

    def find_and_parse_secondary_excel_files(self, directories):
        if directories:
            self.input_column_name = self.input_column_names_parse(self.text_input_two.text())
            self.box.appendPlainText(f"Aplanke '{os.path.basename(self.folder_path)}' rasti šie dokumentai"
                                     f" ir jų stulpeliai:")
            self.executor_and_columns(directories)
            if self.main_file_column_status:
                self.box.appendPlainText("Jeigu viskas teisingai - spauskite "
                                         "'Ieškoti atitikimų ir eksportuoti' mygtuką.")
                if not all(self.if_column_found_checks):
                    column_not_found_count = self.if_column_found_checks.count(False)
                    self.box.appendPlainText(f"{column_not_found_count} dokument"
                                             f"{'ai' if column_not_found_count > 1 else 'as'} "
                                             f"neturėjo ieškomo stulpelio, "
                                             f"prašoma dar kartą patikrinti ar viską teisingai suvedėte.")
            else:
                self.box.appendPlainText("Klaida!\nPatikrinkite ar gerai įvedėte stuleplių pavadinimus"
                                         "/ar gerai parinkote reikalingą pagrindinį failą.\n"
                                         "Bandykite viską iš naujo.")
        else:
            self.box.appendPlainText(f"Aplanke '{os.path.basename(self.folder_path)}' nerasta excel failų.")
            self.export_btn.setEnabled(False)

    def print_found(self):
        for idx, (file_name, column_info) in enumerate(self.column_name_list):
            try:
                column_info = ", ".join(column_info)
            except TypeError:
                column_info = "    Reikalingas stulpelis nerastas."
            self.box.appendPlainText(f"{idx}. {file_name}:\n    Stulpelis {column_info}.\n")

    def find_secondary_file_columns(self, directory):
        if str(self.main_file) not in str(directory) and "rezultatai" not in str(directory):
            file = str(directory).split("\\")[-1]

            try:
                column_abc, column_name = Parsing.find_needed_column(Path(directory), self.input_column_name)
            except TypeError:
                column_abc = None

            if column_abc is not None:
                self.if_column_found_checks.append(True)
                return file, [column_abc, column_name]
            else:
                self.if_column_found_checks.append(False)
                return file, None

    def find_and_parse_main_excel_file(self):
        input_column_name = self.input_column_names_parse(self.text_input_one.text())
        self.box.appendPlainText(f"\nFaile '{self.main_file}' rastas šis stulpelis:")
        text = Parsing.find_needed_column(Path(self.file_path), input_column_name, main=True)
        if text is not None:
            self.box.appendPlainText(f"1. {text[6:]}")
            self.main_file_column_status = True
        else:
            self.box.appendPlainText(f"1. Reikalingas stulpelis nerastas.\n")
            self.main_file_column_status = False
            self.export_btn.setEnabled(False)

    def input_column_names_parse(self, text):
        return text.split(" ")

    def main_file_text_input_box(self):
        self.main_text_input_label()
        self.text_input_one = QLineEdit(self)
        self.text_input_one.move(265, 35)
        self.text_input_one.resize(150, 20)
        self.text_input_one.setEnabled(False)

    def main_text_input_label(self):
        lbl1 = QLabel('Failo, pagal kurį ieškoma\nstulpelio pavadinimas:', self)
        lbl1.move(280, 5)
        lbl1.resize(150, 30)

    def secondary_files_text_input_box(self):
        self.secondary_text_input_label()
        self.text_input_two = QLineEdit(self)
        self.text_input_two.move(265, 85)
        self.text_input_two.resize(150, 20)
        self.text_input_two.setEnabled(False)

    def secondary_text_input_label(self):
        lbl1 = QLabel('Failo, kuriame ieškoma\nstulpelio pavadinimas:', self)
        lbl1.move(280, 55)
        lbl1.resize(150, 30)

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.TopLeft())

    def text_field(self):
        self.box = QPlainTextEdit(self)
        self.box.insertPlainText(self.text)
        self.box.setReadOnly(True)
        self.box.move(10, 115)
        self.box.resize(430, 195)

    def file_button(self):
        lbl1 = QLabel('<b>1.</b>', self)
        lbl1.move(10, 10)
        lbl2 = QLabel("Pasirinkite failą:", self)
        lbl2.move(25, 9)
        self.file_btn = QPushButton('Pasirinkti', self)
        self.file_btn.setToolTip('Pasirinkite failą iš kurio bus ieškoma.')
        self.file_btn.clicked.connect(self.select_file)
        self.file_btn.resize(self.file_btn.sizeHint())
        self.file_btn.move(150, 10)

    def folder_button(self):
        lbl1 = QLabel('<b>2.</b>', self)
        lbl1.move(10, 40)
        lbl2 = QLabel('Pasirinkite aplanką:', self)
        lbl2.move(25, 40)
        self.folder_btn = QPushButton('Pasirinkti', self)
        self.folder_btn.setToolTip('Pasirinkite aplanką kuriame bus ieškoma.')
        self.folder_btn.setEnabled(False)
        self.folder_btn.clicked.connect(self.select_folder)
        self.folder_btn.resize(self.folder_btn.sizeHint())
        self.folder_btn.move(150, 42)

    def where_to_put_button(self):
        lbl1 = QLabel('<b>3.</b>', self)
        lbl1.move(10, 70)
        lbl2 = QLabel("Pasirinkite rastų\nrezultatų aplanką:", self)
        lbl2.move(25, 70)
        self.where_to_btn = QPushButton('Pasirinkti', self)
        self.where_to_btn.setToolTip('Pasirinkite aplanką kuriame bus išsaugoti rezultatai.')
        self.where_to_btn.setEnabled(False)
        self.where_to_btn.clicked.connect(self.select_where_to_put)
        self.where_to_btn.resize(self.where_to_btn.sizeHint())
        self.where_to_btn.move(150, 74)

    def select_file(self):
        dialog = QFileDialog()
        file_path, _ = dialog.getOpenFileName(None, caption="Pasirinkite failą", directory=expanduser("~"),
                                              filter="Excel files (*.xlsx *.xls *.xlt)")
        if file_path:
            self.file_path = file_path
            file = file_path.split("/")[-1]
            self.main_file = file
            self.box.appendPlainText(f"Pasirinktas '{file}' kaip pagrindinis dokumentas.")
            self.folder_btn.setEnabled(True)
            self.file_btn.setText("Pasirinkta")
            self.file_btn.setEnabled(False)
            self.text_input_one.setEnabled(True)

    def select_folder(self):
        dialog = QFileDialog()
        folder_path = dialog.getExistingDirectory(None, caption="Pasirinkite aplanką", directory=expanduser("~"))
        if folder_path:
            self.folder_path = folder_path
            self.box.appendPlainText(f"Pasirinktas '{os.path.basename(folder_path)}' aplankas.")
            self.where_to_btn.setEnabled(True)
            self.folder_btn.setText("Pasirinkta")
            self.folder_btn.setEnabled(False)
            self.text_input_two.setEnabled(True)

    def select_where_to_put(self):
        dialog = QFileDialog()
        folder_path = dialog.getExistingDirectory(None, caption="Pasirinkite aplanką", directory=expanduser("~"))
        if folder_path:
            self.results_f_path = folder_path
            self.box.appendPlainText(f"Rezultatai bus talpinami '{os.path.basename(folder_path)}' aplanke.")
            self.where_to_btn.setText("Pasirinkta")
            self.where_to_btn.setEnabled(False)
            self.start_btn.setEnabled(True)

    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Uždarymo langas', 'Ar tikrai norite uždaryti programą?',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())
