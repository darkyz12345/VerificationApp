import sys
import os
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt5.QtGui import QRegExpValidator
from PyQt5.QtCore import QRegExp

from mainwindow import Ui_MainWindow
from create_excel_file import create_excel


class MAkeVerificationApp(QMainWindow):
    first_file_pdf_path = None
    second_file_pdf_path = None
    path_to_save = None
    def __init__(self):
        super(MAkeVerificationApp, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.second_file_path.adjustSize()
        self.ui.first_year_line_edit.setText(str(datetime.now().year - 1))
        self.ui.second_year_line_edit.setText(str(datetime.now().year))
        validator_int_point = QRegExpValidator(QRegExp(r'(-)*([0-9]|\.)+'))
        validator_int = QRegExpValidator(QRegExp(r'[0-9]*'))
        self.ui.method_top_line_edit.setValidator(validator_int_point)
        self.ui.method_bottom_line_edit.setValidator(validator_int_point)
        self.ui.first_year_line_edit.setValidator(validator_int)
        self.ui.second_year_line_edit.setValidator(validator_int)
        self.ui.file_first_certify_btn.clicked.connect(self.show_dialog_first_certify)
        self.ui.file_second_certify_btn.clicked.connect(self.show_dialog_second_certify)
        self.ui.save_btn.clicked.connect(self.save)

    def show_dialog_first_certify(self):
        file_name = QFileDialog.getOpenFileName(self, 'Выберите файл')
        if file_name:
            self.first_file_pdf_path = file_name[0]
            self.ui.first_file_path.setText(self.first_file_pdf_path.split('/')[-1])
            self.ui.first_file_path.adjustSize()

    def show_dialog_second_certify(self):
        file_name = QFileDialog.getOpenFileName(self, 'Выберите файл')
        if file_name:
            self.second_file_pdf_path = file_name[0]
            self.ui.second_file_path.setText(self.second_file_pdf_path.split('/')[-1])
            self.ui.second_file_path.adjustSize()

    def save(self):
        if not (self.ui.first_year_line_edit.text() and self.ui.second_year_line_edit.text()):
            QMessageBox.warning(self, 'WARNING', 'Вы не заполнили поле "Год выпуска сертификата')
            return None
        if not (self.ui.method_top_line_edit.text() and self.ui.method_bottom_line_edit.text()):
            QMessageBox.warning(self, 'WARNING', 'Вы не заполнили поле "Критерии приемки')
            return None
        if not (self.first_file_pdf_path and self.second_file_pdf_path):
            QMessageBox.warning(self, 'WARNING', 'Вы не указали путь к сертификату')
            return None
        if (self.ui.first_file_path.text() == self.ui.second_file_path.text()):
            QMessageBox.warning(self, 'WARNING', 'Для двух РАЗНЫХ сертификатов вы указали одинаковый путь')
            return None
        last_year = int(self.ui.first_year_line_edit.text())
        this_year = int(self.ui.second_year_line_edit.text())
        method_top = float(self.ui.method_top_line_edit.text())
        method_bottom = float(self.ui.method_bottom_line_edit.text())
        self.showSaveDialog()
        try:
            save_file_path = create_excel(self.first_file_pdf_path, self.second_file_pdf_path, last_year, this_year, method_top,
                         method_bottom, self.path_to_save)
        except Exception as err:
            QMessageBox.warning(self, 'Warning', 'Что-то пошло не так. Сообщите разработчику о проблеме. '
                                                 'Лучше всего передать вводимые файлы сертификатов и все вводимые значения в приложение. '
                                                 'Для связи с разработчиком: t.me/sapless')
        else:
            text = f"Поздравляем, таблица успешно сохранена по пути {save_file_path}"
            QMessageBox.warning(self, 'Успешно', text)


    def showSaveDialog(self):
        options = QFileDialog.Options()
        directory = QFileDialog.getExistingDirectory(self, 'Сохранить таблицу', os.path.expanduser("~"), options=options)
        if directory:
            self.path_to_save = directory

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MAkeVerificationApp()
    window.show()
    sys.exit(app.exec())
