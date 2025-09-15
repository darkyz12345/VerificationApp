
**VerificationApp** - это настольное приложение для автоматизированной верификации калибровочных сертификатов измерительного оборудования [2](#0-1) . Приложение анализирует PDF-сертификаты калибровки за два разных года и генерирует Excel-отчеты с анализом дрейфа измерений и графиками [3](#0-2) .

### Основные возможности

- **Извлечение данных из PDF**: Автоматическое извлечение метаданных оборудования и таблиц неопределенности из PDF-сертификатов [4](#0-3) 
- **Сравнительный анализ**: Сопоставление результатов калибровки между двумя временными периодами [5](#0-4) 
- **Генерация отчетов**: Создание структурированных Excel-файлов с таблицами и графиками дрейфа [6](#0-5) 
- **Графический интерфейс**: Интуитивный интерфейс на PyQt5 с валидацией входных данных [7](#0-6) 

### Технические требования

Приложение использует следующие основные зависимости [8](#0-7) :
- PyQt5 для графического интерфейса
- pdfplumber для извлечения данных из PDF
- openpyxl для генерации Excel-файлов
- pandas для обработки табличных данных

### Архитектура

Приложение построено по трехуровневой архитектуре:
- **Уровень представления**: PyQt5 интерфейс с валидацией входных данных [9](#0-8) 
- **Бизнес-логика**: Обработка данных и генерация отчетов [10](#0-9) 
- **Уровень данных**: Извлечение информации из PDF-документов [11](#0-10) 

### Использование

1. Выберите два PDF-сертификата калибровки (за разные годы)
2. Укажите годы выпуска сертификатов
3. Введите критерии приемки метода
4. Выберите папку для сохранения отчета
5. Получите Excel-файл с анализом дрейфа измерений

**Notes**

Приложение специализировано для работы с русскоязычными калибровочными сертификатами и использует специфические паттерны извлечения данных для поиска полей "Объект калибровки", "Тип", "номер" и таблиц неопределенности. Интерфейс полностью на русском языке с предустановленными значениями текущего и предыдущего года.

Wiki pages you might want to explore:
- [Application Architecture (darkyz12345/VerificationApp)](/wiki/darkyz12345/VerificationApp#2)
- [User Interface (darkyz12345/VerificationApp)](/wiki/darkyz12345/VerificationApp#3)

### Citations

**File:** main.py (L12-87)
```python
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
```

**File:** create_excel_file.py (L16-24)
```python
def create_chart(row_last_year, unit, sheet):
    chart = LineChart()
    chart.title = f'Дрейф({row_last_year.values[0]} {unit})'
    # chart.style = 13
    data = Reference(sheet, min_col=4, min_row=7, max_col=7, max_row=9)
    categories = Reference(sheet, min_col=3, min_row=8, max_row=9)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    return chart
```

**File:** create_excel_file.py (L26-107)
```python
def create_excel(path_to_file_last_year, path_to_file_this_year, last_year, this_year, method_top, method_bottom, save_path):
    name_last_year, serie_number_last_year, uncertainty_last_year = find_name_number_uncertainty(path_to_file_last_year)
    # print(type(uncertainty_last_year))
    name_this_year, serie_number_this_year, uncertainty_this_year = find_name_number_uncertainty(path_to_file_this_year)
    # print(type(uncertainty_this_year))
    unit = uncertainty_this_year.columns[0].split(',')[-1].strip().split()[-1]
    for name_header in uncertainty_this_year.columns:
        if 'неопределенность' in str(name_header).lower() or 'u,' in str(name_header).lower():
            uncertainty_index = list(uncertainty_this_year).index(name_header)
            uncertainty_unit = name_header.split()[-1]
        if 'отклонение' in str(name_header).lower() or 'q,' in str(name_header).lower():
            deviation_index = list(uncertainty_this_year).index(name_header)
            deviation_unit = name_header.split()[-1]
    workbook = openpyxl.Workbook()
    for (index_last_year, row_last_year), (index_this_year, row_this_year) in zip(uncertainty_last_year.iterrows(),
                                                                                  uncertainty_this_year.iterrows()):
        tmp_func = lambda x: x.isdigit() or ',' == x or '.' == x
        if all(tmp_func(i) for i in str(row_last_year.values[0])) is False:
            continue
        sheet = workbook.create_sheet(title=f'point_{row_last_year.values[0]}')
        sheet['A1'] = 'Дата проведения верификации'
        sheet['A2'] = 'Марка, модель оборудования'
        sheet['A3'] = 'Заводской номер оборудования'
        sheet['A4'] = 'Результаты верификации'
        sheet['A5'] = 'Ответственный за верификацию оборудования'
        sheet['B1'] = datetime.now().strftime('%d.%m.%Y')
        sheet['B2'] = name_this_year
        sheet['B3'] = serie_number_this_year
        sheet['B5'] = 'Захарченко Е. Д.'
        sheet['C8'] = last_year
        sheet['C9'] = this_year
        sheet['D7'] = 'треб. метода'
        sheet['E7'] = 'треб. метода'
        sheet['F7'] = 'откл.'
        sheet['G7'] = 'неопр.'
        sheet['F6'] = unit
        sheet['G6'] = unit
        if '%' in uncertainty_unit or 'q' == uncertainty_unit:
            uncertainty_value_last_year = float(row_last_year.values[uncertainty_index].replace(',', '.')) * float(
                row_last_year.values[0].replace(',', '.')) / 100
            deviation_value_last_year = float(row_last_year.values[deviation_index].replace(',', '.')) * (
                    float(row_last_year.values[0].replace(',', '.')) / 100)
            uncertainty_value_this_year = float(row_this_year.values[uncertainty_index].replace(',', '.')) * float(
                row_this_year.values[0].replace(',', '.')) / 100
            deviation_value_this_year = (float(row_this_year.values[deviation_index].replace(',', '.')) *
                                         float(row_this_year.values[0].replace(',', '.')) / 100)
        else:
            uncertainty_value_last_year = row_last_year.values[uncertainty_index]
            deviation_value_last_year = row_last_year.values[deviation_index]
            uncertainty_value_this_year = row_this_year.values[uncertainty_index]
            deviation_value_this_year = row_this_year.values[deviation_index]
        sheet['F8'] = deviation_value_last_year
        sheet['F9'] = deviation_value_this_year
        sheet['G8'] = uncertainty_value_last_year
        sheet['G9'] = uncertainty_value_this_year
        sheet['D8'] = method_top
        sheet['D9'] = method_top
        sheet['E8'] = method_bottom
        sheet['E9'] = method_bottom
        dims = {}
        for row in sheet.rows:
            for cell in row:
                dims[cell.column_letter] = max(dims.get(cell.column_letter, 0), len(str(cell.value)))
        for col, value in dims.items():
            sheet.column_dimensions[col].width = value + 5
        border_style = Border(left=Side(border_style='thin', color='000000'),
                              right=Side(border_style='thin', color='000000'),
                              top=Side(border_style='thin', color='000000'),
                              bottom=Side(border_style='thin', color='000000'))
        for row in sheet.iter_rows(min_row=1, max_row=5, min_col=1, max_col=2):
            for cell in row:
                cell.border = border_style
        for row in sheet.iter_rows(min_row=7, max_row=9, min_col=3, max_col=7):
            for cell in row:
                cell.border = border_style
        sheet["F6"].border = border_style
        sheet["G6"].border = border_style
        sheet.add_chart(create_chart(row_last_year, unit, sheet))
    default_sheet = workbook['Sheet']
    workbook.remove(default_sheet)
    workbook.save(f'{save_path}/{name_this_year}.xlsx')
    return f'{save_path}/{name_this_year}.xlsx'
```

**File:** pdf_extract_data.py (L10-42)
```python
def find_name_number_uncertainty(path_to_file: str):
    name_list = []
    name = ''
    serie_number = ''
    uncertainty = None
    with pdfplumber.open(path_to_file) as pdf:
        text = pdf.pages[0].extract_text()
        text_info = text.split('\n')
        for txt in text_info:
            if "Объект калибровки:" in txt:
                name_list.append(txt.split(':')[-1].strip())
            if "Тип" in txt:
                name_list.append(txt.split(':')[-1].strip())
                name = ', '.join(name_list)
            if "номер" in txt:
                serie_number = txt.split(':')[-1].strip()
        for page_num, page in enumerate(pdf.pages):
            # Извлекаем таблицы
            tables = page.extract_tables()
            for table in tables:
                for header in (i.replace('\n', ' ').replace('ё', 'е').lower() if type(i) is str else i for i in
                               table[0]):
                    if type(header) is str:
                        if 'неопределенность' in header or "u" == header:
                            if 'u' == header:
                                ts = table[0]
                                units = table[1]
                                uncertainty = pd.DataFrame(table[2:], columns=tuple(
                                    str(ts[i]).replace('\n', ' ') + ', ' + str(units[i]) for i in range(len(ts))))
                            else:
                                uncertainty = pd.DataFrame(table[1:], columns=tuple(
                                    str(i).replace('\n', ' ') for i in table[0]))
    return name, serie_number, uncertainty
```

**File:** requirements.txt (L19-38)
```text
openpyxl==3.1.5
packaging==24.1
pandas==2.2.2
pdfminer.six==20231228
pdfplumber==0.11.1
pillow==10.3.0
pycparser==2.22
pyinstaller==6.8.0
pyinstaller-hooks-contrib==2024.7
pyparsing==3.1.2
PyPDF2==3.0.1
pypdfium2==4.30.0
PyQt5==5.15.10
PyQt5-Qt5==5.15.2
PyQt5-sip==12.13.0
python-dateutil==2.9.0.post0
pytz==2024.1
requests==2.32.3
six==1.16.0
tabula-py==2.9.3
```
