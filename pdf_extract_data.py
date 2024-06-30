import PyPDF2
import tabula
import pandas as pd
import pdfplumber


PATH_TO_FILE = 'Цифровой_динамометрический_ключ_AWM_100.pdf'
PATH_TO_FILE_1 = 'WDW-100D  CY20230366.pdf'

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