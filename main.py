import copy

from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet
from deep_translator import GoogleTranslator
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from docx.shared import RGBColor

import csv
import time
import os
from pathlib import Path


def parser(file='test.csv'):  # запись данных в список
    """
    Безопасная функция чтения CSV файла с валидацией пути и обработкой ошибок.
    
    Args:
        file: Путь к CSV файлу (по умолчанию 'test.csv')
    
    Returns:
        list: Список строк из CSV файла
    
    Raises:
        FileNotFoundError: Если файл не найден
        PermissionError: Если нет доступа к файлу
        ValueError: Если файл имеет недопустимый формат
    """
    data = []
    
    # Защита от Path Traversal: нормализация пути и проверка на безопасность
    try:
        # Получаем абсолютный путь и нормализуем его
        file_path = Path(file).resolve()
        
        # Проверяем, что файл существует
        if not file_path.exists():
            raise FileNotFoundError(f"Файл не найден: {file}")
        
        # Проверяем, что это файл, а не директория
        if not file_path.is_file():
            raise ValueError(f"Указанный путь не является файлом: {file}")
        
        # Проверяем расширение файла
        if file_path.suffix.lower() != '.csv':
            raise ValueError(f"Файл должен иметь расширение .csv: {file}")
        
        # Проверяем размер файла (ограничение до 100MB для безопасности)
        file_size = file_path.stat().st_size
        max_size = 100 * 1024 * 1024  # 100 MB
        if file_size > max_size:
            raise ValueError(f"Файл слишком большой (максимум 100MB): {file_size / (1024*1024):.2f}MB")
        
    except (OSError, ValueError) as e:
        print(f"Ошибка валидации файла: {e}")
        raise
    
    # Безопасное чтение CSV файла с обработкой ошибок
    try:
        with open(file_path, 'r', encoding='utf-8', newline='') as f:
            reader = csv.reader(f)
            for row_num, row in enumerate(reader, start=1):
                # Валидация структуры CSV: проверяем минимальное количество колонок
                if row_num == 1:
                    # Проверка заголовка
                    if len(row) < 6:
                        raise ValueError(f"CSV файл должен содержать минимум 6 колонок. Найдено: {len(row)}")
                    expected_headers = ['Host', 'Port', 'Name', 'Description', 'Solution', 'Risk Factor']
                    if row[0].strip() != expected_headers[0] or row[5].strip() != expected_headers[5]:
                        print(f"Предупреждение: Заголовки CSV могут не соответствовать ожидаемым. "
                              f"Ожидается: {expected_headers[0]}, {expected_headers[5]}")
                
                # Проверка количества колонок в строке
                if len(row) < 6:
                    print(f"Предупреждение: Строка {row_num} содержит меньше 6 колонок. Пропущена.")
                    continue
                
                data.append(row)
        
        if len(data) == 0:
            raise ValueError("CSV файл пуст или не содержит данных")
        
        print(f'Количество строк: {len(data)}')
        
    except UnicodeDecodeError as e:
        raise ValueError(f"Ошибка декодирования файла. Убедитесь, что файл в кодировке UTF-8: {e}")
    except csv.Error as e:
        raise ValueError(f"Ошибка парсинга CSV файла: {e}")
    except PermissionError:
        raise PermissionError(f"Нет доступа к файлу: {file}")
    except Exception as e:
        raise RuntimeError(f"Неожиданная ошибка при чтении файла: {e}")

    return data


def edit(data: list, selected_risks=None):  # функция приведения списка к надлежащему виду
    """
    Функция обработки данных с фильтрацией уязвимостей по выбранным уровням риска.
    
    Args:
        data: Список строк из CSV файла
        selected_risks: Список выбранных уровней риска (например, ['Critical', 'High', 'Medium', 'Low', 'Info'])
                      Если None, по умолчанию используются ['Critical', 'High']
    
    Returns:
        list: Обработанный список данных
    
    Raises:
        ValueError: Если данные имеют неверный формат
        IndexError: Если структура данных некорректна
    """
    if not data or len(data) == 0:
        raise ValueError("Входные данные пусты")
    
    # По умолчанию используем Critical и High, если не указано иное
    if selected_risks is None:
        selected_risks = ['Critical', 'High']
    
    temp = []
    clean = []
    sort = []
    fin = []
    
    # Фильтрация по выбранным уровням риска
    risk = ['Risk Factor'] + selected_risks
    
    try:
        # Фильтрация по риск-фактору (только Critical и High)
        for i in range(len(data)):
            if len(data[i]) < 6:
                continue  # Пропускаем строки с недостаточным количеством колонок
            if data[i][5] in risk:
                temp.append(data[i])
        
        risk_names = ', '.join(selected_risks)
        print(f'Количество строк после фильтрации по риск-фактору ({risk_names}):', len(temp))
        
        if len(temp) == 0:
            print(f"Предупреждение: Не найдено уязвимостей с выбранными уровнями риска ({risk_names}). "
                  f"Убедитесь, что в CSV файле есть записи с Risk Factor из списка: {selected_risks}")
            return [['УРОВЕНЬ РИСКА', 'УЯЗВИМОСТЬ ИЛИ НЕДОСТАТОК МЕХАНИЗМА ЗАЩИТЫ', 
                    'ОПИСАНИЕ', 'РЕКОМЕНДАЦИИ', 'УЯЗВИМЫЕ РЕСУРСЫ']]
    except IndexError as e:
        raise IndexError(f"Ошибка доступа к данным: {e}. Проверьте структуру CSV файла.")

    try:
        # Объединение колонок хост и порт
        for i in range(len(temp)):
            if len(temp[i]) < 2:
                continue
            host = str(temp[i][0]).strip() if temp[i][0] else ''
            port = str(temp[i][1]).strip() if temp[i][1] else ''
            temp[i].append(f'{host}:{port}')
            temp[i].pop(0)
            temp[i].pop(0)

        # Удаление дублируемых строк
        for i in range(len(temp)):
            if temp[i] not in clean:
                clean.append(temp[i])
        
        print('Количество строк после удаления дубликатов:', len(clean))

        # Перестановка столбцов
        for i in range(len(clean)):
            if len(clean[i]) < 4:
                continue
            clean[i].insert(0, clean[i][3])
            if len(clean[i]) > 4:
                clean[i].pop(4)

        # Сортировка по риск-фактору в порядке: Critical, High, Medium, Low, Info
        risk_order = ['Critical', 'High', 'Medium', 'Low', 'Info', 'None']
        # Фильтруем только выбранные уровни риска
        for name in risk_order:
            if name in selected_risks:
                for i in range(len(clean)):
                    if len(clean[i]) > 0 and clean[i][0] == name:
                        sort.append(clean[i])
    except (IndexError, ValueError) as e:
        raise ValueError(f"Ошибка обработки данных: {e}. Проверьте структуру CSV файла.")

    for i in range(len(sort)):  # объединение строк с одинаковым названием
        for j in range(len(sort)):
            if sort[i][1] == sort[j][1] and i != j:
                sort[i][4] = f'{sort[i][4]}\n{sort[j][4]}'
                sort[j][1] += 'na_uda1enie'

    for i in range(len(sort)):  # объединение строк с одинаковым названием
        if 'na_uda1enie' not in sort[i][1]:
            fin.append(sort[i])
    print('Количество строк после объединения:', len(fin))

    while True:  # Удаление лишних пробелов
        a = 0
        for i in range(len(fin)):
            for j in range(len(fin[i])):
                if fin[i][j].find('  ') != -1:
                    fin[i][j] = fin[i][j].replace('  ', ' ')
                    a += 1
        if not a:
            break

    # Разбиение строк для переводчика (<5000 символов) и для полного отображения на странице
    try:
        while True:
            count = 0  # счетчик количества новых строк после разбиения ячеек до 3000 символов
            for i in range(len(fin)):
                if len(fin[i]) > 2 and len(str(fin[i][2])) > 3000:
                    try:
                        index = str(fin[i][2]).rindex('\n\n', 0, 3000)
                        fin.insert(i + 1, [fin[i][0], '', str(fin[i][2])[index + 1:], '', ''])
                        fin[i][2] = str(fin[i][2])[:index]
                        count += 1
                    except ValueError:
                        # Если нет '\n\n' в первых 3000 символах, разбиваем по позиции
                        fin.insert(i + 1, [fin[i][0], '', str(fin[i][2])[3000:], '', ''])
                        fin[i][2] = str(fin[i][2])[:3000]
                        count += 1
            if not count:
                break

        while True:
            count = 0
            for i in range(len(fin)):
                if len(fin[i]) > 4 and str(fin[i][4]).count('\n') > 41:
                    try:
                        index = str(fin[i][4]).rindex('\n', 0, 15 * 40)
                        fin.insert(i+1, [fin[i][0], fin[i][1], fin[i][2], fin[i][3], str(fin[i][4])[index + 1:]])
                        fin[i][4] = str(fin[i][4])[:index]
                        count += 1
                    except ValueError:
                        break
            if not count:
                break
    except (IndexError, ValueError) as e:
        print(f"Предупреждение при разбиении строк: {e}")

    return fin


def translate(data: list, skip_translation=False):  # перевод списка с английского на русский
    """
    Перевод данных с английского на русский с обработкой ошибок API.
    
    Args:
        data: Список данных для перевода
        skip_translation: Если True, пропускает перевод (для тестирования)
    
    Returns:
        list: Переведенный список данных
    """
    if not data or len(data) == 0:
        return data
    
    # Замена заголовка
    if len(data) > 0:
        data[0] = ['УРОВЕНЬ РИСКА', 'УЯЗВИМОСТЬ ИЛИ НЕДОСТАТОК МЕХАНИЗМА ЗАЩИТЫ', 
                   'ОПИСАНИЕ', 'РЕКОМЕНДАЦИИ', 'УЯЗВИМЫЕ РЕСУРСЫ']
    
    # Если перевод пропущен, возвращаем данные с заголовками
    if skip_translation:
        print("Перевод пропущен. Используется оригинальный текст.")
        return data
    
    # Перевод данных с обработкой ошибок
    translator = GoogleTranslator(source='en', target='ru')
    
    # Подсчет общего количества ячеек для перевода
    total_cells = 0
    cells_to_translate = []
    for i in range(1, len(data)):
        for j in range(len(data[i])):
            text = str(data[i][j])
            if text and text.strip() != '':
                total_cells += 1
                cells_to_translate.append((i, j))
    
    if total_cells == 0:
        print("Нет данных для перевода.")
        return data
    
    print(f"Найдено {total_cells} ячеек для перевода. Начало перевода...")
    
    translated_count = 0
    error_count = 0
    
    # Оптимизированный перевод с прогресс-баром
    for idx, (i, j) in enumerate(cells_to_translate, 1):
        try:
            text = str(data[i][j])
            
            # Пропускаем пустые строки (уже проверено, но на всякий случай)
            if not text or text.strip() == '':
                continue
            
            # Ограничение длины текста для API (максимум 5000 символов)
            original_text = text
            if len(text) > 5000:
                text = text[:5000]
                print(f"\nПредупреждение: Текст в строке {i}, колонке {j} обрезан до 5000 символов")
            
            # Перевод
            translated = translator.translate(text)
            data[i][j] = translated
            translated_count += 1
            
            # Прогресс-бар
            progress = (idx / total_cells) * 100
            print(f"\rПеревод: {idx}/{total_cells} ({progress:.1f}%) - Строка {i}", end='', flush=True)
            
            # Адаптивная задержка: уменьшаем задержку для ускорения, но оставляем минимальную
            # Задержка только каждые 5 запросов для ускорения
            if idx % 5 == 0:
                time.sleep(0.05)  # Уменьшена задержка
            elif idx % 10 == 0:
                time.sleep(0.1)   # Немного больше задержка каждые 10 запросов
                
        except Exception as e:
            error_count += 1
            print(f"\nОшибка перевода строки {i}, колонки {j}: {e}. Используется оригинальный текст.")
            # В случае ошибки оставляем оригинальный текст
            continue
    
    print(f"\n\nПеревод завершен: успешно переведено {translated_count} ячеек, ошибок: {error_count}")
    
    return data


def to_pdf(data: list, name='test.pdf'):  # генерация pdf-файла
    critical = 0
    high = 0
    medium = 0
    low = 0
    info = 0
    risk = ['Критический', 'Высокий', 'Средний', 'Низкий', 'Информационный', 'Critical', 'High', 'Medium', 'Low', 'Info', 'None']

    for i in range(len(data)):  # Удаление лишних переносов (кроме правого столбца) и исправление ошибки по тэгам
        for j in range(1, len(data[i]) - 1):
            data[i][j] = data[i][j].replace('<', ' < ')
            data[i][j] = data[i][j].replace('\n\n\n', '\n\n')
            data[i][j] = data[i][j].replace('\n\n', '<br />')
            data[i][j] = data[i][j].replace('\n', ' ')

    for i in range(len(data)):  # Удаление лишних пробелов
        for j in range(1, len(data[i]) - 1):
            data[i][j] = data[i][j].replace('  ', ' ')

    for i in range(len(data)):  # Замена переноса на тег (pdf не воспринимает \n) и коррекция перевода
        data[i][4] = data[i][4].replace('\n', '<br />')
        data[i][0] = data[i][0].replace('Середина', 'Средний')

    for i in range(1, len(data)):  # Счетчики количества строк по риск-фактору
        if data[i][0] == risk[0] or data[i][0] == risk[4]:
            critical += 1
        elif data[i][0] == risk[1] or data[i][0] == risk[5]:
            high += 1
        elif data[i][0] == risk[2] or data[i][0] == risk[6]:
            medium += 1

    # Регистрация шрифтов с проверкой существования файлов
    font_dir = Path(__file__).parent
    arial_path = font_dir / 'arial.ttf'
    arialbd_path = font_dir / 'arialbd.ttf'
    
    try:
        if arial_path.exists():
            pdfmetrics.registerFont(TTFont('arial', str(arial_path)))
        else:
            print("Предупреждение: Файл arial.ttf не найден. Используется стандартный шрифт.")
    except Exception as e:
        print(f"Ошибка загрузки шрифта arial.ttf: {e}")
    
    try:
        if arialbd_path.exists():
            pdfmetrics.registerFont(TTFont('arialbd', str(arialbd_path)))
        else:
            print("Предупреждение: Файл arialbd.ttf не найден. Используется стандартный шрифт.")
    except Exception as e:
        print(f"Ошибка загрузки шрифта arialbd.ttf: {e}")

    stylesHead = getSampleStyleSheet()
    styleHead = stylesHead['BodyText']
    styleHead.fontName = 'arialbd'
    styleHead.textColor = '#5a5a5a'
    styleHead.alignment = 1  # центрирует текст внутри параграфа
    styleHead.fontSize = 8

    stylesName = getSampleStyleSheet()
    styleName = stylesName['BodyText']
    styleName.fontName = 'arialbd'
    styleName.textColor = '#000000'
    styleName.alignment = 1  # центрирует текст внутри параграфа
    styleName.fontSize = 8

    stylesText = getSampleStyleSheet()
    styleText = stylesText['BodyText']
    styleText.fontName = 'arial'
    styleText.textColor = '#000000'
    styleText.alignment = 1  # центрирует текст внутри параграфа
    styleText.fontSize = 6

    stylesRisk = getSampleStyleSheet()
    styleRisk = stylesRisk['BodyText']
    styleRisk.fontName = 'arialbd'
    styleRisk.textColor = '#ffffff'
    styleRisk.alignment = 1  # центрирует текст внутри параграфа
    styleRisk.fontSize = 8

    for i in range(len(data[0])):
        data[0][i] = Paragraph(data[0][i], styleHead)

    for i in range(1, len(data)):
        data[i][1] = Paragraph(data[i][1], styleName)

    for i in range(1, len(data)):
        for j in range(2, len(data[i])):
            data[i][j] = Paragraph(data[i][j], styleText)

    for i in range(1, len(data)):
        data[i][0] = Paragraph(data[i][0], styleRisk)

    t = Table(data, colWidths=[23 * mm, 40 * mm, 138 * mm, 52 * mm, 28 * mm], repeatRows=1,
              rowHeights=[15 * mm] + [None] * (len(data) - 1))
    t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), (0.7, 0.7, 0.7)),
                           ('BOX', (0, 0), (-1, 0), 2, (1, 1, 1)),
                           ('INNERGRID', (0, 0), (-1, 0), 2, (1, 1, 1)),
                           ('BACKGROUND', (1, 1), (1, -1), (0.85, 0.85, 0.85)),
                           ('BOX', (1, 1), (1, -1), 2, (1, 1, 1)),
                           ('INNERGRID', (1, 1), (1, -1), 2, (1, 1, 1)),
                           ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                           ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                           ('LINEAFTER', (2, 1), (-1, -1), 0.5, (0.7, 0.7, 0.7)),
                           ('LINEBELOW', (2, 1), (-1, -1), 0.5, (0.7, 0.7, 0.7)),
                           ('BACKGROUND', (0, 1), (0, critical), '#670101'),
                           ('BACKGROUND', (0, critical + 1), (0, critical + high), '#970101'),
                           ('BACKGROUND', (0, critical + high + 1), (0, critical + high + medium), '#b29700'),
                           ('BACKGROUND', (0, critical + high + medium + 1), (0, critical + high + medium + low), '#3f9500'),
                           ('BACKGROUND', (0, critical + high + medium + low + 1), (0, -1), '#0066cc'),
                           ('BOX', (0, 1), (0, -1), 2, (1, 1, 1)),
                           ('INNERGRID', (0, 1), (0, -1), 2, (1, 1, 1)), ]))

    e = []
    e.append(t)
    
    # Валидация имени выходного файла
    output_path = Path(name).resolve()
    if output_path.suffix.lower() != '.pdf':
        output_path = output_path.with_suffix('.pdf')
    
    try:
        doc = SimpleDocTemplate(str(output_path), pagesize=landscape(A4), rightMargin=10 * mm, 
                               leftMargin=10 * mm, topMargin=6 * mm, bottomMargin=6 * mm, 
                               title='Convert csv to pdf', author='41')
        doc.build(e)
        print(f"PDF файл успешно создан: {output_path}")
    except Exception as e:
        raise RuntimeError(f"Ошибка создания PDF файла: {e}")


def to_docx(data: list, name='test.docx'):  # генерация docx-файла
    critical = 0
    high = 0
    medium = 0
    low = 0
    info = 0
    risk = ['Критический', 'Высокий', 'Средний', 'Низкий', 'Информационный', 'Critical', 'High', 'Medium', 'Low', 'Info', 'None']

    for i in range(len(data)):  # Удаление лишних переносов для docx и коррекция перевода
        data[i][0] = data[i][0].replace('Середина', 'Средний')
        for j in range(1, len(data[i]) - 1):
            data[i][j] = data[i][j].replace('\n\n', '<br />')
            data[i][j] = data[i][j].replace('\n', ' ')
            data[i][j] = data[i][j].replace('<br />', '\n')

    for i in range(1, len(data)):  # Счетчики количества строк по риск-фактору
        if data[i][0] == risk[0] or data[i][0] == risk[5]:
            critical += 1
        elif data[i][0] == risk[1] or data[i][0] == risk[6]:
            high += 1
        elif data[i][0] == risk[2] or data[i][0] == risk[7]:
            medium += 1
        elif data[i][0] == risk[3] or data[i][0] == risk[8]:
            low += 1
        elif data[i][0] == risk[4] or data[i][0] == risk[9] or data[i][0] == risk[10]:
            info += 1

    document = Document()
    section = document.sections[0]
    # section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Cm(29.7)
    section.page_height = Cm(21)
    section.left_margin = Cm(1)
    section.right_margin = Cm(1)
    section.top_margin = Cm(1)
    section.bottom_margin = Cm(1)
    table = document.add_table(rows=len(data), cols=len(data[1]))
    table.autofit = False  # обязательный параметр для ручной настройки ширины колонок таблицы
    table.allow_autofit = False  # обязательный параметр для ручной настройки ширины колонок таблицы

    widths = (Cm(2.5), Cm(4), Cm(13), Cm(5.4), Cm(3.2))
    for row in table.rows:  # настройка ширины колонок таблицы
        for idx, width in enumerate(widths):
            row.cells[idx].width = width

    def set_repeat_table_header(row):  # для повтора шапки таблицы
        tr = row._tr
        trPr = tr.get_or_add_trPr()
        tblHeader = OxmlElement('w:tblHeader')
        tblHeader.set(qn('w:val'), "true")
        trPr.append(tblHeader)
        return row

    def preventDocumentBreak(document):  # запрет разделения строки на несколько страниц
        tags = document.element.xpath('//w:tr')
        rows = len(tags)
        for row in range(0, rows):
            tag = tags[row]
            child = OxmlElement('w:cantSplit')
            tag.append(child)

    def set_color_cell(cell, color_cell=None):  # цвет ячейки
        tblCell = cell._tc
        tblCellProperties = tblCell.get_or_add_tcPr()
        if color_cell:
            clShading = OxmlElement('w:shd')
            clShading.set(qn('w:fill'), color_cell)
            tblCellProperties.append(clShading)
        return cell

    def set_cell_border(cell, **kwargs):  # настройка границ ячейки
        """
        Пример использования:
        set_cell_border(
            cell,
            top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
            bottom={"sz": 12, "color": "#00FF00", "val": "single"},
            start={"sz": 24, "val": "dashed", "shadow": "true"},
            end={"sz": 12, "val": "dashed"},
        )
        """
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()

        tcBorders = tcPr.first_child_found_in("w:tcBorders")
        if tcBorders is None:
            tcBorders = OxmlElement('w:tcBorders')
            tcPr.append(tcBorders)

        for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
            edge_data = kwargs.get(edge)
            if edge_data:
                tag = 'w:{}'.format(edge)

                element = tcBorders.find(qn(tag))
                if element is None:
                    element = OxmlElement(tag)
                    tcBorders.append(element)

                for key in ["sz", "val", "color", "space", "shadow"]:
                    if key in edge_data:
                        element.set(qn('w:{}'.format(key)), str(edge_data[key]))

    for i in range(len(data[0])):
        cell = table.cell(0, i)
        cell.text = data[0][i]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        paragraph = cell.paragraphs[0]
        paragraph.paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        paragraph.paragraph_format.space_before = Pt(3)
        paragraph.paragraph_format.space_after = Pt(3)
        run = paragraph.runs
        font = run[0].font
        font.name = 'Arial'
        font.bold = True
        font.color.rgb = RGBColor(0x5a, 0x5a, 0x5a)
        font.size = Pt(8)
        set_color_cell(cell, color_cell="b4b4b4")
        set_cell_border(cell, top={"sz": 24, "val": "single", "color": "#ffffff"},
                        bottom={"sz": 24, "val": "single", "color": "#ffffff"},
                        start={"sz": 24, "val": "single", "color": "#ffffff"},
                        end={"sz": 24, "val": "single", "color": "#ffffff"})

    for i in range(1, len(data)):
        cell = table.cell(i, 0)
        cell.text = data[i][0]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        paragraph = cell.paragraphs[0]
        paragraph.paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        paragraph.paragraph_format.space_before = Pt(3)
        paragraph.paragraph_format.space_after = Pt(3)
        run = paragraph.runs
        font = run[0].font
        font.name = 'Arial'
        font.bold = True
        font.color.rgb = RGBColor(0xff, 0xff, 0xff)
        font.size = Pt(8)
        if i <= critical:
            set_color_cell(cell, color_cell="670101")
        elif i <= critical + high:
            set_color_cell(cell, color_cell="970101")
        elif i <= critical + high + medium:
            set_color_cell(cell, color_cell="b29700")
        elif i <= critical + high + medium + low:
            set_color_cell(cell, color_cell="3f9500")
        else:
            set_color_cell(cell, color_cell="0066cc")
        set_cell_border(cell, top={"sz": 24, "val": "single", "color": "#ffffff"},
                        bottom={"sz": 24, "val": "single", "color": "#ffffff"},
                        start={"sz": 24, "val": "single", "color": "#ffffff"},
                        end={"sz": 24, "val": "single", "color": "#ffffff"})

    for i in range(1, len(data)):
        cell = table.cell(i, 1)
        cell.text = data[i][1]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        paragraph = cell.paragraphs[0]
        paragraph.paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        paragraph.paragraph_format.space_before = Pt(3)
        paragraph.paragraph_format.space_after = Pt(3)
        run = paragraph.runs
        font = run[0].font
        font.name = 'Arial'
        font.bold = True
        font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        font.size = Pt(8)
        set_color_cell(cell, color_cell="dadada")
        set_cell_border(cell, top={"sz": 24, "val": "single", "color": "#ffffff"},
                        bottom={"sz": 24, "val": "single", "color": "#ffffff"},
                        start={"sz": 24, "val": "single", "color": "#ffffff"},
                        end={"sz": 24, "val": "single", "color": "#ffffff"})

    for i in range(1, len(data)):
        for j in range(2, len(data[i])):
            cell = table.cell(i, j)
            cell.text = data[i][j]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            paragraph = cell.paragraphs[0]
            paragraph.paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
            paragraph.paragraph_format.space_before = Pt(3)
            paragraph.paragraph_format.space_after = Pt(3)
            run = paragraph.runs
            font = run[0].font
            font.name = 'Arial'
            font.color.rgb = RGBColor(0x00, 0x00, 0x00)
            font.size = Pt(7)
            set_cell_border(cell, bottom={"sz": 6, "val": "single", "color": "#dadada"},
                            end={"sz": 6, "val": "single", "color": "#dadada"})

    set_repeat_table_header(table.rows[0])
    preventDocumentBreak(document)
    
    # Валидация имени выходного файла
    output_path = Path(name).resolve()
    if output_path.suffix.lower() != '.docx':
        output_path = output_path.with_suffix('.docx')
    
    try:
        document.save(str(output_path))
        print(f"DOCX файл успешно создан: {output_path}")
    except PermissionError:
        raise PermissionError(f"Нет доступа для записи файла: {output_path}. "
                            f"Убедитесь, что файл не открыт в другой программе.")
    except Exception as e:
        raise RuntimeError(f"Ошибка создания DOCX файла: {e}")


if __name__ == '__main__':
    try:
        a = time.time()
        
        # Парсинг CSV файла
        print("Начало обработки CSV файла...")
        data = parser()  # при необходимости в аргументе функции указать название файла csv
        
        # Обработка данных (фильтрация только Critical и High)
        print("Фильтрация данных (только Critical и High уязвимости)...")
        data = edit(data)
        
        if len(data) <= 1:
            print("Внимание: После фильтрации не осталось данных для обработки.")
            print("Убедитесь, что в CSV файле есть записи с Risk Factor = 'Critical' или 'High'")
            exit(0)
        
        # Перевод данных
        print("Перевод данных на русский язык...")
        # Для пропуска перевода раскомментируйте следующую строку и закомментируйте translate(data)
        # data = translate(data, skip_translation=True)
        data = translate(data)
        
        # Создание копии для DOCX
        data_docx = copy.deepcopy(data)
        
        # Генерация файлов
        # to_pdf(data)  # при необходимости в аргументе функции указать название файла pdf
        print("Создание DOCX файла...")
        to_docx(data_docx)  # при необходимости в аргументе функции указать название файла docx
        
        b = (time.time() - a) / 60
        print(f'Время выполнения скрипта: {b:.1f} min')
        print("Обработка завершена успешно!")
        
    except FileNotFoundError as e:
        print(f"Ошибка: Файл не найден. {e}")
        exit(1)
    except PermissionError as e:
        print(f"Ошибка доступа: {e}")
        exit(1)
    except ValueError as e:
        print(f"Ошибка валидации данных: {e}")
        exit(1)
    except Exception as e:
        print(f"Критическая ошибка: {e}")
        import traceback
        traceback.print_exc()
        exit(1)
