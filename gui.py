import sys
import os
from pathlib import Path
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QPushButton, QFileDialog, 
                             QLineEdit, QCheckBox, QGroupBox, QMessageBox, 
                             QProgressBar, QTextEdit, QComboBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
import copy
import time

from main import parser, edit, translate, to_pdf, to_docx


class WorkerThread(QThread):
    """Поток для выполнения обработки данных в фоновом режиме"""
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool, str)
    
    def __init__(self, input_file, output_file, selected_risks, output_format, skip_translation):
        super().__init__()
        self.input_file = input_file
        self.output_file = output_file
        self.selected_risks = selected_risks
        self.output_format = output_format
        self.skip_translation = skip_translation
    
    def run(self):
        try:
            self.progress.emit("Начало обработки CSV файла...")
            data = parser(self.input_file)
            
            self.progress.emit(f"Фильтрация данных (уровни: {', '.join(self.selected_risks)})...")
            data = edit(data, self.selected_risks)
            
            if len(data) <= 1:
                self.finished.emit(False, "После фильтрации не осталось данных для обработки.\nУбедитесь, что в CSV файле есть записи с выбранными уровнями риска.")
                return
            
            if not self.skip_translation:
                self.progress.emit("Перевод данных на русский язык...")
                data = translate(data, skip_translation=False)
            else:
                self.progress.emit("Пропуск перевода...")
                data = translate(data, skip_translation=True)
            
            self.progress.emit(f"Создание {self.output_format.upper()} файла...")
            
            if self.output_format == 'pdf':
                to_pdf(data, self.output_file)
            else:
                to_docx(data, self.output_file)
            
            self.finished.emit(True, f"Файл успешно создан: {self.output_file}")
            
        except Exception as e:
            import traceback
            error_msg = f"Ошибка: {str(e)}\n\n{traceback.format_exc()}"
            self.finished.emit(False, error_msg)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker_thread = None
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle('Конвертер CSV в PDF/DOCX - Отчеты по уязвимостям')
        self.setGeometry(100, 100, 600, 650)
        
        # Центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Главный layout
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # Заголовок
        title = QLabel('Конвертер CSV в PDF/DOCX')
        title.setFont(QFont('Arial', 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title)
        
        # Группа выбора входного файла
        input_group = QGroupBox('Входной файл')
        input_layout = QVBoxLayout()
        
        input_file_layout = QHBoxLayout()
        self.input_file_edit = QLineEdit()
        self.input_file_edit.setPlaceholderText('Выберите CSV файл...')
        input_file_btn = QPushButton('Обзор...')
        input_file_btn.clicked.connect(self.select_input_file)
        input_file_layout.addWidget(self.input_file_edit)
        input_file_layout.addWidget(input_file_btn)
        input_layout.addLayout(input_file_layout)
        
        input_group.setLayout(input_layout)
        main_layout.addWidget(input_group)
        
        # Группа выбора уровней уязвимостей
        risk_group = QGroupBox('Выберите уровни уязвимостей для отображения')
        risk_layout = QVBoxLayout()
        
        self.critical_check = QCheckBox('Critical (Критический)')
        self.critical_check.setChecked(True)
        risk_layout.addWidget(self.critical_check)
        
        self.high_check = QCheckBox('High (Высокий)')
        self.high_check.setChecked(True)
        risk_layout.addWidget(self.high_check)
        
        self.medium_check = QCheckBox('Medium (Средний)')
        self.medium_check.setChecked(False)
        risk_layout.addWidget(self.medium_check)
        
        self.low_check = QCheckBox('Low (Низкий)')
        self.low_check.setChecked(False)
        risk_layout.addWidget(self.low_check)
        
        self.info_check = QCheckBox('Info (Информационный)')
        self.info_check.setChecked(False)
        risk_layout.addWidget(self.info_check)
        
        self.none_check = QCheckBox('None (Отсутствует)')
        self.none_check.setChecked(False)
        risk_layout.addWidget(self.none_check)
        
        risk_group.setLayout(risk_layout)
        main_layout.addWidget(risk_group)
        
        # Группа выбора выходного файла
        output_group = QGroupBox('Выходной файл')
        output_layout = QVBoxLayout()
        
        # Формат файла
        format_layout = QHBoxLayout()
        format_label = QLabel('Формат:')
        self.format_combo = QComboBox()
        self.format_combo.addItems(['DOCX', 'PDF'])
        format_layout.addWidget(format_label)
        format_layout.addWidget(self.format_combo)
        format_layout.addStretch()
        output_layout.addLayout(format_layout)
        
        # Имя файла
        output_file_layout = QHBoxLayout()
        self.output_file_edit = QLineEdit()
        self.output_file_edit.setPlaceholderText('Введите имя выходного файла...')
        output_file_btn = QPushButton('Обзор...')
        output_file_btn.clicked.connect(self.select_output_file)
        output_file_layout.addWidget(self.output_file_edit)
        output_file_layout.addWidget(output_file_btn)
        output_layout.addLayout(output_file_layout)
        
        output_group.setLayout(output_layout)
        main_layout.addWidget(output_group)
        
        # Опции
        options_group = QGroupBox('Опции')
        options_layout = QVBoxLayout()
        
        self.skip_translation_check = QCheckBox('Пропустить перевод (для тестирования)')
        self.skip_translation_check.setChecked(False)
        options_layout.addWidget(self.skip_translation_check)
        
        options_group.setLayout(options_layout)
        main_layout.addWidget(options_group)
        
        # Кнопка обработки
        self.process_btn = QPushButton('Обработать')
        self.process_btn.setFont(QFont('Arial', 12, QFont.Bold))
        self.process_btn.clicked.connect(self.process_file)
        main_layout.addWidget(self.process_btn)
        
        # Прогресс-бар
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # Лог-окно
        log_label = QLabel('Лог выполнения:')
        main_layout.addWidget(log_label)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        main_layout.addWidget(self.log_text)
        
        # Обновление расширения файла при изменении формата
        self.format_combo.currentTextChanged.connect(self.update_output_extension)
    
    def select_input_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            'Выберите CSV файл', 
            '', 
            'CSV Files (*.csv);;All Files (*)'
        )
        if file_path:
            self.input_file_edit.setText(file_path)
    
    def select_output_file(self):
        format_ext = '.docx' if self.format_combo.currentText() == 'DOCX' else '.pdf'
        file_path, _ = QFileDialog.getSaveFileName(
            self, 
            'Сохранить как', 
            '', 
            f'{self.format_combo.currentText()} Files (*{format_ext});;All Files (*)'
        )
        if file_path:
            # Убедимся, что расширение правильное
            if not file_path.endswith(format_ext):
                file_path += format_ext
            self.output_file_edit.setText(file_path)
    
    def update_output_extension(self):
        """Обновляет расширение выходного файла при изменении формата"""
        current_text = self.output_file_edit.text()
        if current_text:
            # Удаляем старое расширение
            base_name = Path(current_text).stem
            # Добавляем новое расширение
            new_ext = '.docx' if self.format_combo.currentText() == 'DOCX' else '.pdf'
            new_path = str(Path(current_text).parent / (base_name + new_ext))
            self.output_file_edit.setText(new_path)
    
    def get_selected_risks(self):
        """Возвращает список выбранных уровней риска"""
        risks = []
        if self.critical_check.isChecked():
            risks.append('Critical')
        if self.high_check.isChecked():
            risks.append('High')
        if self.medium_check.isChecked():
            risks.append('Medium')
        if self.low_check.isChecked():
            risks.append('Low')
        if self.info_check.isChecked():
            risks.append('Info')
        if self.none_check.isChecked():
            risks.append('None')
        return risks
    
    def log_message(self, message):
        """Добавляет сообщение в лог"""
        self.log_text.append(message)
        QApplication.processEvents()
    
    def process_file(self):
        # Проверка входного файла
        input_file = self.input_file_edit.text().strip()
        if not input_file:
            QMessageBox.warning(self, 'Ошибка', 'Пожалуйста, выберите входной CSV файл!')
            return
        
        if not os.path.exists(input_file):
            QMessageBox.warning(self, 'Ошибка', f'Файл не найден: {input_file}')
            return
        
        # Проверка выбранных уровней риска
        selected_risks = self.get_selected_risks()
        if not selected_risks:
            QMessageBox.warning(self, 'Ошибка', 'Пожалуйста, выберите хотя бы один уровень уязвимости!')
            return
        
        # Проверка выходного файла
        output_file = self.output_file_edit.text().strip()
        if not output_file:
            QMessageBox.warning(self, 'Ошибка', 'Пожалуйста, укажите имя выходного файла!')
            return
        
        # Получаем формат
        output_format = self.format_combo.currentText().lower()
        
        # Убедимся, что расширение правильное
        if output_format == 'docx' and not output_file.endswith('.docx'):
            output_file += '.docx'
        elif output_format == 'pdf' and not output_file.endswith('.pdf'):
            output_file += '.pdf'
        
        # Обновляем поле с правильным расширением
        self.output_file_edit.setText(output_file)
        
        # Блокируем кнопку и показываем прогресс
        self.process_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Неопределенный прогресс
        self.log_text.clear()
        
        # Создаем и запускаем поток обработки
        self.worker_thread = WorkerThread(
            input_file,
            output_file,
            selected_risks,
            output_format,
            self.skip_translation_check.isChecked()
        )
        self.worker_thread.progress.connect(self.log_message)
        self.worker_thread.finished.connect(self.on_processing_finished)
        self.worker_thread.start()
    
    def on_processing_finished(self, success, message):
        """Обработчик завершения обработки"""
        self.progress_bar.setVisible(False)
        self.process_btn.setEnabled(True)
        
        if success:
            QMessageBox.information(self, 'Успех', message)
            self.log_message(f"✓ {message}")
        else:
            QMessageBox.critical(self, 'Ошибка', message)
            self.log_message(f"✗ {message}")


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()

