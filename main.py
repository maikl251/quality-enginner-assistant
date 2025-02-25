import sys
import pandas as pd
import json
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QMessageBox,
    QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView,
    QCompleter, QGroupBox, QComboBox, QFormLayout
)
from PyQt5.QtCore import Qt, QStringListModel
from PyQt5.QtGui import QColor
from datetime import datetime
import os


class EngineeringAssistant(QMainWindow):
    def __init__(self):
        super().__init__()
        self.data = pd.DataFrame(columns=[
            'ID', 'Деталь', 'Участок',
            'Количество_брака_1', 'Тип_дефекта_1',
            'Количество_брака_2', 'Тип_дефекта_2',
            'Примечание', 'Дата'
        ])
        self.history = {
            'ids': [],
            'details': [],
            'areas': []
        }
        self.current_id = None
        self.load_data()
        self.init_ui()

    def load_data(self):
        """Загружает данные из файлов"""
        if os.path.exists('engineering_data.xlsx'):
            try:
                self.data = pd.read_excel('engineering_data.xlsx', dtype={'ID': str})
                # Проверяем, что все необходимые столбцы существуют
                required_columns = [
                    'ID', 'Деталь', 'Участок',
                    'Количество_брака_1', 'Тип_дефекта_1',
                    'Количество_брака_2', 'Тип_дефекта_2',
                    'Примечание', 'Дата'
                ]
                for col in required_columns:
                    if col not in self.data.columns:
                        raise ValueError(f"Отсутствует столбец {col} в файле engineering_data.xlsx")
                # Заменяем NaN на пустые строки или 0
                self.data = self.data.fillna({
                    'Количество_брака_1': 0,
                    'Количество_брака_2': 0,
                    'Тип_дефекта_1': '-',
                    'Тип_дефекта_2': '-',
                    'Примечание': '',
                    'Дата': ''
                })
            except Exception as e:
                QMessageBox.warning(self, "Ошибка загрузки", f"Не удалось загрузить данные: {str(e)}")
                self.data = pd.DataFrame(columns=[
                    'ID', 'Деталь', 'Участок',
                    'Количество_брака_1', 'Тип_дефекта_1',
                    'Количество_брака_2', 'Тип_дефекта_2',
                    'Примечание', 'Дата'
                ])
        else:
            self.data = pd.DataFrame(columns=[
                'ID', 'Деталь', 'Участок',
                'Количество_брака_1', 'Тип_дефекта_1',
                'Количество_брака_2', 'Тип_дефекта_2',
                'Примечание', 'Дата'
            ])

        if os.path.exists('input_history.json'):
            try:
                with open('input_history.json', 'r') as f:
                    self.history = json.load(f)
            except Exception as e:
                QMessageBox.warning(self, "Ошибка загрузки", f"Не удалось загрузить историю: {str(e)}")

    def save_data(self):
        """Сохраняет все данные с форматированием в Excel"""
        try:
            # Сохраняем данные в Excel
            with pd.ExcelWriter("engineering_data.xlsx", engine='xlsxwriter') as writer:
                self.data.to_excel(writer, index=False, sheet_name='Данные')
                workbook = writer.book
                worksheet = writer.sheets['Данные']

                # Настраиваем ширину столбцов
                for i, col in enumerate(self.data.columns):
                    max_len = max(self.data[col].astype(str).map(len).max(), len(col)) + 2
                    worksheet.set_column(i, i, max_len)

                # Включаем перенос текста для всех ячеек
                cell_format = workbook.add_format({'text_wrap': True})
                worksheet.set_column(0, len(self.data.columns) - 1, None, cell_format)

            # Сохраняем историю
            with open('input_history.json', 'w') as f:
                json.dump(self.history, f)
        except Exception as e:
            QMessageBox.warning(self, "Ошибка сохранения", f"Не удалось сохранить данные: {str(e)}")

    def init_ui(self):
        self.setWindowTitle("Умный помощник инженера качества")
        self.setGeometry(100, 100, 1400, 900)
        self.setStyleSheet("""
            QWidget {
                font-size: 14px;
            }
            QGroupBox {
                font-weight: bold;
                margin-top: 15px;
                padding: 10px 0;
            }
            QComboBox, QLineEdit, QTextEdit {
                padding: 5px;
                min-height: 28px;
            }
            QTableWidget {
                font-size: 13px;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QTableWidget::item:nth-child(even) {
                background-color: #f2f2f2;
            }
            QTableWidget::item:nth-child(odd) {
                background-color: #ffffff;
            }
            QPushButton {
                padding: 8px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton#add_btn {
                background-color: #4CAF50;
                color: white;
            }
            QPushButton#finish_btn {
                background-color: #2196F3;
                color: white;
            }
            QPushButton#clear_btn {
                background-color: #f44336;
                color: white;
            }
            QPushButton#export_btn {
                background-color: #FF9800;
                color: white;
            }
        """)
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout()
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_widget.setLayout(main_layout)

        self.tabs = QTabWidget()
        self.tabs.addTab(self.create_input_tab(), "Ввод данных")
        self.tabs.addTab(self.create_table_tab(), "Общая таблица")
        main_layout.addWidget(self.tabs)
        self.statusBar().showMessage("Готов к работе! Начните вводить данные")

    def closeEvent(self, event):
        self.save_data()
        event.accept()

    def create_input_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(10, 10, 10, 10)

        detail_group = QGroupBox("Основные данные детали")
        detail_layout = QFormLayout()
        detail_layout.setVerticalSpacing(10)

        self.id_combo = QComboBox()
        self.id_combo.setEditable(True)
        self.id_combo.addItems(self.history['ids'])
        detail_layout.addRow(QLabel("Уникальный ID детали:"), self.id_combo)

        self.detail_combo = QComboBox()
        self.detail_combo.setEditable(True)
        self.detail_combo.addItems(self.history['details'])
        detail_layout.addRow(QLabel("Название детали:"), self.detail_combo)

        detail_group.setLayout(detail_layout)

        area_group = QGroupBox("Данные участка")
        area_layout = QVBoxLayout()
        area_layout.setSpacing(10)

        self.area_combo = QComboBox()
        self.area_combo.setEditable(True)
        self.area_combo.addItems(self.history['areas'])
        area_layout.addWidget(QLabel("Участок обработки:"))
        area_layout.addWidget(self.area_combo)

        defects_layout = QHBoxLayout()
        defects_layout.setSpacing(15)

        defect1_group = QGroupBox("Брак 1")
        defect1_layout = QFormLayout()
        self.defect_count_1 = QLineEdit()
        self.defect_type_1 = QComboBox()
        self.defect_type_1.setEditable(True)
        defect1_layout.addRow(QLabel("Количество:"), self.defect_count_1)
        defect1_layout.addRow(QLabel("Тип дефекта:"), self.defect_type_1)
        defect1_group.setLayout(defect1_layout)

        defect2_group = QGroupBox("Брак 2")
        defect2_layout = QFormLayout()
        self.defect_count_2 = QLineEdit()
        self.defect_type_2 = QComboBox()
        self.defect_type_2.setEditable(True)
        defect2_layout.addRow(QLabel("Количество:"), self.defect_count_2)
        defect2_layout.addRow(QLabel("Тип дефекта:"), self.defect_type_2)
        defect2_group.setLayout(defect2_layout)

        defects_layout.addWidget(defect1_group)
        defects_layout.addWidget(defect2_group)

        area_layout.addLayout(defects_layout)

        self.note_input = QTextEdit()
        area_layout.addWidget(QLabel("Примечания:"))
        area_layout.addWidget(self.note_input)

        area_group.setLayout(area_layout)

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(15)

        self.add_btn = QPushButton("Добавить запись по участку (Ctrl+Enter)")
        self.add_btn.setObjectName("add_btn")
        self.add_btn.setShortcut("Ctrl+Return")
        self.add_btn.clicked.connect(self.add_area_data)

        self.finish_btn = QPushButton("Завершить ввод по детали (Ctrl+S)")
        self.finish_btn.setObjectName("finish_btn")
        self.finish_btn.setShortcut("Ctrl+S")
        self.finish_btn.clicked.connect(self.finish_detail_input)

        self.clear_btn = QPushButton("Очистить все поля")
        self.clear_btn.setObjectName("clear_btn")
        self.clear_btn.clicked.connect(self.clear_all_fields)

        btn_layout.addWidget(self.add_btn)
        btn_layout.addWidget(self.finish_btn)
        btn_layout.addWidget(self.clear_btn)

        layout.addWidget(detail_group)
        layout.addWidget(area_group)
        layout.addLayout(btn_layout)

        tab.setLayout(layout)
        self.id_combo.currentTextChanged.connect(self.update_details_from_history)
        return tab

    def update_details_from_history(self, id_str):
        if id_str in self.data['ID'].values:
            details = self.data[self.data['ID'] == id_str]
            if not details.empty:
                self.detail_combo.setCurrentText(details.iloc[0]['Деталь'])

    def add_area_data(self):
        try:
            current_id = self.id_combo.currentText().strip()
            if not current_id:
                raise ValueError("Введите ID детали!")

            self.update_input_history()

            if not all([
                self.detail_combo.currentText().strip(),
                self.area_combo.currentText().strip()
            ]):
                raise ValueError("Заполните все обязательные поля!")

            current_date = datetime.now().strftime("%Y-%m-%d %H:%M")
            area = self.area_combo.currentText().strip()
            defect1_count = int(self.defect_count_1.text() or 0)
            defect2_count = int(self.defect_count_2.text() or 0)

            if defect1_count + defect2_count == 0:
                raise ValueError("Введите данные хотя бы для одного вида брака!")

            note = self.note_input.toPlainText().strip()
            formatted_note = f"{note}" if note else "добавлена запись"

            # Проверяем, существует ли уже запись для данного участка
            existing_record = self.data[
                (self.data['ID'] == current_id) &
                (self.data['Участок'] == area)
            ]

            if not existing_record.empty:
                idx = existing_record.index[0]
                # Убедимся, что значения числовые
                self.data.at[idx, 'Количество_брака_1'] = int(self.data.at[idx, 'Количество_брака_1']) + defect1_count
                self.data.at[idx, 'Количество_брака_2'] = int(self.data.at[idx, 'Количество_брака_2']) + defect2_count

                # Обновляем тип брака, если он был изменен
                defect_type_1 = self.defect_type_1.currentText().strip()
                defect_type_2 = self.defect_type_2.currentText().strip()
                if defect_type_1:
                    self.data.at[idx, 'Тип_дефекта_1'] = defect_type_1
                if defect_type_2:
                    self.data.at[idx, 'Тип_дефекта_2'] = defect_type_2

                # Добавляем примечание и дату
                self.data.at[idx, 'Примечание'] += f", {formatted_note}"
                self.data.at[idx, 'Дата'] += f", {current_date}"
            else:
                new_row = {
                    'ID': current_id,
                    'Деталь': self.detail_combo.currentText().strip(),
                    'Участок': area,
                    'Количество_брака_1': defect1_count,
                    'Тип_дефекта_1': self.defect_type_1.currentText().strip() or "-",
                    'Количество_брака_2': defect2_count,
                    'Тип_дефекта_2': self.defect_type_2.currentText().strip() or "-",
                    'Примечание': formatted_note,
                    'Дата': current_date
                }
                self.data = pd.concat([self.data, pd.DataFrame([new_row])], ignore_index=True)

            self.clear_input_fields()
            self.update_table()
            self.statusBar().showMessage("Данные успешно добавлены!", 3000)
        except Exception as e:
            QMessageBox.warning(self, "Ошибка ввода", str(e))

    def update_input_history(self):
        for field, widget in [
            ('ids', self.id_combo),
            ('details', self.detail_combo),
            ('areas', self.area_combo)
        ]:
            text = widget.currentText().strip()
            if text and text not in self.history[field]:
                self.history[field].insert(0, text)
                widget.addItem(text)
                widget.setCurrentText(text)

    def clear_input_fields(self):
        self.defect_count_1.clear()
        self.defect_type_1.setCurrentIndex(-1)
        self.defect_count_2.clear()
        self.defect_type_2.setCurrentIndex(-1)
        self.note_input.clear()

    def clear_all_fields(self):
        self.id_combo.setCurrentIndex(-1)
        self.detail_combo.setCurrentIndex(-1)
        self.area_combo.setCurrentIndex(-1)
        self.clear_input_fields()

    def finish_detail_input(self):
        reply = QMessageBox.question(self, 'Подтверждение', 'Вы уверены, что хотите завершить ввод по детали?',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.save_data()
            self.clear_all_fields()
            self.tabs.setCurrentIndex(1)
            self.id_combo.setFocus()

    def create_table_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        self.table = QTableWidget()
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels([
            'ID', 'Деталь', 'Участок',
            'Брак1 (кол)', 'Брак1 (тип)',
            'Брак2 (кол)', 'Брак2 (тип)',
            'Примечание', 'Дата'
        ])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSortingEnabled(True)
        self.table.setWordWrap(True)  # Включаем перенос текста
        layout.addWidget(self.table)

        export_btn = QPushButton("Экспорт в Excel")
        export_btn.setObjectName("export_btn")
        export_btn.clicked.connect(self.manual_export_to_excel)
        layout.addWidget(export_btn)

        tab.setLayout(layout)
        return tab

    def update_table(self):
        self.table.setRowCount(len(self.data))
        prev_id = None
        start_row = 0
        sorted_data = self.data.sort_values(['ID', 'Дата'])

        for i, row in sorted_data.iterrows():
            # Заполняем все ячейки для текущей строки
            for col in range(9):
                item = QTableWidgetItem(str(row.iloc[col]))
                item.setFlags(item.flags() ^ Qt.ItemIsEditable)
                self.table.setItem(i, col, item)
            
            # Объединяем ячейки для одинаковых ID и Детали
            if row['ID'] == prev_id:
                # Объединяем ячейки ID и Деталь
                self.table.setSpan(start_row, 0, i - start_row + 1, 1)  # ID
                self.table.setSpan(start_row, 1, i - start_row + 1, 1)  # Деталь
            else:
                start_row = i
                prev_id = row['ID']

        # Настраиваем отображение таблицы
        self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)  # Участок
        self.table.horizontalHeader().setSectionResizeMode(7, QHeaderView.Stretch)  # Примечание

    def manual_export_to_excel(self):
        """Экспортирует данные в Excel"""
        try:
            self.save_data()
            QMessageBox.information(self, "Экспорт завершен", 
                "Данные успешно экспортированы в engineering_data.xlsx!")
        except Exception as e:
            QMessageBox.warning(self, "Ошибка экспорта", 
                f"Произошла ошибка при экспорте: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = EngineeringAssistant()
    window.show()
    sys.exit(app.exec_())
