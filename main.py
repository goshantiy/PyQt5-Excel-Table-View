
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableView, QDialog, QVBoxLayout, QHBoxLayout, QMessageBox, QFileDialog
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QSortFilterProxyModel, QItemSelectionModel, QMimeData
import pandas as pd
import matplotlib.pyplot as plt
from PyQt5.QtWidgets import QPushButton, QWidget, QLineEdit, QInputDialog, QMenu, QAction


class ExcelTableModel(QAbstractTableModel):
    def __init__(self, data, headers, parent=None):
        super(ExcelTableModel, self).__init__(parent)
        self._data = data
        self._headers = headers

    def rowCount(self, parent=QModelIndex()):
        return len(self._data)

    def columnCount(self, parent=QModelIndex()):
        return len(self._headers)

    def data(self, index, role=Qt.DisplayRole):
        if role == Qt.DisplayRole or role == Qt.EditRole:
            return str(self._data[index.row()][index.column()])
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal:
            return str(self._headers[section])
        return None
    
    def setData(self, index, value, role=Qt.EditRole):
        if role == Qt.EditRole:
            self._data[index.row()][index.column()] = value
            self.dataChanged.emit(index, index)
            return True
        return False
    
    def flags(self, index):
        return super().flags(index) | Qt.ItemIsEditable

    

    def sort(self, column, order=Qt.AscendingOrder):
            self.layoutAboutToBeChanged.emit()

            column_to_sort = [row[column] for row in self._data]

            sorted_column = sorted(enumerate(column_to_sort), key=lambda x: x[1], reverse=(order == Qt.DescendingOrder))

            for i, (index, value) in enumerate(sorted_column):
                self._data[i][column] = value

                self.layoutChanged.emit()

    
    def saveToExcel(self, file_path):
        df = pd.DataFrame(self._data, columns=self._headers)
        df.to_excel(file_path, index=False)

    def removeRows(self, position, rows, parent=QModelIndex()):
        self.beginRemoveRows(parent, position, position + rows - 1)
        del self._data[position:position + rows]
        self.endRemoveRows()
        return True

    def removeColumns(self, position, columns, parent=QModelIndex()):
        self.beginRemoveColumns(parent, position, position + columns - 1)
        for i in range(len(self._data)):
            del self._data[i][position:position + columns]
        self._headers = self._headers[:position] + self._headers[position + columns:]
        self.endRemoveColumns()
        return True
    
    def addColumn(self):
        self.beginInsertColumns(QModelIndex(), len(self._headers), len(self._headers))
        self._headers.append(f'Column {len(self._headers)}')
        for row in self._data:
            row.append('')
        self.endInsertColumns()

    def addRow(self):
        self.beginInsertRows(QModelIndex(), len(self._data), len(self._data))
        new_row = [''] * len(self._headers)
        self._data.append(new_row)
        self.endInsertRows()
    
    def addData(self, new_data):
        self.layoutAboutToBeChanged.emit()

        self._data.append(new_data)

        self.layoutChanged.emit()

    def printData(self):
        print(f"Data row:")
        for row in self._data:
            print(row)
        print()

class SearchDialog(QDialog):
    def __init__(self, parent=None):
        super(SearchDialog, self).__init__(parent)

        self.search_input = QLineEdit()
        self.search_button = QPushButton('Search')
        self.search_button.clicked.connect(self.search)

        layout = QVBoxLayout(self)
        layout.addWidget(self.search_input)
        layout.addWidget(self.search_button)

    def search(self):
        search_term = self.search_input.text()
        self.accept() 

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        self.setWindowTitle("Программа")
        self.resize(800, 600)

        self.model = ExcelTableModel([], [])
        self.setAcceptDrops(True)

        # Кнопки главного окна
        self.load_button = QPushButton("Загрузить базу данных")
        self.load_button.clicked.connect(self.load_data)

        self.search_button = QPushButton("Поиск")
        self.search_button.clicked.connect(self.show_search_dialog)

        self.filter_button = QPushButton("Фильтр по строке")
        self.filter_button.clicked.connect(self.show_filter_dialog)

        self.plot_button = QPushButton("Построить график")
        self.plot_button.clicked.connect(self.plot_data)

        self.addColumnButton = QPushButton('Добавить столбец')
        self.addRowButton = QPushButton('Добавить колонку')

        self.addColumnButton.clicked.connect(self.add_column)
        self.addRowButton.clicked.connect(self.add_row)

        self.save_button = QPushButton("Сохранить таблицу")
        self.save_button.clicked.connect(self.save_data)

        # Таблица
        self.table_view = QTableView()
        self.table_view.setSortingEnabled(True)
        self.table_view.setSelectionBehavior(QTableView.SelectColumns)
        self.table_view.setAcceptDrops(True)
        self.table_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table_view.customContextMenuRequested.connect(self.show_context_menu)

        header_view = self.table_view.horizontalHeader()
        header_view.setSectionsClickable(True)
        header_view.sectionClicked.connect(self.sortByColumn)

        # Автоматическое размещение виджетов в главном окне
        central_widget = QWidget(self)
        layout = QVBoxLayout(central_widget)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.load_button)
        button_layout.addWidget(self.search_button)
        button_layout.addWidget(self.filter_button)
        button_layout.addWidget(self.plot_button)
        button_layout.addWidget(self.addColumnButton)
        button_layout.addWidget(self.addRowButton)
        button_layout.addWidget(self.save_button)

        layout.addLayout(button_layout)
        layout.addWidget(self.table_view)

        self.setCentralWidget(central_widget)

    def dragEnterEvent(self, event):
        mime_data = event.mimeData()
        if mime_data.hasUrls() and mime_data.urls()[0].toString().endswith('.xlsx'):
            event.acceptProposedAction()
                
    def dropEvent(self, event):
        mime_data = event.mimeData()
        if mime_data.hasUrls():
            file_path = mime_data.urls()[0].toLocalFile()
            try:
                data_frame = pd.read_excel(file_path)
                data_headers = data_frame.columns.tolist()

                self.model.beginResetModel()
                self.model._data = data_frame.values.tolist()
                self.model._headers = data_headers
                self.model.endResetModel()
                self.proxy_model = QSortFilterProxyModel()
                self.proxy_model.setSourceModel(self.model)
                self.table_view.setModel(self.proxy_model)

            except Exception as e:
                QMessageBox.warning(self, "Error", str(e))

    def show_filter_dialog(self):
        search_dialog = SearchDialog(self)
        result = search_dialog.exec_()

        if result == QDialog.Accepted:
            search_term = search_dialog.search_input.text()
            self.proxy_model.setFilterRegExp(search_term)
            self.proxy_model.setFilterKeyColumn(-1) 

    def show_search_dialog(self):
            search_dialog = SearchDialog(self)
            result = search_dialog.exec_()

            if result == QDialog.Accepted:
                search_term = search_dialog.search_input.text()

                self.table_view.clearSelection()

                for row in range(self.model.rowCount()):
                    for col in range(self.model.columnCount()):
                        index = self.model.index(row, col)
                        item_text = index.data(Qt.DisplayRole).lower()

                        if search_term.lower() == item_text:
                            self.table_view.selectionModel().select(index, QItemSelectionModel.Select)

    def show_context_menu(self, pos):
        context_menu = QMenu(self)

        add_column_action = QAction('Добавить столбец', self)
        add_row_action = QAction('Добавить строку', self)
        delete_row_action = QAction('Удалить столбец', self)
        delete_column_action = QAction('Удалить строку', self)

        delete_row_action.triggered.connect(self.delete_row)
        delete_column_action.triggered.connect(self.delete_column)

        add_column_action.triggered.connect(self.add_column_below)
        add_row_action.triggered.connect(self.add_row_below)

        context_menu.addAction(add_column_action)
        context_menu.addAction(add_row_action)
        context_menu.addAction(delete_row_action)
        context_menu.addAction(delete_column_action)

        context_menu.exec_(self.table_view.mapToGlobal(pos))

    def add_column_below(self):
        current_column = self.table_view.currentIndex().column() if self.table_view.currentIndex().isValid() else 0
        column_name, ok = QInputDialog.getText(self, 'Добавить столбец', 'Введите название столбца:')
        if ok and column_name:
            self.model.beginInsertColumns(QModelIndex(), current_column + 1, current_column + 1)
            self.model._headers.insert(current_column + 1, column_name)
            for row in self.model._data:
                row.insert(current_column + 1, '')
            self.model.endInsertColumns()

    def add_row_below(self):
        current_row = self.table_view.currentIndex().row() if self.table_view.currentIndex().isValid() else 0
        self.model.beginInsertRows(QModelIndex(), current_row + 1, current_row + 1)
        new_row = [''] * len(self.model._headers)
        self.model._data.insert(current_row + 1, new_row)
        self.model.endInsertRows()
    
    def delete_row(self):
        current_row = self.table_view.currentIndex().row() if self.table_view.currentIndex().isValid() else 0
        if 0 <= current_row < self.model.rowCount():
            self.model.beginRemoveRows(QModelIndex(), current_row, current_row)
            del self.model._data[current_row]
            self.model.endRemoveRows()

    def delete_column(self):
        current_column = self.table_view.currentIndex().column() if self.table_view.currentIndex().isValid() else 0
        if 0 <= current_column < self.model.columnCount():
            self.model.beginRemoveColumns(QModelIndex(), current_column, current_column)
            del self.model._headers[current_column]
            for row in self.model._data:
                del row[current_column]
            self.model.endRemoveColumns()

    def add_column(self):
        self.model.addColumn()

    def add_row(self):
        self.model.addRow()

    def add_row(self):
        self.model.addRow()


    def sortByColumn(self, logical_index):
        # Get the order of the current sorting
        current_order = self.table_view.horizontalHeader().sortIndicatorOrder()

        # Reverse the order if already sorted by the clicked column
        current_order = Qt.DescendingOrder if current_order == Qt.AscendingOrder else Qt.AscendingOrder
        print(f"current Order {current_order}")
        # Sort the model by the clicked column
        self.table_view.model().sort(logical_index, current_order)


    def load_data(self):
        """
        Загрузка данных из файла формата .xlsx.
        """
        file_dialog = QFileDialog(self)
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        file_dialog.setNameFilter("Файлы .xlsx (*.xlsx)")

        if file_dialog.exec():
            file_path = file_dialog.selectedFiles()[0]

            try:
                data_frame = pd.read_excel(file_path)
                data_headers = data_frame.columns.tolist()

                self.model = ExcelTableModel(data_frame.values.tolist(),data_headers)
                self.proxy_model = QSortFilterProxyModel()
                self.proxy_model.setSourceModel(self.model)

                self.table_view.setModel(self.proxy_model)

            except Exception as e:
                QMessageBox.warning(self, "Ошибка", str(e))

    def sort_data(self):
        """
        Сортировка данных по выбранным столбцам.
        """
        selected_columns = self.table_view.selectionModel().selectedColumns()
        if(selected_columns != None):
            model = self.table_view.model()
            sort_keys = [model.headerData(column, Qt.Horizontal) for column in selected_columns]

            model.sort(sort_keys, Qt.AscendingOrder)

    def search_data(self):
        """
        Поиск информации в выбранном столбце.
        """
        selected_columns = self.table_view.selectionModel().selectedColumns()

        if len(selected_columns) > 0:
            model = self.table_view.model()
            search_column = model.headerData(selected_columns[0], Qt.Horizontal)

            text, ok = QInputDialog.getText(self, "Поиск", "Введите значение для поиска:")

            if ok:
                result = model[model[search_column].astype(str).str.contains(text)]
                if len(result) > 0:
                    self.table_view.setModel(TableModel(result))
                else:
                    QMessageBox.information(self, "Поиск", "Совпадений не найдено.")

        else:
            QMessageBox.warning(self, "Внимание", "Выберите столбец для поиска.")

    def plot_data(self):
        """
        Построение графика на основе выбранных данных.
        """
        selected_columns = self.table_view.selectionModel().selectedColumns()

        if len(selected_columns) >= 1:
            model = self.table_view.model()

            # Create subplots
            fig, (ax_lines, ax_hist, ax_scatter) = plt.subplots(1, 3, figsize=(15, 5))

            for column in selected_columns:
                column_index = column.column()

                key = model.headerData(column_index, Qt.Horizontal)
                column_data = [model.data(model.index(row, column_index)) for row in range(model.rowCount())]

                # Line plot
                ax_lines.plot(column_data, label=f'{key}')
                ax_lines.set_title('Line Plot')
                ax_lines.set_xlabel('Index')
                ax_lines.set_ylabel('Value')

                # Histogram
                ax_hist.hist(column_data, bins='auto', alpha=0.7, label=f'{key}')
                ax_hist.set_title('Histogram')
                ax_hist.set_xlabel('Value')
                ax_hist.set_ylabel('Frequency')

                # Scatter plot
                ax_scatter.scatter(range(len(column_data)), column_data, label=f'{key}')
                ax_scatter.set_title('Scatter Plot')
                ax_scatter.set_xlabel('Index')
                ax_scatter.set_ylabel('Value')

            # Add legends to subplots
            ax_lines.legend()
            ax_hist.legend()
            ax_scatter.legend()

            # Adjust layout
            plt.tight_layout()

            # Show figures
            plt.show()
        else:
            QMessageBox.warning(self, "Внимание", "Выберите столбец для построения графика.")

    def add_data(self):
        """
        Добавление данных в таблицу.
        """
        model = self.table_view.model()
        rows, cols = model.shape

        # Открытие диалогового окна для ввода новых данных
        dialog = QDialog(self)
        dialog.setWindowTitle("Добавить данные")

        layout = QVBoxLayout(dialog)

        line_edits = []
        for i in range(cols):
            line_edit = QLineEdit()
            line_edits.append(line_edit)

            layout.addWidget(line_edit)

        ok_button = QPushButton("ОК")
        ok_button.clicked.connect(lambda: self.add_data_to_table(line_edits, dialog))
        layout.addWidget(ok_button)

        dialog.setLayout(layout)
        dialog.exec()

    def add_data_to_table(self, line_edits, dialog):
        """
        Функция добавления новых данных в таблицу.
        """
        model = self.table_view.model()

        new_data = {}
        for i, line_edit in enumerate(line_edits):
            new_data[model.headerData(i, Qt.Horizontal)] = [line_edit.text()]

        new_row = pd.DataFrame.from_dict(new_data)
        model = pd.concat([model, new_row], ignore_index=True)

        self.table_view.setModel(ExcelTableModel(model))

        dialog.close()


    def save_data(self):
        """
        Сохранение текущей таблицы в файл формата .xlsx.
        """
        file_dialog = QFileDialog(self)
        file_dialog.setFileMode(QFileDialog.AnyFile)
        file_dialog.setDefaultSuffix("xlsx")
        file_dialog.setNameFilter("Файлы .xlsx (*.xlsx)")

        if file_dialog.exec():
            file_path = file_dialog.selectedFiles()[0]

            try:
                model = self.model
                model.saveToExcel(file_path)

            except Exception as e:
                QMessageBox.warning(self, "Ошибка", str(e))


if __name__ == "__main__":
    app = QApplication(sys.argv)

    font = QFont("Arial", 12)
    app.setFont(font)

    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec())
