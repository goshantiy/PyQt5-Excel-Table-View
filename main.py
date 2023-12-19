
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableView, QDialog, QVBoxLayout, QHBoxLayout, QMessageBox, QFileDialog
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex
import pandas as pd
import matplotlib.pyplot as plt
from PyQt5.QtWidgets import QPushButton, QWidget, QLineEdit, QInputDialog


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
    
    def addData(self, new_data):
        self.layoutAboutToBeChanged.emit()

        self._data.append(new_data)

        self.layoutChanged.emit()

    def printData(self):
        print(f"Data row:")
        for row in self._data:
            print(row)
        print()


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        self.setWindowTitle("Программа")
        self.resize(800, 600)

        # Кнопки главного окна
        self.load_button = QPushButton("Загрузить базу данных")
        self.load_button.clicked.connect(self.load_data)

        self.search_button = QPushButton("Поиск")
        self.search_button.clicked.connect(self.search_data)

        self.plot_button = QPushButton("Построить график")
        self.plot_button.clicked.connect(self.plot_data)

        self.add_button = QPushButton("Добавить данные")
        self.add_button.clicked.connect(self.add_data)

        self.delete_button = QPushButton("Удалить данные")
        self.delete_button.clicked.connect(self.delete_data)

        self.save_button = QPushButton("Сохранить таблицу")
        self.save_button.clicked.connect(self.save_data)

        # Таблица
        self.table_view = QTableView()
        self.table_view.setSortingEnabled(True)
        self.table_view.setSelectionBehavior(QTableView.SelectColumns)

        header_view = self.table_view.horizontalHeader()
        header_view.setSectionsClickable(True)
        header_view.sectionClicked.connect(self.sortByColumn)

        # Автоматическое размещение виджетов в главном окне
        central_widget = QWidget(self)
        layout = QVBoxLayout(central_widget)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.load_button)
        button_layout.addWidget(self.search_button)
        button_layout.addWidget(self.plot_button)
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.delete_button)
        button_layout.addWidget(self.save_button)

        layout.addLayout(button_layout)
        layout.addWidget(self.table_view)

        self.setCentralWidget(central_widget)

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

                # Создание кастомной модели данных на основе загруженного DataFrame
                self.model = ExcelTableModel(data_frame.values.tolist(),data_headers)

                self.table_view.setModel(self.model)

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

        if len(selected_columns) >= 2:
            model = self.table_view.model()

            plt.figure()
            for column in selected_columns:
                column_index = column.column()

                # Get header data for the selected column
                key = model.headerData(column_index, Qt.Horizontal)

                # Extract data for the specified column
                column_data = [model.data(model.index(row, column_index)) for row in range(model.rowCount())]

                # Plot data
                plt.plot(column_data, label=key)

            plt.legend()
            plt.show()
        else:
            QMessageBox.warning(self, "Внимание", "Выберите по крайней мере два столбца для построения графика.")

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

    def delete_data(self):
        """
        Удаление выбранных данных из таблицы.
        """
        selected_rows = self.table_view.selectionModel().selectedRows()

        if len(selected_rows) > 0:
            model = self.table_view.model()
            model.drop(selected_rows, inplace=True)

            self.table_view.setModel(ExcelTableModel(model))

        else:
            QMessageBox.warning(self, "Внимание", "Выберите строки для удаления.")

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
                model = self.table_view.model()
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
