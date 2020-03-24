from PyQt5 import QtWidgets, QtCore, uic
import csv
import sys
import pyodbc
from re import compile

# DB Connection settings
SQL_SERVER_NAME = 'ru1rdb16.group.ad'
DATABASE_NAME = 'TBM_Downtimes'
SOURCE_CSV_FILE_FOR_LISTS = ".\\TBM_List.csv"
MAIN_UI_FORM_PATH = ".\\Edit_Tool_Main_Window.ui"

# Sources for lists
SHIFT_TIME = ('День', 'Ночь')
SHIFT_NUMBERS = ('', '1', '2', '3', '4')
OPERATOR_LIST = (
'', 'A - Adjuster', 'M - Mechanical', 'O - Operator', 'P - PLC', 'Q - Quality Engineer', 'V - VMI Spec')
RECORD_TABLE_HEADERS = (
'ID', 'Номер\nстанка', 'Код', 'Описание\nсбоя', 'Длительность', 'Оператор', 'Доп.\nинформация', 'Дата', 'Номер\nсмены',
'Смена', 'ХЗ Что')

# Source data for TBM list
TBM_MAX_NUMBER = 49
TBM_LIST = tuple("TBM{:02d}".format(i) for i in range(1, TBM_MAX_NUMBER + 1))

# SQL connection string
conn = pyodbc.connect('Driver={SQL Server};Server=' + SQL_SERVER_NAME + ';Database=' + DATABASE_NAME + ';Trusted_Connection=yes;')


# This gets other lists
def read_CSV_for_lists(CSV_Path):
    raw_data = []

    with open(CSV_Path, mode='r', newline='') as csvfile:
        csv_reader = csv.DictReader(csvfile, delimiter=";")

        for row in csv_reader:
            error_code = row["ErrorCode"]
            error_code_main = row["ErrorCodeMain"]
            description_main = row["DescriptionMain"]
            description_main_with_code = row["DescriptionMainWithCode"]
            code_and_description = row["CodeAndDescription"]
            error_description = row["ErrorDescription"]
            error_group = row["ErrorGroup"]

            line = [error_code, error_code_main, description_main, description_main_with_code, code_and_description,
                    error_description, error_group]
            raw_data.append(line)

    return raw_data


raw_data = read_CSV_for_lists(SOURCE_CSV_FILE_FOR_LISTS)


def get_list_for_code_combobox(raw_data):
    result = [i[3] for i in raw_data]
    return sorted(list(set(result)))


def get_list_for_fault_description_combobox(fault_code, raw_data):
    result = [i[4] for i in raw_data if i[3] == fault_code]
    return result


def data_checker(shift, shift_number, TBM_number, fault_code, fault_code_description, fault_delay, operator):
    result = ''

    if shift == '':
        result += 'Поле "Смена" не может быть пустым\n'

    if shift_number == '':
        result += 'Поле "Номер смены" не может быть пустым\n'

    if TBM_number == '':
        result += 'Поле "Станок" не может быть пустым\n'

    if fault_code == '':
        result += 'Поле "Код" не может быть пустым\n'

    if fault_code_description == '':
        result += 'Поле "Описание сбоя" не может быть пустым\n'

    if (fault_delay is None) or (fault_delay == ''):
        result += 'Поле "Длительность сбоя" не может быть пустым\n'
    else:
        pattern = compile("^\d+$")
        if not pattern.match(fault_delay):
            result += 'Поле "Длительность сбоя" должно содержать только цифры\n'

    if operator == '':
        result += 'Поле "Оператор" не может быть пустым\n'

    return result


# Callable for deleting from DB
def delete_from_DB(conn, id):
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM [TBM_Downtimes].[dbo].[Main] WHERE [ID] = '" + id + "'")
        conn.commit()
    except BaseException as e:
        main_window.plainTextEdit_Errors_Message.insertPlainText(e)
    else:
        main_window.plainTextEdit_Errors_Message.insertPlainText("Запись удалена")


# Callable for updating data
def update_data_in_DB(conn, id, list_data):
    text_line = "[TBM Number] = '{0}', " \
                "[Defect Code] = '{1}', " \
                "[Fault Description] = '{2}', " \
                "[Idle Time] = '{3}', " \
                "[Operator] = '{4}', " \
                "[Description Notes] = '{6}', " \
                "[Date] = '{5}', " \
                "[Shift Number] = '{7}', " \
                "[Shift Time] = '{8}'".format(*list_data)

    cursor = conn.cursor()

    try:
        cursor.execute("UPDATE [TBM_Downtimes].[dbo].[Main]"
                        "SET " + text_line +
                        "WHERE [ID] = '" + id + "'")
        conn.commit()
    except BaseException as e:
        main_window.plainTextEdit_Errors_Message.insertPlainText(str(e))
    else:
        main_window.plainTextEdit_Errors_Message.insertPlainText("Запись обновлена")


def read_data_from_DB(conn, date_to_select, shift_to_select):
    cursor = conn.cursor()
    main_window.plainTextEdit_Errors_Message.clear()
    cursor.execute("SELECT * FROM TBM_Downtimes.dbo.Main WHERE [Date] = '" + date_to_select + "' AND [Shift Time] = '" + shift_to_select + "'")

    raw_data = []
    for row in cursor:
        raw_data.append(row)

    return raw_data


def populate_table_with_data(raw_data):
    main_window.tableWidget_Records.setRowCount(len(raw_data))

    for i, row in enumerate(raw_data):
        for j, col in enumerate(row):
            item = QtWidgets.QTableWidgetItem(str(col))
            main_window.tableWidget_Records.setItem(i, j, item)

    main_window.tableWidget_Records.resizeColumnsToContents()
    main_window.tableWidget_Records.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)


class main_UI(QtWidgets.QMainWindow):
    def __init__(self):
        super(main_UI, self).__init__()
        uic.loadUi(MAIN_UI_FORM_PATH, self)
        self.setFixedSize(self.size())

        # Set button handlers
        self.pushButton_Get_Records.clicked.connect(self.action_pushButton_Get_Records)
        self.pushButton_Update_Record.clicked.connect(self.action_pushButton_Update_Record)
        self.pushButton_Delete_Record.clicked.connect(self.action_pushButton_Delete_Record)
        self.tableWidget_Records.clicked.connect(self.action_table_click)

        # Set date for date choosers
        self.current_date = QtCore.QDate.currentDate()
        self.dateEdit_Date.setDate(self.current_date)
        self.dateEdit_Selector.setDate(self.current_date)

        self.comboBox_Fault_Code.activated.connect(self.change_list_values)

        self.FAULT_CODES_LIST = get_list_for_code_combobox(raw_data)
        self.FAULT_CODES_LIST.append("")
        self.comboBox_Fault_Code.addItems(sorted(self.FAULT_CODES_LIST))
        self.change_list_values()

        # Set list values
        self.comboBox_Shift_for_Query.addItems(SHIFT_TIME)
        self.comboBox_Shift.addItems(SHIFT_TIME)
        self.comboBox_Shift_Number.addItems(SHIFT_NUMBERS)
        self.comboBox_TBM_number.addItems(TBM_LIST)
        self.comboBox_Operator.addItems(OPERATOR_LIST)

        # Preparing status text area
        self.plainTextEdit_Errors_Message.setReadOnly(True)
        self.plainTextEdit_Errors_Message.setStyleSheet("background-color: rgb(217, 217, 217);")

        # Preparing QTable Widget
        self.tableWidget_Records.setColumnCount(len(RECORD_TABLE_HEADERS))
        self.tableWidget_Records.setHorizontalHeaderLabels(RECORD_TABLE_HEADERS)
        self.tableWidget_Records.resizeColumnsToContents()

        self.show()

    def change_list_values(self):
        global FAULT_CODES_DESCRIPTION_LIST
        FAULT_CODES_DESCRIPTION_LIST = get_list_for_fault_description_combobox(self.comboBox_Fault_Code.currentText(),
                                                                               raw_data)
        FAULT_CODES_DESCRIPTION_LIST.append("")
        self.comboBox_Fault_Description.clear()
        self.comboBox_Fault_Description.addItems(sorted(FAULT_CODES_DESCRIPTION_LIST))

    def action_pushButton_Get_Records(self):
        date = (self.dateEdit_Selector.date()).toPyDate().strftime('%Y-%m-%d')
        shift_time = self.comboBox_Shift_for_Query.currentText()
        raw_data = read_data_from_DB(conn, date, shift_time)
        populate_table_with_data(raw_data)

    def action_table_click(self):
        indexes = self.tableWidget_Records.selectionModel().selectedIndexes()
        row_data = []

        for index in indexes:
            cell_value = index.sibling(index.row(), index.column()).data()
            row_data.append(cell_value)

        global record_id
        record_id = row_data[0]
        record_TBM_number = row_data[1]
        try:
            self.comboBox_TBM_number.setCurrentIndex(TBM_LIST.index(record_TBM_number))
        except BaseException as e:
            print(e)

        record_fault_code = row_data[2]
        try:
            self.comboBox_Fault_Code.setCurrentIndex(self.FAULT_CODES_LIST.index(record_fault_code) + 1)
            self.change_list_values()
        except BaseException as e:
            print(e)

        record_fault_code_description = row_data[3]
        try:
            self.comboBox_Fault_Description.setCurrentIndex(
                FAULT_CODES_DESCRIPTION_LIST.index(record_fault_code_description) + 1)
        except BaseException as e:
            print(e)

        record_fault_delay = row_data[4]
        self.lineEdit_Fault_Delay.setText(record_fault_delay)

        record_operator = row_data[5]
        try:
            self.comboBox_Operator.setCurrentIndex(OPERATOR_LIST.index(record_operator))
        except BaseException as e:
            print(e)

        record_additional_info = row_data[6]
        self.lineEdit_Additional_Info.setText(record_additional_info)

        record_date = row_data[7]
        record_date_splitted = record_date.split('-')
        year = int(record_date_splitted[0])
        month = int(record_date_splitted[1])
        day = int(record_date_splitted[2])
        record_date = QtCore.QDate(year, month, day)
        self.dateEdit_Date.setDate(record_date)

        record_shift_number = row_data[8]
        try:
            self.comboBox_Shift_Number.setCurrentIndex(SHIFT_NUMBERS.index(record_shift_number))
        except BaseException as e:
            print(e)

        record_shift = row_data[9]
        try:
            self.comboBox_Shift.setCurrentIndex(SHIFT_TIME.index(record_shift))
        except BaseException as e:
            print(e)

    def action_pushButton_Delete_Record(self):
        if self.checkBox_Confirm_Delete.isChecked():
            print("Delete " + record_id)
            delete_from_DB(conn, record_id)
            self.action_pushButton_Get_Records()
            self.checkBox_Confirm_Delete.setChecked(False)
        else:
            self.plainTextEdit_Errors_Message.insertPlainText("Для удаления нужно поставить флажок")

    def action_pushButton_Update_Record(self):
        try:
            date = ((self.dateEdit_Date.date()).toPyDate()).strftime('%Y-%m-%d')
            shift = self.comboBox_Shift.currentText()
            shift_number = self.comboBox_Shift_Number.currentText()
            TBM_number = self.comboBox_TBM_number.currentText()
            fault_code = self.comboBox_Fault_Code.currentText()
            fault_code_description = self.comboBox_Fault_Description.currentText()
            fault_delay = self.lineEdit_Fault_Delay.text()
            operator = self.comboBox_Operator.currentText()
            additional_info = self.lineEdit_Additional_Info.text()

            error_message = data_checker(shift, shift_number, TBM_number, fault_code, fault_code_description, fault_delay, operator)
            self.plainTextEdit_Errors_Message.clear()
            self.plainTextEdit_Errors_Message.insertPlainText(error_message)
            list_data = [TBM_number, fault_code, fault_code_description, fault_delay, operator, date, additional_info, shift_number, shift]

            if error_message == '':
                update_data_in_DB(conn, record_id, list_data)
                self.action_pushButton_Get_Records()
                self.plainTextEdit_Errors_Message.insertPlainText("Запись обновлена")
        except BaseException as e:
            print(e)

    def closeEvent(self, event):
        reply = QtWidgets.QMessageBox.question(self, 'Quit',
                                               "Закрыть окно?", QtWidgets.QMessageBox.Yes |
                                               QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No)

        if reply == QtWidgets.QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    main_window = main_UI()
    app.exec_()
