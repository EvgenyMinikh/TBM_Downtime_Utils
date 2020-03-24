from PyQt5 import QtWidgets, QtCore, uic
import csv
import sys
import pyodbc
from re import compile

SQL_SERVER_NAME = 'ru1rdb16.group.ad'
DATABASE_NAME = 'TBM_Downtimes'
SOURCE_CSV_FILE_FOR_LISTS = ".\\TBM_List.csv"
MAIN_UI_FORM_PATH = ".\\Main_Window_Form.ui"

SHIFT_TIME = ('', 'День', 'Ночь')
SHIFT_NUMBERS = ('', '1', '2', '3', '4')
OPERATOR_LIST = ('', 'A - Adjuster', 'M - Mechanical', 'O - Operator', 'P - PLC', 'Q - Quality Engineer', 'V - VMI Spec')

TBM_MAX_NUMBER = 49
TBM_LIST = tuple("TBM{:02d}".format(i) for i in range(1, TBM_MAX_NUMBER + 1))

conn = pyodbc.connect('Driver={SQL Server};Server=' + SQL_SERVER_NAME + ';Database=' + DATABASE_NAME + ';Trusted_Connection=yes;')


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


def write_data_into_DB(conn, db_values):
    cursor = conn.cursor()

    try:
        cursor.execute("INSERT INTO [" + DATABASE_NAME + "].[dbo].[Main]"
                                                         "([TBM Number],"
                                                         "[Defect Code],"
                                                         "[Fault Description],"
                                                         "[Idle Time],"
                                                         "[Operator],"
                                                         "[Description Notes],"
                                                         "[Date],"
                                                         "[Shift Number],"
                                                         "[Shift Time]) "
                                                         "VALUES (" + db_values + ")")
    except BaseException as e:
        main_window.plainTextEdit_Errors_Message.clear()
        main_window.plainTextEdit_Errors_Message.insertPlainText(str(e))
    else:
        conn.commit()
        main_window.plainTextEdit_Errors_Message.clear()
        main_window.plainTextEdit_Errors_Message.insertPlainText("Данные записаны")


class main_Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(main_Ui, self).__init__()
        uic.loadUi(MAIN_UI_FORM_PATH, self)
        self.setFixedSize(self.size())

        self.pushButton_Clean.clicked.connect(self.action_pushButton_Clean)
        self.pushButton_Save.clicked.connect(self.action_pushButton_Save)

        self.current_date = QtCore.QDate.currentDate()
        self.dateEdit_Date.setDate(self.current_date)

        self.comboBox_Fault_Code.activated.connect(self.change_list_values)

        FAULT_CODES_LIST = get_list_for_code_combobox(raw_data)
        FAULT_CODES_LIST.append("")
        self.comboBox_Fault_Code.addItems(sorted(FAULT_CODES_LIST))
        self.change_list_values()

        self.comboBox_Shift.addItems(SHIFT_TIME)
        self.comboBox_Shift_Number.addItems(SHIFT_NUMBERS)
        self.comboBox_TBM_number.addItems(TBM_LIST)
        self.comboBox_Operator.addItems(OPERATOR_LIST)

        self.plainTextEdit_Errors_Message.setReadOnly(True)
        self.plainTextEdit_Errors_Message.setStyleSheet("background-color: rgb(217, 217, 217);")
        self.plainTextEdit_Additional_Info.insertPlainText('-')
        self.show()

    def change_list_values(self):
        FAULT_CODES_DESCRIPTION_LIST = get_list_for_fault_description_combobox(self.comboBox_Fault_Code.currentText(),
                                                                               raw_data)
        FAULT_CODES_DESCRIPTION_LIST.append("")
        self.comboBox_Fault_Description.clear()
        self.comboBox_Fault_Description.addItems(sorted(FAULT_CODES_DESCRIPTION_LIST))

    def action_pushButton_Clean(self):
        self.current_date = QtCore.QDate.currentDate()
        self.dateEdit_Date = self.findChild(QtWidgets.QDateEdit, 'dateEdit_Date')
        self.dateEdit_Date.setDate(self.current_date)
        self.comboBox_Fault_Code.setCurrentIndex(0)
        #self.comboBox_Shift.setCurrentIndex(0)
        #self.comboBox_Shift_Number.setCurrentIndex(0)
        self.comboBox_TBM_number.setCurrentIndex(0)
        self.comboBox_Fault_Description.setCurrentIndex(0)
        self.comboBox_Operator.setCurrentIndex(0)
        self.lineEdit_Delay.clear()
        self.plainTextEdit_Additional_Info.clear()
        self.plainTextEdit_Additional_Info.insertPlainText('-')
        self.plainTextEdit_Errors_Message.clear()

    def action_pushButton_Save(self):
        date = (self.dateEdit_Date.date()).toPyDate()
        shift = self.comboBox_Shift.currentText()
        shift_number = self.comboBox_Shift_Number.currentText()
        TBM_number = self.comboBox_TBM_number.currentText()
        fault_code = self.comboBox_Fault_Code.currentText()
        fault_code_description = self.comboBox_Fault_Description.currentText()
        fault_delay = self.lineEdit_Delay.text()
        operator = self.comboBox_Operator.currentText()
        additional_info = self.plainTextEdit_Additional_Info.toPlainText()

        data_checker(shift, shift_number, TBM_number, fault_code, fault_code_description, fault_delay, operator)

        text_line = "'{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}'".format(TBM_number,
                                                                                           fault_code,
                                                                                           fault_code_description,
                                                                                           fault_delay,
                                                                                           operator,
                                                                                           additional_info,
                                                                                           date.strftime('%Y-%m-%d'),
                                                                                           shift_number,
                                                                                           shift)

        error_message = data_checker(shift, shift_number, TBM_number, fault_code, fault_code_description, fault_delay, operator).strip()
        self.plainTextEdit_Errors_Message.clear()
        self.plainTextEdit_Errors_Message.insertPlainText(error_message)

        if error_message == '':
            write_data_into_DB(conn, text_line)
            self.action_pushButton_Clean()

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
    main_window = main_Ui()
    app.exec_()
