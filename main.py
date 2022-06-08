# from PyQt6.QtWidgets import  QApplication, QMainWindow
# from PyQt6 import QtCore, QtGui, QtWidgets
# from PyQt6.QtCore import Qt

import sys
from datetime import datetime

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import Qt


class TableModel(QtCore.QAbstractTableModel):
    def __init__(self, data):
        super(TableModel, self).__init__()
        self._data = data

    headerNames = ['Институт', 'Направление', 'Профиль', 'Семестр', 'Вид практики', 'Тип практики', 'Трудоемкость',
                   'Дата начала', 'Дата окончания', 'Компетенции']

    def data(self, index, role):
        if role == Qt.DisplayRole or role == Qt.EditRole:
            value = self._data[index.row()][index.column()]

            if isinstance(value, datetime):
                # Render time to YYY-MM-DD.
                return value.strftime("%Y-%m-%d")

            if isinstance(value, float):
                # Render float to 2 dp
                return "%.2f" % value

            if isinstance(value, str):
                # Render strings with quotes
                return '%s' % value

                # Default (anything not captured above: e.g. int)
            return value

    def rowCount(self, index):
        # The length of the outer list.
        return len(self._data)

    def columnCount(self, index):
        # The following takes the first sub-list, and returns
        # the length (only works if all rows are an equal length)
        return len(self._data[0])

    def headerData(self, section, orientation, role):
        # section is the index of the column/row.
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return str(self.headerNames[section])

    def setData(self, index, value, role):
        if role == Qt.EditRole:
            if index.column() == 7 or index.column() == 8:
                self._data[index.row()][index.column()] = value
                return True
            return False
        return False

    def flags(self, index):
        return Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable

class Ui_MainWindow(object):
    ## Data for table
    data = [
        ['', '', '', '', '', '', '', '', '', '']
    ]

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 600)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName("verticalLayout")
        self.fisrtStageTable = QtWidgets.QTableView(self.centralwidget)
        self.fisrtStageTable.setObjectName("fisrtStageTable")

        self.tableModel = TableModel(self.data)
        self.fisrtStageTable.setModel(self.tableModel)

        self.verticalLayout.addWidget(self.fisrtStageTable)
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setObjectName("label")
        self.verticalLayout.addWidget(self.label)
        self.continueBtn = QtWidgets.QPushButton(self.centralwidget)
        self.continueBtn.setObjectName("continueBtn")
        self.verticalLayout.addWidget(self.continueBtn)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 22))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label.setText(_translate("MainWindow", "informLabel"))
        self.continueBtn.setText(_translate("MainWindow", "Подтвердить"))

    def addRecordToTable(self, record):
        lenghtList = len(record)
        if lenghtList < 10:
            return 0
        if self.data[0][0] == '':
            self.data.clear()
        self.data.append(record)
        self.fisrtStageTable.model().layoutChanged.emit()


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    ui.addRecordToTable(["text", "text", "text", "text", "text", "text", "text", "text", "text", "text"])
    MainWindow.show()
    sys.exit(app.exec())
