from PyQt5 import QtCore, QtGui, QtWidgets



class ValidatedItemDelegate(QtWidgets.QItemDelegate):
    def createEditor(self, widget, option, index):
        if not index.isValid():
            return 0
        if index.column() == 7 or index.column() == 8:
            editor = QtWidgets.QLineEdit(widget)
            validator = QtGui.QRegExpValidator(QtCore.QRegExp("\d{2} [a-zA-Zа-яА-Я]{0,10} \d{4}"), editor)
            editor.setValidator(validator)
            return editor

        return super(ValidatedItemDelegate, self).createEditor(widget, option, index)

class DateEditDelegate(QtWidgets.QItemDelegate):
    def __init__(self, parent = None):
        super(QtWidgets.QItemDelegate, self).__init__(parent)

    def createEditor(self, parent, option, index):
        dateEdit = QtWidgets.QDateEdit(parent)
        dateEdit.setDisplayFormat("dd.MM.yyyy")
        return dateEdit
