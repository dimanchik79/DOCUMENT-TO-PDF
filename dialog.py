# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'dialog.ui'
#
# Created by: PyQt5 UI code generator 5.15.9
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(790, 622)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("icon.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.btn_add_doc = QtWidgets.QPushButton(self.centralwidget)
        self.btn_add_doc.setGeometry(QtCore.QRect(30, 10, 291, 51))


        self.btn_add_dir = QtWidgets.QPushButton(self.centralwidget)
        self.btn_add_dir.setGeometry(QtCore.QRect(470, 10, 291, 51))

        self.lbl_files_add = QtWidgets.QLabel(self.centralwidget)
        self.lbl_files_add.setGeometry(QtCore.QRect(10, 550, 331, 61))
        self.lbl_files_add.setStyleSheet(" background-color: green;color: white;border-radius: 5px;")
        self.lbl_files_add.setAlignment(QtCore.Qt.AlignCenter)

        self.lbl_path_add = QtWidgets.QLabel(self.centralwidget)
        self.lbl_path_add.setGeometry(QtCore.QRect(450, 550, 331, 61))
        self.lbl_path_add.setStyleSheet(" background-color: green;color: white;border-radius: 5px;")
        self.lbl_path_add.setAlignment(QtCore.Qt.AlignCenter)

        self.btn_convert = QtWidgets.QPushButton(self.centralwidget)
        self.btn_convert.setGeometry(QtCore.QRect(350, 250, 93, 91))
        self.btn_convert.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.btn_convert.setStyleSheet("background-image: url(\"convert.ico\")")


        self.progress = QtWidgets.QProgressBar(self.centralwidget)
        self.progress.setGeometry(QtCore.QRect(350, 350, 91, 23))
        self.progress.setProperty("value", 0)

        self.list_doc = QtWidgets.QListWidget(self.centralwidget)
        self.list_doc.setGeometry(QtCore.QRect(10, 70, 331, 471))
        font = QtGui.QFont()
        font.setFamily("Calibri")
        self.list_doc.setFont(font)
        self.list_doc.setViewMode(QtWidgets.QListView.ListMode)

        self.list_pdfs = QtWidgets.QListWidget(self.centralwidget)
        self.list_pdfs.setGeometry(QtCore.QRect(450, 70, 331, 471))
        font = QtGui.QFont()
        font.setFamily("Calibri")
        self.list_pdfs.setFont(font)
        self.list_pdfs.setViewMode(QtWidgets.QListView.ListMode)

        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "WORD DOCUMENT TO PDF CONVERTER"))
        self.btn_add_doc.setText(_translate("MainWindow", "Add DOC, DOCX files"))
        self.btn_add_dir.setText(_translate("MainWindow", "Add directory for convert"))
        self.lbl_files_add.setText(_translate("MainWindow", "Files not added"))
        self.lbl_path_add.setText(_translate("MainWindow", "Directotry not added"))
        self.btn_convert.setText(_translate("MainWindow", "CONVERT"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())