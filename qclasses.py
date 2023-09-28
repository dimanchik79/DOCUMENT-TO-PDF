from PyQt5 import QtCore, QtWidgets, QtGui
import comtypes.client
from PyQt5.QtGui import QColor
from win32com import client
import webbrowser


def string_error(err):
    text = ''
    for txt in err:
        text += f"{txt}"
    return text


class Ui_MainWindow(object):
    def __init__(self):
        self.file_name = None
        self.directory = None
        self.lbl_path_add = None
        self.lbl_files_add = None
        self.list_pdfs = None
        self.list_docs = None
        self.btn_add_dir = None
        self.btn_add_doc = None
        self.centralwidget = None
        self.btn_convert = None
        self.progress = None

        self.doc_to_convert = {}
        self.lists = None

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(790, 622)
        MainWindow.setWindowIcon(QtGui.QIcon("1.png"))

        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.btn_add_doc = QtWidgets.QPushButton(self.centralwidget)
        self.btn_add_doc.setGeometry(QtCore.QRect(30, 10, 291, 51))
        self.btn_add_doc.clicked.connect(self.get_add_files)

        self.btn_add_dir = QtWidgets.QPushButton(self.centralwidget)
        self.btn_add_dir.setGeometry(QtCore.QRect(470, 10, 291, 51))
        self.btn_add_dir.clicked.connect(self.add_folder)

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
        self.btn_convert.clicked.connect(self.convert)

        self.progress = QtWidgets.QProgressBar(self.centralwidget)
        self.progress.setGeometry(QtCore.QRect(355, 350, 91, 23))
        self.progress.setProperty("value", 0)

        self.list_docs = QtWidgets.QListWidget(self.centralwidget)
        self.list_docs.setGeometry(QtCore.QRect(10, 70, 331, 471))
        font = QtGui.QFont()
        font.setFamily("Calibri")
        self.list_docs.setFont(font)
        self.list_docs.setViewMode(QtWidgets.QListView.ListMode)
        self.list_docs.doubleClicked.connect(self.ondoubleclick_docum)

        self.list_pdfs = QtWidgets.QListWidget(self.centralwidget)
        self.list_pdfs.setGeometry(QtCore.QRect(450, 70, 331, 471))
        font = QtGui.QFont()
        font.setFamily("Calibri")
        self.list_pdfs.setFont(font)
        self.list_pdfs.setViewMode(QtWidgets.QListView.ListMode)
        self.list_pdfs.doubleClicked.connect(self.ondoubleclick_pdfs)

        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "WORD DOCUMENT TO PDF CONVERTER"))
        self.btn_add_doc.setText(_translate("MainWindow", "Add DOC, DOCX, XLS, XLSX files"))
        self.btn_add_dir.setText(_translate("MainWindow", "Add directory for convert"))
        self.lbl_files_add.setText(_translate("MainWindow", "Files not added"))
        self.lbl_path_add.setText(_translate("MainWindow", "Directotry not added"))
        self.btn_convert.setText(_translate("MainWindow", "CONVERT"))

    def get_add_files(self):
        files = QtWidgets.QFileDialog.getOpenFileNames(filter="*.doc *.docx *.xls *xlsx")[0]
        if not files:
            return
        for file in files:
            self.doc_to_convert[file] = file[file.rindex("/") + 1:]
        for key, value in self.doc_to_convert.items():
            self.list_docs.addItem(value)
        self.lbl_files_add.setText(f"{len(self.doc_to_convert)} files selected")
        self.list_docs.setFocus()

    def add_folder(self):
        self.directory = QtWidgets.QFileDialog.getExistingDirectory()
        self.lbl_path_add.setText(self.directory)

    def convert(self):
        error, index, index_path, text_error, err = [], 0, 0, "", True
        if self.directory is None:
            text_error = "Add directory when convert"
        elif self.doc_to_convert == {}:
            text_error = "Add files for convert"
        else:
            err = False
        if err:
            err_msg = QtWidgets.QMessageBox()
            err_msg.setWindowIcon(QtGui.QIcon("1.png"))
            err_msg.setWindowTitle("ERROR")
            err_msg.setText(text_error)
            err_msg.exec_()
            return

        self.btn_convert.setText("WAIT")
        self.progress.setMinimum(0)
        self.progress.setMaximum(len(self.doc_to_convert))
        for path, file in self.doc_to_convert.items():
            try:
                file_path = f"{self.directory}/{file[:file.rindex('.')]}.pdf"
                path = path.replace("/", chr(92))
                file_path = file_path.replace("/", chr(92))
                if file[file.rindex("."):] in [".doc", ".docx"]:
                    word = comtypes.client.CreateObject('Word.Application')
                    doc = word.Documents.Open(path)
                    doc.SaveAs(file_path, FileFormat=17)
                    doc.Close()
                    word.Quit()

                elif file[file.rindex("."):] in [".xls", ".xlsx"]:
                    excel = client.Dispatch("Excel.Application")
                    sheets = excel.Workbooks.Open(path)
                    work_sheets = sheets.Worksheets[0]
                    work_sheets.ExportAsFixedFormat(0, file_path)
                self.list_pdfs.addItem(f"{file[:file.rindex('.')]}.pdf")
                index_path += 1
                index += 1
                self.progress.setValue(index)
            except Exception:
                self.list_pdfs.addItem(f"{file} - error")
                self.list_pdfs.item(index).setForeground(QColor('red'))
                self.progress.setValue(index)
                index += 1
                error.append(f"{file}\n")
        if error:
            text_error = f"{string_error(error)}dont converted\n{index_path} files converted"
        else:
            text_error = "All files succesfuly converted!"
        if text_error != "":
            err_msg = QtWidgets.QMessageBox()
            err_msg.setWindowIcon(QtGui.QIcon("1.png"))
            err_msg.setWindowTitle("WARRNING")
            err_msg.setText(text_error)
            err_msg.exec_()

        self.btn_convert.setText("CONVERT")
        self.progress.setValue(0)

    def ondoubleclick_docum(self):
        for key, word in self.doc_to_convert.items():
            if word == self.list_docs.currentItem().text():
                self.file_name = key
        index = self.list_docs.currentIndex().row()
        self.list_docs.item(index).setForeground(QColor('green'))
        webbrowser.open(self.file_name, new=0, autoraise=True)

    def ondoubleclick_pdfs(self):
        index = self.list_pdfs.currentIndex().row()
        self.list_pdfs.item(index).setForeground(QColor('green'))
        self.file_name = f"{self.directory}/{self.list_pdfs.currentItem().text()}"
        webbrowser.open(self.file_name, new=0, autoraise=True)
