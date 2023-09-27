from PyQt5 import QtCore, QtGui, QtWidgets
from tkinter import *
from tkinter.messagebox import showerror
from tkinter.filedialog import askopenfilenames, askdirectory


def string_error(err):
    text = ''
    for txt in err:
        text += f"{txt}"
    return text


class Ui_MainWindow(object):
    def __init__(self):
        self.lbl_path_add = None
        self.lbl_files_add = None
        self.tree_pdfs = None
        self.tree_docs = None
        self.btn_add_dir = None
        self.btn_add_doc = None
        self.centralwidget = None
        self.btn_convert = None
        self.progress = None

        self.doc_to_convert = {}
        self.index = 0
        self.index_path = 0
        self.lists = None

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 622)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.btn_add_doc = QtWidgets.QPushButton(self.centralwidget)
        self.btn_add_doc.setGeometry(QtCore.QRect(40, 10, 261, 31))
        self.btn_add_doc.setObjectName("btn_add_doc")
        self.btn_add_doc.clicked.connect(self.get_add_files)

        self.btn_add_dir = QtWidgets.QPushButton(self.centralwidget)
        self.btn_add_dir.setGeometry(QtCore.QRect(470, 10, 281, 31))
        self.btn_add_dir.setObjectName("btn_add_dir")

        self.tree_docs = QtWidgets.QListWidget(self.centralwidget)
        self.tree_docs.setGeometry(QtCore.QRect(10, 50, 331, 491))
        self.tree_docs.setObjectName("tree_docs")

        self.tree_pdfs = QtWidgets.QListView(self.centralwidget)
        self.tree_pdfs.setGeometry(QtCore.QRect(450, 50, 331, 491))
        self.tree_pdfs.setObjectName("tree_pdfs")

        self.lbl_files_add = QtWidgets.QLabel(self.centralwidget)
        self.lbl_files_add.setGeometry(QtCore.QRect(10, 550, 331, 61))
        self.lbl_files_add.setStyleSheet(" background-color: green;color: white;")
        self.lbl_files_add.setAlignment(QtCore.Qt.AlignCenter)
        self.lbl_files_add.setObjectName("lbl_files_add")

        self.lbl_path_add = QtWidgets.QLabel(self.centralwidget)
        self.lbl_path_add.setGeometry(QtCore.QRect(450, 550, 341, 61))
        self.lbl_path_add.setStyleSheet(" background-color: green;color: white;")
        self.lbl_path_add.setAlignment(QtCore.Qt.AlignCenter)
        self.lbl_path_add.setObjectName("lbl_path_add")

        self.btn_convert = QtWidgets.QPushButton(self.centralwidget)
        self.btn_convert.setGeometry(QtCore.QRect(350, 250, 93, 91))
        self.btn_convert.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.btn_convert.setStyleSheet("background-image: url(\"convert.ico\")")
        self.btn_convert.setObjectName("btn_convert")

        self.progress = QtWidgets.QProgressBar(self.centralwidget)
        self.progress.setGeometry(QtCore.QRect(350, 350, 91, 23))
        self.progress.setProperty("value", 0)
        self.progress.setObjectName("progress")


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

    def get_add_files(self):
        error = []
        files = askopenfilenames()
        for file in files:
            if file[file.rindex("."):] not in ['.docx', '.doc']:
                error.append(f"{file[file.rindex('/') + 1:]}\n")
            else:
                self.doc_to_convert[file] = file[file.rindex("/") + 1:]
        if error:
            showerror(title="WARRNING!", message=f"This files dont convert:\nRename to .docx\n{string_error(error)}")

        for key, value in self.doc_to_convert.items():
            self.tree_docs.addItem(value)

        # self.docx_label.configure(text=f"{len(self.doc_to_convert)} files selected")
        # self.listbox_docx.select_set(0, 0)
        # self.listbox_docx.focus_set()
