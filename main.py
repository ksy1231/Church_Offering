# Icon made by Freepik from www.flaticon.com

# -*-coding: UTF-8-*-

import os
import pandas as pd
import sys
from PyQt5 import QtCore
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *

if hasattr(Qt, 'AA_EnableHighDpiScaling'):
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)

if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)


class typeDialog(QDialog):
    def __init__(self, files, parent=None):
        super(typeDialog, self).__init__(parent)
        self.setupUI()

        self.files = files

    def setupUI(self):
        self.resize(400, 300)
        self.verticalLayoutWidget = QWidget(self)
        self.verticalLayoutWidget.setGeometry(QRect(60, 20, 280, 190))
        self.verticalLayout = QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)

        self.titheButton = QPushButton(self.verticalLayoutWidget)
        font = QFont()
        font.setPointSize(20)
        self.titheButton.setFont(font)
        self.verticalLayout.addWidget(self.titheButton)
        self.titheButton.clicked.connect(self.tithe_clicked)

        self.thanksOfferingButton = QPushButton(self.verticalLayoutWidget)
        font = QFont()
        font.setPointSize(20)
        self.thanksOfferingButton.setFont(font)
        self.verticalLayout.addWidget(self.thanksOfferingButton)
        self.thanksOfferingButton.clicked.connect(self.thanksOffering_clicked)

        self.offeringButton = QPushButton(self.verticalLayoutWidget)
        font = QFont()
        font.setPointSize(20)
        self.offeringButton.setFont(font)
        self.verticalLayout.addWidget(self.offeringButton)
        self.offeringButton.clicked.connect(self.offering_clicked)

        self.exitButton = QPushButton(self)
        font = QFont()
        font.setPointSize(20)
        self.exitButton.setFont(font)
        self.exitButton.setGeometry(QtCore.QRect(150, 240, 110, 30))
        self.exitButton.clicked.connect(self.close)

        self.retranslateUi(self)
        QMetaObject.connectSlotsByName(self)

    def tithe_clicked(self):
        df = pd.DataFrame()
        for f in self.files:
            data = pd.read_excel(f, sheet_name="십일조", header=None)
            df = df.append(data)
        df = df.groupby(0).size().reset_index(name="1")
        df.to_excel('십일조 합산 결과.xlsx', sheet_name='십일조', index=False, header=None)

    def thanksOffering_clicked(self):
        df = pd.DataFrame()
        for f in self.files:
            data = pd.read_excel(f, sheet_name="감사헌금", header=None)
            df = df.append(data)
        df = df.groupby(0).size().reset_index(name="1")
        df.to_excel('감사헌금 합산 결과.xlsx', sheet_name='감사헌금', index=False, header=None)

    def offering_clicked(self):
        df = pd.DataFrame()
        for f in self.files:
            data = pd.read_excel(f, sheet_name="주정헌금", header=None)
            df = df.append(data)
        df = df.groupby(0).size().reset_index(name="1")
        df.to_excel('주정헌금 합산 결과.xlsx', sheet_name='주정헌금', index=False, header=None)

    def retranslateUi(self, Dialog):
        _translate = QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "재정 프로그램"))
        self.titheButton.setText(_translate("Dialog", "십일조"))
        self.thanksOfferingButton.setText(_translate("Dialog", "감사헌금"))
        self.offeringButton.setText(_translate("Dialog", "주정헌금"))
        self.exitButton.setText(_translate("Dialog", "Exit"))


class MainWindow(QWidget):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setupUI()

        self.files_xls = None

    def setupUI(self):
        scriptDir = os.path.dirname(os.path.realpath(__file__))
        self.setWindowIcon(QIcon(scriptDir + os.path.sep + 'church.ico'))
        self.resize(400, 250)
        # self.openFileNamesDialog()
        # self.show()
        self.horizontalLayoutWidget = QWidget(self)
        self.horizontalLayoutWidget.setGeometry(QRect(20, 40, 360, 80))
        self.horizontalLayout = QHBoxLayout(self.horizontalLayoutWidget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)

        self.label = QLabel(self.horizontalLayoutWidget)
        font = QFont()
        font.setPointSize(20)
        self.label.setFont(font)
        self.horizontalLayout.addWidget(self.label)

        self.lineEdit = QLineEdit(self.horizontalLayoutWidget)
        font = QFont()
        font.setPointSize(20)
        self.lineEdit.setFont(font)
        self.horizontalLayout.addWidget(self.lineEdit)
        self.lineEdit.textChanged.connect(self.getFolderPath)

        self.toolButton = QToolButton(self.horizontalLayoutWidget)
        font = QFont()
        font.setPointSize(20)
        self.toolButton.setFont(font)
        self.horizontalLayout.addWidget(self.toolButton)
        self.toolButton.clicked.connect(self.toolButton_clicked)

        self.horizontalLayoutWidget_2 = QWidget(self)
        self.horizontalLayoutWidget_2.setGeometry(QRect(180, 130, 200, 80))
        self.horizontalLayout_2 = QHBoxLayout(self.horizontalLayoutWidget_2)
        self.horizontalLayout_2.setContentsMargins(0, 0, 0, 0)

        self.exitButton = QPushButton(self.horizontalLayoutWidget_2)
        font = QFont()
        font.setPointSize(20)
        self.exitButton.setFont(font)
        self.horizontalLayout_2.addWidget(self.exitButton)
        self.exitButton.clicked.connect(self.close)

        self.okButton = QPushButton(self.horizontalLayoutWidget_2)
        font = QFont()
        font.setPointSize(20)
        self.okButton.setFont(font)
        self.horizontalLayout_2.addWidget(self.okButton)

        self.retranslateUi(self)
        QMetaObject.connectSlotsByName(self)

    def toolButton_clicked(self):
        fname = QFileDialog.getExistingDirectory(self)
        self.lineEdit.setText(fname)

    # def openFileNamesDialog(self):
    #     options = QFileDialog.Options()
    #
    #     options |= QFileDialog.DontUseNativeDialog
    #     self.files, _ = QFileDialog.getOpenFileNames(self, "QFileDialog.getOpenFileNames()", "", "All Files (*);;Excel "
    #                                                                                     "Files(*.xlsx)", options=options)
    #     if self.files:
    #         self.calCount()
    #
    # def calCount(self):
    #     df = pd.DataFrame()
    #     for f in self.files:
    #         data = pd.read_excel(f, sheet_name="십일조", header=None)
    #         df = df.append(data)
    #     df = df.groupby(0).size().reset_index(name="1")
    #     df.to_excel('합산 결과.xlsx', sheet_name='십일조', index=False, header=None)

    def getFolderPath(self, text):
        path = os.chdir(text)
        files = os.listdir(path)
        self.files_xls = [f for f in files if f[-4:] == 'xlsx']
        self.okButton.clicked.connect(self.okButton_clicked)

    @QtCore.pyqtSlot()
    def okButton_clicked(self):
        dig = typeDialog(self.files_xls)
        dig.exec_()

    def retranslateUi(self, Dialog):
        _translate = QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "재정 프로그램"))
        self.label.setText(_translate("Dialog", "Folder:"))
        self.toolButton.setText(_translate("Dialog", "..."))
        self.exitButton.setText(_translate("Dialog", "Exit"))
        self.okButton.setText(_translate("Dialog", "OK"))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    app.exec_()
