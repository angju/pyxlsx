# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'E:\python\pyxlsx\pyxlsx.ui'
#
# Created by: PyQt5 UI code generator 5.7.1
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(897, 508)
        Dialog.setMaximumSize(QtCore.QSize(897, 508))
        Dialog.setSizeGripEnabled(True)
        self.pushButtonRun = QtWidgets.QPushButton(Dialog)
        self.pushButtonRun.setGeometry(QtCore.QRect(100, 440, 141, 41))
        self.pushButtonRun.setObjectName("pushButtonRun")
        self.pushButtonClose = QtWidgets.QPushButton(Dialog)
        self.pushButtonClose.setGeometry(QtCore.QRect(270, 440, 131, 41))
        self.pushButtonClose.setObjectName("pushButtonClose")
        self.textBrowserOutput = QtWidgets.QTextBrowser(Dialog)
        self.textBrowserOutput.setGeometry(QtCore.QRect(500, 40, 321, 441))
        self.textBrowserOutput.setObjectName("textBrowserOutput")
        self.label = QtWidgets.QLabel(Dialog)
        self.label.setGeometry(QtCore.QRect(630, 20, 54, 12))
        self.label.setObjectName("label")
        self.layoutWidget = QtWidgets.QWidget(Dialog)
        self.layoutWidget.setGeometry(QtCore.QRect(30, 40, 451, 381))
        self.layoutWidget.setObjectName("layoutWidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.layoutWidget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.pushButtonSelectFile = QtWidgets.QPushButton(self.layoutWidget)
        self.pushButtonSelectFile.setObjectName("pushButtonSelectFile")
        self.verticalLayout.addWidget(self.pushButtonSelectFile)
        self.pushButtonClear = QtWidgets.QPushButton(self.layoutWidget)
        self.pushButtonClear.setObjectName("pushButtonClear")
        self.verticalLayout.addWidget(self.pushButtonClear)
        spacerItem = QtWidgets.QSpacerItem(0, 0, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem)
        self.horizontalLayout.addLayout(self.verticalLayout)
        self.textBrowser = QtWidgets.QTextBrowser(self.layoutWidget)
        self.textBrowser.setObjectName("textBrowser")
        self.horizontalLayout.addWidget(self.textBrowser)

        self.retranslateUi(Dialog)
        self.pushButtonClose.clicked.connect(Dialog.close)
        QtCore.QMetaObject.connectSlotsByName(Dialog)
        Dialog.setTabOrder(self.pushButtonRun, self.pushButtonClose)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.pushButtonRun.setText(_translate("Dialog", "运行"))
        self.pushButtonClose.setText(_translate("Dialog", "退出"))
        self.label.setText(_translate("Dialog", "log输出"))
        self.pushButtonSelectFile.setText(_translate("Dialog", "添加文件"))
        self.pushButtonClear.setText(_translate("Dialog", "清空文件"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())

