# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'D:\ZRenyinquan\eric6\eric6-21.11\AutoPro_CPy_GUI\PIForm.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(400, 300)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/bat.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Form.setWindowIcon(icon)
        self.btn_PI = QtWidgets.QPushButton(Form)
        self.btn_PI.setGeometry(QtCore.QRect(100, 110, 181, 41))
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/button.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btn_PI.setIcon(icon1)
        self.btn_PI.setObjectName("btn_PI")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "正极电流测试项"))
        self.btn_PI.setText(_translate("Form", "正极电流测试项"))
import images_rc


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())
