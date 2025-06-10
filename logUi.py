# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'logUi.명령어'
##
## Created by: Qt User Interface Compiler version 6.6.3
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication,     QMetaObject,   QRect,     Qt)

from PySide6.QtWidgets import (QAbstractScrollArea,   QFrame,    QLabel, QPlainTextEdit, QPushButton    )

class Ui_Dialog(object):
    def setupUi(self, Dialog):
        if not Dialog.objectName():
            Dialog.setObjectName(u"Dialog")
        Dialog.resize(615, 480)
        self.plainTextEdit = QPlainTextEdit(Dialog)
        self.plainTextEdit.setObjectName(u"plainTextEdit")
        self.plainTextEdit.setEnabled(True)
        self.plainTextEdit.setGeometry(QRect(10, 20, 591, 411))
        self.plainTextEdit.setContextMenuPolicy(Qt.DefaultContextMenu)
        self.plainTextEdit.setFrameShape(QFrame.Box)
        self.plainTextEdit.setLineWidth(1)
        self.plainTextEdit.setMidLineWidth(1)
        self.plainTextEdit.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.plainTextEdit.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.plainTextEdit.setSizeAdjustPolicy(QAbstractScrollArea.AdjustToContentsOnFirstShow)
        self.plainTextEdit.setReadOnly(True)
        self.plainTextEdit.setCenterOnScroll(False)
        self.pushButton = QPushButton(Dialog)
        self.pushButton.setObjectName(u"pushButton")
        self.pushButton.setGeometry(QRect(540, 440, 61, 31))
        self.label = QLabel(Dialog)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(10, 0, 81, 21))

        self.retranslateUi(Dialog)

        QMetaObject.connectSlotsByName(Dialog)
    # setupUi

    def retranslateUi(self, Dialog):
        Dialog.setWindowTitle(QCoreApplication.translate("Dialog", u"\uc2e4\ud589\uacb0\uacfc \uc870\ud68c", None))
        self.pushButton.setText(QCoreApplication.translate("Dialog", u"\ub2eb\uae30", None))
        self.label.setText(QCoreApplication.translate("Dialog", u"\uc2e4\ud589\uacb0\uacfc \uc870\ud68c", None))
    # retranslateUi

