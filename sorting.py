# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Sorting.ui'
#
# Created by: PyQt5 UI code generator 5.15.9
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets
import os


class Button(QtWidgets.QLineEdit):

	def __init__(self, parent):
		super(Button, self).__init__(parent)

		self.setAcceptDrops(True)

	def dragEnterEvent(self, e):

		if e.mimeData().hasUrls():
			e.accept()
		else:
			super(Button, self).dragEnterEvent(e)

	def dragMoveEvent(self, e):

		super(Button, self).dragMoveEvent(e)

	def dropEvent(self, e):

		if e.mimeData().hasUrls():
			for url in e.mimeData().urls():
				self.setText(os.path.normcase(url.toLocalFile()))
				e.accept()
		else:
			super(Button, self).dropEvent(e)


class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(607, 298)
        self.gridLayout = QtWidgets.QGridLayout(Dialog)
        self.gridLayout.setObjectName("gridLayout")
        self.label_main = QtWidgets.QLabel(Dialog)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_main.sizePolicy().hasHeightForWidth())
        self.label_main.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label_main.setFont(font)
        self.label_main.setAlignment(QtCore.Qt.AlignCenter)
        self.label_main.setObjectName("label_main")
        self.gridLayout.addWidget(self.label_main, 0, 2, 1, 3)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.pushButton_cancel = QtWidgets.QPushButton(Dialog)
        self.pushButton_cancel.setEnabled(True)
        font = QtGui.QFont()
        font.setPointSize(12)
        self.pushButton_cancel.setFont(font)
        self.pushButton_cancel.setObjectName("pushButton_cancel")
        self.horizontalLayout.addWidget(self.pushButton_cancel)
        self.pushButton_sorting = QtWidgets.QPushButton(Dialog)
        font = QtGui.QFont()
        font.setPointSize(12)
        self.pushButton_sorting.setFont(font)
        self.pushButton_sorting.setObjectName("pushButton_sorting")
        self.horizontalLayout.addWidget(self.pushButton_sorting)
        self.gridLayout.addLayout(self.horizontalLayout, 9, 2, 1, 3)
        self.label_gk = QtWidgets.QLabel(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_gk.setFont(font)
        self.label_gk.setObjectName("label_gk")
        self.gridLayout.addWidget(self.label_gk, 6, 2, 1, 1)
        self.lineEdit_path_file_sp = Button(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.lineEdit_path_file_sp.setFont(font)
        self.lineEdit_path_file_sp.setObjectName("lineEdit_path_file_sp")
        self.gridLayout.addWidget(self.lineEdit_path_file_sp, 2, 3, 1, 1)
        self.lineEdit_name_gk = QtWidgets.QLineEdit(Dialog)
        self.lineEdit_name_gk.setEnabled(True)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.lineEdit_name_gk.setFont(font)
        self.lineEdit_name_gk.setObjectName("lineEdit_name_gk")
        self.gridLayout.addWidget(self.lineEdit_name_gk, 6, 3, 1, 2)
        self.pushButton_folder_document = QtWidgets.QPushButton(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.pushButton_folder_document.setFont(font)
        self.pushButton_folder_document.setObjectName("pushButton_folder_document")
        self.gridLayout.addWidget(self.pushButton_folder_document, 1, 4, 1, 1)
        self.lineEdit_path_folder_document = Button(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.lineEdit_path_folder_document.setFont(font)
        self.lineEdit_path_folder_document.setObjectName("lineEdit_path_folder_document")
        self.gridLayout.addWidget(self.lineEdit_path_folder_document, 1, 3, 1, 1)
        self.label_document_sp = QtWidgets.QLabel(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.label_document_sp.setFont(font)
        self.label_document_sp.setObjectName("label_document_sp")
        self.gridLayout.addWidget(self.label_document_sp, 8, 2, 1, 1)
        self.label_folder_document = QtWidgets.QLabel(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.label_folder_document.setFont(font)
        self.label_folder_document.setObjectName("label_folder_document")
        self.gridLayout.addWidget(self.label_folder_document, 1, 2, 1, 1)
        self.label_file_sp = QtWidgets.QLabel(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.label_file_sp.setFont(font)
        self.label_file_sp.setObjectName("label_file_sp")
        self.gridLayout.addWidget(self.label_file_sp, 2, 2, 1, 1)
        self.pushButton_file_sp = QtWidgets.QPushButton(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.pushButton_file_sp.setFont(font)
        self.pushButton_file_sp.setObjectName("pushButton_file_sp")
        self.gridLayout.addWidget(self.pushButton_file_sp, 2, 4, 1, 1)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.checkBox_conclusion_sp = QtWidgets.QCheckBox(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.checkBox_conclusion_sp.setFont(font)
        self.checkBox_conclusion_sp.setObjectName("checkBox_conclusion_sp")
        self.horizontalLayout_3.addWidget(self.checkBox_conclusion_sp)
        self.checkBox_protocol_sp = QtWidgets.QCheckBox(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.checkBox_protocol_sp.setFont(font)
        self.checkBox_protocol_sp.setObjectName("checkBox_protocol_sp")
        self.horizontalLayout_3.addWidget(self.checkBox_protocol_sp)
        self.checkBox_preciption_sp = QtWidgets.QCheckBox(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.checkBox_preciption_sp.setFont(font)
        self.checkBox_preciption_sp.setObjectName("checkBox_preciption_sp")
        self.horizontalLayout_3.addWidget(self.checkBox_preciption_sp)
        self.checkBox_infocard_sp = QtWidgets.QCheckBox(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.checkBox_infocard_sp.setFont(font)
        self.checkBox_infocard_sp.setObjectName("checkBox_infocard_sp")
        self.horizontalLayout_3.addWidget(self.checkBox_infocard_sp)
        self.gridLayout.addLayout(self.horizontalLayout_3, 8, 3, 1, 2)
        self.pushButton_finish_folder = QtWidgets.QPushButton(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.pushButton_finish_folder.setFont(font)
        self.pushButton_finish_folder.setObjectName("pushButton_finish_folder")
        self.gridLayout.addWidget(self.pushButton_finish_folder, 3, 4, 1, 1)
        self.lineEdit_path_folder_finish = Button(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.lineEdit_path_folder_finish.setFont(font)
        self.lineEdit_path_folder_finish.setObjectName("lineEdit_path_folder_finish")
        self.gridLayout.addWidget(self.lineEdit_path_folder_finish, 3, 3, 1, 1)
        self.label_finish_folder = QtWidgets.QLabel(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_finish_folder.setFont(font)
        self.label_finish_folder.setObjectName("label_finish_folder")
        self.gridLayout.addWidget(self.label_finish_folder, 3, 2, 1, 1)

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Сортировка"))
        self.label_main.setText(_translate("Dialog", "Сортировка документов"))
        self.pushButton_cancel.setText(_translate("Dialog", "Отмена"))
        self.pushButton_sorting.setText(_translate("Dialog", "Сортировка"))
        self.label_gk.setText(_translate("Dialog", "Имя ГК"))
        self.pushButton_folder_document.setText(_translate("Dialog", "Открыть"))
        self.label_document_sp.setText(_translate("Dialog", "Проверка документов"))
        self.label_folder_document.setText(_translate("Dialog", "Каталог с документами"))
        self.label_file_sp.setText(_translate("Dialog", "Файл с номерами"))
        self.pushButton_file_sp.setText(_translate("Dialog", "Открыть"))
        self.checkBox_conclusion_sp.setText(_translate("Dialog", "Заключение"))
        self.checkBox_protocol_sp.setText(_translate("Dialog", "Протокол"))
        self.checkBox_preciption_sp.setText(_translate("Dialog", "Предписание"))
        self.checkBox_infocard_sp.setText(_translate("Dialog", "Инфокарта"))
        self.pushButton_finish_folder.setText(_translate("Dialog", "Открыть"))
        self.label_finish_folder.setText(_translate("Dialog", "Конечная папка"))