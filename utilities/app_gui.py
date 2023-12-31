# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'app_gui.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(745, 527)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label_45 = QtWidgets.QLabel(self.centralwidget)
        self.label_45.setGeometry(QtCore.QRect(20, 280, 181, 20))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setBold(False)
        font.setWeight(50)
        self.label_45.setFont(font)
        self.label_45.setObjectName("label_45")
        self.label_51 = QtWidgets.QLabel(self.centralwidget)
        self.label_51.setGeometry(QtCore.QRect(12, 200, 381, 21))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(75)
        self.label_51.setFont(font)
        self.label_51.setObjectName("label_51")
        self.label_48 = QtWidgets.QLabel(self.centralwidget)
        self.label_48.setGeometry(QtCore.QRect(20, 240, 231, 20))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setBold(False)
        font.setWeight(50)
        self.label_48.setFont(font)
        self.label_48.setObjectName("label_48")
        self.browse_ha_files_folder = QtWidgets.QPushButton(self.centralwidget)
        self.browse_ha_files_folder.setGeometry(QtCore.QRect(580, 240, 151, 30))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setBold(True)
        font.setWeight(75)
        self.browse_ha_files_folder.setFont(font)
        self.browse_ha_files_folder.setObjectName("browse_ha_files_folder")
        self.browse_ppt_template_file = QtWidgets.QPushButton(self.centralwidget)
        self.browse_ppt_template_file.setGeometry(QtCore.QRect(580, 280, 151, 30))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setBold(True)
        font.setWeight(75)
        self.browse_ppt_template_file.setFont(font)
        self.browse_ppt_template_file.setObjectName("browse_ppt_template_file")
        self.ha_files_folder = QtWidgets.QLabel(self.centralwidget)
        self.ha_files_folder.setEnabled(False)
        self.ha_files_folder.setGeometry(QtCore.QRect(250, 240, 311, 25))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setBold(False)
        font.setWeight(50)
        self.ha_files_folder.setFont(font)
        self.ha_files_folder.setFrameShape(QtWidgets.QFrame.Box)
        self.ha_files_folder.setMidLineWidth(1)
        self.ha_files_folder.setObjectName("ha_files_folder")
        self.ppt_template_file = QtWidgets.QLabel(self.centralwidget)
        self.ppt_template_file.setEnabled(False)
        self.ppt_template_file.setGeometry(QtCore.QRect(250, 280, 311, 25))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setBold(False)
        font.setWeight(50)
        self.ppt_template_file.setFont(font)
        self.ppt_template_file.setFrameShape(QtWidgets.QFrame.Box)
        self.ppt_template_file.setMidLineWidth(1)
        self.ppt_template_file.setObjectName("ppt_template_file")
        self.label_54 = QtWidgets.QLabel(self.centralwidget)
        self.label_54.setGeometry(QtCore.QRect(10, 90, 381, 21))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(75)
        self.label_54.setFont(font)
        self.label_54.setObjectName("label_54")
        self.label_55 = QtWidgets.QLabel(self.centralwidget)
        self.label_55.setGeometry(QtCore.QRect(18, 130, 121, 20))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setBold(False)
        font.setWeight(50)
        self.label_55.setFont(font)
        self.label_55.setObjectName("label_55")
        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setGeometry(QtCore.QRect(10, 320, 720, 21))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.exit_app = QtWidgets.QPushButton(self.centralwidget)
        self.exit_app.setGeometry(QtCore.QRect(530, 440, 151, 30))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setBold(True)
        font.setWeight(75)
        self.exit_app.setFont(font)
        self.exit_app.setObjectName("exit_app")
        self.reset_selections = QtWidgets.QPushButton(self.centralwidget)
        self.reset_selections.setGeometry(QtCore.QRect(100, 440, 151, 30))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setBold(True)
        font.setWeight(75)
        self.reset_selections.setFont(font)
        self.reset_selections.setObjectName("reset_selections")
        self.generate_reports = QtWidgets.QPushButton(self.centralwidget)
        self.generate_reports.setGeometry(QtCore.QRect(280, 440, 221, 30))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setBold(True)
        font.setWeight(75)
        self.generate_reports.setFont(font)
        self.generate_reports.setObjectName("generate_reports")
        self.line_4 = QtWidgets.QFrame(self.centralwidget)
        self.line_4.setGeometry(QtCore.QRect(10, 170, 720, 21))
        self.line_4.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_4.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_4.setObjectName("line_4")
        self.cr_number = QtWidgets.QLineEdit(self.centralwidget)
        self.cr_number.setGeometry(QtCore.QRect(138, 120, 171, 30))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.cr_number.setFont(font)
        self.cr_number.setObjectName("cr_number")
        self.label_58 = QtWidgets.QLabel(self.centralwidget)
        self.label_58.setGeometry(QtCore.QRect(348, 130, 121, 20))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setBold(False)
        font.setWeight(50)
        self.label_58.setFont(font)
        self.label_58.setObjectName("label_58")
        self.sr_number = QtWidgets.QLineEdit(self.centralwidget)
        self.sr_number.setGeometry(QtCore.QRect(468, 120, 171, 30))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.sr_number.setFont(font)
        self.sr_number.setObjectName("sr_number")
        self.browse_output_ppt_folder = QtWidgets.QPushButton(self.centralwidget)
        self.browse_output_ppt_folder.setGeometry(QtCore.QRect(580, 390, 151, 30))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setBold(True)
        font.setWeight(75)
        self.browse_output_ppt_folder.setFont(font)
        self.browse_output_ppt_folder.setObjectName("browse_output_ppt_folder")
        self.output_ppt_folder = QtWidgets.QLabel(self.centralwidget)
        self.output_ppt_folder.setEnabled(False)
        self.output_ppt_folder.setGeometry(QtCore.QRect(250, 390, 311, 25))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setBold(False)
        font.setWeight(50)
        self.output_ppt_folder.setFont(font)
        self.output_ppt_folder.setFrameShape(QtWidgets.QFrame.Box)
        self.output_ppt_folder.setMidLineWidth(1)
        self.output_ppt_folder.setObjectName("output_ppt_folder")
        self.label_50 = QtWidgets.QLabel(self.centralwidget)
        self.label_50.setGeometry(QtCore.QRect(20, 390, 231, 20))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setBold(False)
        font.setWeight(50)
        self.label_50.setFont(font)
        self.label_50.setObjectName("label_50")
        self.label_52 = QtWidgets.QLabel(self.centralwidget)
        self.label_52.setGeometry(QtCore.QRect(10, 350, 381, 21))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(75)
        self.label_52.setFont(font)
        self.label_52.setObjectName("label_52")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(10, 10, 141, 61))
        self.label.setStyleSheet("image: url(:/newPrefix/quest_logo.jpg);")
        self.label.setText("")
        self.label.setScaledContents(True)
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(570, 10, 161, 61))
        self.label_2.setStyleSheet("image: url(:/newPrefix2/ihi_logo.jpg);")
        self.label_2.setText("")
        self.label_2.setScaledContents(True)
        self.label_2.setObjectName("label_2")
        self.label_56 = QtWidgets.QLabel(self.centralwidget)
        self.label_56.setGeometry(QtCore.QRect(270, 20, 211, 31))
        font = QtGui.QFont()
        font.setFamily("Verdana")
        font.setPointSize(15)
        font.setBold(True)
        font.setUnderline(True)
        font.setWeight(75)
        self.label_56.setFont(font)
        self.label_56.setObjectName("label_56")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 745, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "GE Passport"))
        self.label_45.setText(_translate("MainWindow", "<html><head/><body><p>PPT TEMPLATE <span style=\" color:#ff0000;\">#</span></p></body></html>"))
        self.label_51.setText(_translate("MainWindow", "INPUT FILES LOCATION"))
        self.label_48.setText(_translate("MainWindow", "<html><head/><body><p>HA FILES FOLDER<span style=\" color:#ff0000;\">#</span></p></body></html>"))
        self.browse_ha_files_folder.setText(_translate("MainWindow", "BROWSE"))
        self.browse_ppt_template_file.setText(_translate("MainWindow", "BROWSE"))
        self.ha_files_folder.setText(_translate("MainWindow", "path...."))
        self.ppt_template_file.setText(_translate("MainWindow", "path...."))
        self.label_54.setText(_translate("MainWindow", "INPUT PARAMS"))
        self.label_55.setText(_translate("MainWindow", "<html><head/><body><p>CR NUMBER <span style=\" color:#ff0000;\">*</span></p></body></html>"))
        self.exit_app.setText(_translate("MainWindow", "EXIT"))
        self.reset_selections.setText(_translate("MainWindow", "RESET"))
        self.generate_reports.setText(_translate("MainWindow", "GENERATE REPORT"))
        self.label_58.setText(_translate("MainWindow", "<html><head/><body><p>SR NUMBER <span style=\" color:#ff0000;\">*</span></p></body></html>"))
        self.browse_output_ppt_folder.setText(_translate("MainWindow", "BROWSE"))
        self.output_ppt_folder.setText(_translate("MainWindow", "path...."))
        self.label_50.setText(_translate("MainWindow", "<html><head/><body><p>SAVE OUTPUT IN FOLDER <span style=\" color:#ff0000;\">#</span></p></body></html>"))
        self.label_52.setText(_translate("MainWindow", "OUTPUT FILES LOCATION"))
        self.label_56.setText(_translate("MainWindow", "GE Passport Tool"))

import utilities.source


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
