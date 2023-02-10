# -*- coding: utf-8 -*-



from __future__ import print_function 
from mailmerge import MailMerge
import os
import sys
import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
import win32com.client
import pythoncom

class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(1142, 644)
        Dialog.setMouseTracking(False)
        Dialog.setFocusPolicy(QtCore.Qt.WheelFocus)
        Dialog.setToolTip("")
        Dialog.setLayoutDirection(QtCore.Qt.LeftToRight)
        Dialog.setStyleSheet("")
        self.widget = QtWidgets.QWidget(Dialog)
        self.widget.setGeometry(QtCore.QRect(140, 80, 821, 491))
        self.widget.setMouseTracking(True)
        self.widget.setFocusPolicy(QtCore.Qt.TabFocus)
        self.widget.setAutoFillBackground(False)
        self.widget.setStyleSheet("QWidget #widget {border-image: url(:/Image/image/IBSS.jpg) rgb(0,23,0);}\n"
"\n"
"\n"
"QWidget #widget {border-top-right-radius:30px;}\n"
"QWidget #widget {border-bottom-left-radius:30px;}")
        self.widget.setObjectName("widget")
        self.verticalLayoutWidget = QtWidgets.QWidget(self.widget)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(630, 110, 121, 121))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.add_word = QtWidgets.QPushButton(self.verticalLayoutWidget)
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.PlaceholderText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.PlaceholderText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.PlaceholderText, brush)
        self.add_word.setPalette(palette)
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(24)
        self.add_word.setFont(font)
        self.add_word.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.add_word.setStyleSheet("#add_word{\n"
"    \n"
"    color: rgb(255, 255, 255);\n"
"    border:3px solid rgb(255,255,255);\n"
"    border-radius:8px;\n"
"}\n"
"#add_word:hover{\n"
"    \n"
"    border-color: rgb(44, 9, 103);\n"
"}\n"
"#add_word:pressed{\n"
"    padding-top:5px;\n"
"    padding-left:5px;\n"
"    \n"
"}")
        self.add_word.setObjectName("add_word")
        self.verticalLayout.addWidget(self.add_word)
        self.add_excel = QtWidgets.QPushButton(self.verticalLayoutWidget)
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.PlaceholderText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.PlaceholderText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.PlaceholderText, brush)
        self.add_excel.setPalette(palette)
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(24)
        self.add_excel.setFont(font)
        self.add_excel.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.add_excel.setStyleSheet("#add_excel{    \n"
"    color: rgb(255, 255, 255);\n"
"    border:3px solid rgb(255,255,255);\n"
"    border-radius:8px;\n"
"}\n"
"#add_excel:hover{\n"
"     border-color: rgb(44, 9, 103);\n"
"}\n"
"#add_excel:pressed{\n"
"    padding-top:5px;\n"
"    padding-left:5px;\n"
"}")
        self.add_excel.setObjectName("add_excel")
        self.verticalLayout.addWidget(self.add_excel)
        self.progressBar = QtWidgets.QProgressBar(self.widget)
        self.progressBar.setGeometry(QtCore.QRect(30, 460, 118, 23))
        self.progressBar.setStyleSheet("QProgressBar {\n"
"        background-color: rgb(98,114,164);\n"
"        color: rgb(200,200,200);\n"
"        border-style:none;\n"
"        border-radius: 10px;\n"
"}\n"
"QProgressBar::chunk {\n"
"    background-color: qlineargradient(spread:pad, x1:0, y1:0.511364, x2:1, y2:0.523, stop:0 rgba(0, 0, 0, 255), stop:1 rgba(100, 0, 223, 255));\n"
"        border-radius: 10px;\n"
"}")
        self.progressBar.setProperty("value", 0)
        self.progressBar.setTextVisible(False)
        self.progressBar.setVisible(False)
        self.progressBar.setObjectName("progressBar")
        self.result = QtWidgets.QLabel(self.widget)
        self.result.setGeometry(QtCore.QRect(120, 260, 621, 81))
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(17)
        font.setBold(True)
        font.setWeight(75)
        self.result.setFont(font)
        self.result.setAutoFillBackground(False)
        self.result.setStyleSheet("color: rgb(255, 255, 255);\n"
"\n"
"border-radius:8px;\n"
"\n"
"background-color:  rgba(0, 0, 0, 150)\n"
"")
        self.result.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.result.setFrameShadow(QtWidgets.QFrame.Plain)
        self.result.setTextFormat(QtCore.Qt.AutoText)
        self.result.setScaledContents(True)
        self.result.setAlignment(QtCore.Qt.AlignCenter)
        self.result.setWordWrap(True)
        self.result.setObjectName("result")
        self.label_2 = QtWidgets.QLabel(self.widget)
        self.label_2.setGeometry(QtCore.QRect(770, 460, 54, 31))
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(16)
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setStyleSheet("color: rgb(255,255,255)\n"
"")
        self.label_2.setObjectName("label_2")
        self.heading = QtWidgets.QLabel(self.widget)
        self.heading.setGeometry(QtCore.QRect(0, 0, 821, 41))
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(17)
        font.setBold(True)
        font.setWeight(75)
        self.heading.setFont(font)
        self.heading.setAutoFillBackground(False)
        self.heading.setStyleSheet("color: rgb(255, 255, 255);\n"
"border-top-right-radius:30px;\n"
"background-color:  rgba(0, 0, 0, 110)\n"
"")
        self.heading.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.heading.setFrameShadow(QtWidgets.QFrame.Plain)
        self.heading.setTextFormat(QtCore.Qt.AutoText)
        self.heading.setScaledContents(True)
        self.heading.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.heading.setWordWrap(True)
        self.heading.setIndent(10)
        self.heading.setObjectName("heading")
        self.pushButton_1 = QtWidgets.QPushButton(self.widget)
        self.pushButton_1.setGeometry(QtCore.QRect(770, 10, 41, 31))
        self.pushButton_1.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.pushButton_1.setStyleSheet("background-color:  rgba(0, 0, 0, 110)\n"
"")
        self.pushButton_1.setText("")
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/Image/image/Picture1.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.pushButton_1.setIcon(icon)
        self.pushButton_1.setIconSize(QtCore.QSize(20, 20))
        self.pushButton_1.setObjectName("pushButton_1")
        self.pushButton_2 = QtWidgets.QPushButton(self.widget)
        self.pushButton_2.setGeometry(QtCore.QRect(720, 10, 41, 31))
        self.pushButton_2.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.pushButton_2.setStyleSheet("background-color:  rgba(0, 0, 0, 110)")
        self.pushButton_2.setText("")
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/Image/image/Picture2.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.pushButton_2.setIcon(icon1)
        self.pushButton_2.setIconSize(QtCore.QSize(20, 20))
        self.pushButton_2.setObjectName("pushButton_2")
        self.save = QtWidgets.QPushButton(self.widget)
        self.save.setGeometry(QtCore.QRect(530, 360, 101, 51))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.PlaceholderText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.PlaceholderText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.PlaceholderText, brush)
        self.save.setPalette(palette)
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(24)
        self.save.setFont(font)
        self.save.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.save.setStyleSheet("#save{\n"
"    background-color: rgb(0, 0, 0);\n"
"    color: rgb(255, 255, 255);\n"
"    border:3px solid rgb(255,255,255);\n"
"    border-radius:8px;\n"
"}\n"
"#save:hover{\n"
"    background-color: rgb(255, 255, 255);\n"
"    border-color: rgb(44, 9, 103);\n"
"    color: rgb(0, 0, 0);\n"
"}\n"
"#save:pressed{\n"
"    padding-top:5px;\n"
"    padding-left:5px;\n"
"}")
        self.save.setObjectName("save")
        self.quit = QtWidgets.QPushButton(self.widget)
        self.quit.setGeometry(QtCore.QRect(640, 360, 101, 51))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.PlaceholderText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.PlaceholderText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.PlaceholderText, brush)
        self.quit.setPalette(palette)
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(24)
        self.quit.setFont(font)
        self.quit.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.quit.setStyleSheet("#quit{\n"
"    background-color: rgb(0, 0, 0);\n"
"    color: rgb(255, 255, 255);\n"
"    border:3px solid rgb(255,255,255);\n"
"    border-radius:8px;\n"
"}\n"
"#quit:hover{\n"
"    background-color: rgb(255, 255, 255);\n"
"    border-color: rgb(44, 9, 103);\n"
"    color: rgb(0, 0, 0);\n"
"}\n"
"#quit:pressed{\n"
"    padding-top:5px;\n"
"    padding-left:5px;\n"
"}")
        self.quit.setObjectName("quit")
        self.lineEdit_3 = QtWidgets.QLineEdit(self.widget)
        self.lineEdit_3.setGeometry(QtCore.QRect(130, 370, 371, 41))
        self.lineEdit_3.setMaximumSize(QtCore.QSize(481, 41))
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(16)
        font.setItalic(True)
        self.lineEdit_3.setFont(font)
        self.lineEdit_3.setCursor(QtGui.QCursor(QtCore.Qt.IBeamCursor))
        self.lineEdit_3.setFocusPolicy(QtCore.Qt.ClickFocus)
        self.lineEdit_3.setAcceptDrops(True)
        self.lineEdit_3.setStyleSheet("border-color: rgb(0, 0, 0,150);\n"
"border-style:outset;\n"
"border-width:1px;")
        self.lineEdit_3.setFrame(False)
        self.lineEdit_3.setDragEnabled(False)
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.BackgroundDim = QtWidgets.QLabel(self.widget)
        self.BackgroundDim.setGeometry(QtCore.QRect(0, 30, 821, 461))
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(17)
        font.setBold(True)
        font.setWeight(75)
        self.BackgroundDim.setFont(font)
        self.BackgroundDim.setAutoFillBackground(False)
        self.BackgroundDim.setStyleSheet("color: rgb(255, 255, 255);\n"
"border-top-right-radius:30px;\n"
"background-color:  rgba(0, 0, 0, 20)\n"
"")
        self.BackgroundDim.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.BackgroundDim.setFrameShadow(QtWidgets.QFrame.Plain)
        self.BackgroundDim.setText("")
        self.BackgroundDim.setTextFormat(QtCore.Qt.AutoText)
        self.BackgroundDim.setScaledContents(True)
        self.BackgroundDim.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.BackgroundDim.setWordWrap(True)
        self.BackgroundDim.setIndent(10)
        self.BackgroundDim.setObjectName("BackgroundDim")
        self.comboBox = QtWidgets.QComboBox(self.widget)
        self.comboBox.setGeometry(QtCore.QRect(450, 180, 151, 41))
        self.comboBox.setStyleSheet("\n"
"border-style:outset;\n"
"\n"
"border-width:1px;\n"
"")
        self.comboBox.setObjectName("comboBox")
        self.comboBox.setStyleSheet("font-size: 18px; font-weight: bold;font-family: Trebuchet MS")
        self.lineEdit_1 = QtWidgets.QLineEdit(self.widget)
        self.lineEdit_1.setGeometry(QtCore.QRect(120, 120, 481, 41))
        self.lineEdit_1.setMaximumSize(QtCore.QSize(481, 41))
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(16)
        font.setItalic(True)
        self.lineEdit_1.setFont(font)
        self.lineEdit_1.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))
        self.lineEdit_1.setFocusPolicy(QtCore.Qt.NoFocus)
        self.lineEdit_1.setContextMenuPolicy(QtCore.Qt.DefaultContextMenu)
        self.lineEdit_1.setAcceptDrops(False)
        self.lineEdit_1.setStyleSheet("color:rgb(0,0,0);\n"
"\n"
"border-style:outset;\n"
"border-color: rgb(0, 0, 0,150);\n"
"border-width:1px;\n"
"border-radius:8px;\n"
"")
        self.lineEdit_1.setInputMask("")
        self.lineEdit_1.setFrame(False)
        self.lineEdit_1.setDragEnabled(True)
        self.lineEdit_1.setCursorMoveStyle(QtCore.Qt.LogicalMoveStyle)
        self.lineEdit_1.setClearButtonEnabled(False)
        self.lineEdit_1.setObjectName("lineEdit_1")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.widget)
        self.lineEdit_2.setGeometry(QtCore.QRect(120, 180, 331, 41))
        self.lineEdit_2.setMaximumSize(QtCore.QSize(481, 41))
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(16)
        font.setItalic(True)
        self.lineEdit_2.setFont(font)
        self.lineEdit_2.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))
        self.lineEdit_2.setFocusPolicy(QtCore.Qt.NoFocus)
        self.lineEdit_2.setAcceptDrops(True)
        self.lineEdit_2.setStyleSheet("border-top-left-radius:8px;\n"
"border-bottom-left-radius:8px;\n"
"border-style:outset;\n"
"border-color: rgb(0, 0, 0,150);\n"
"border-width:1px;")
        self.lineEdit_2.setFrame(False)
        self.lineEdit_2.setDragEnabled(False)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.BackgroundDim.raise_()
        self.heading.raise_()
        self.quit.raise_()
        self.result.raise_()
        self.verticalLayoutWidget.raise_()
        self.progressBar.raise_()
        self.label_2.raise_()
        self.pushButton_1.raise_()
        self.pushButton_2.raise_()
        self.lineEdit_3.raise_()
        self.save.raise_()
        self.comboBox.raise_()

        self.retranslateUi(Dialog)
        self.pushButton_1.clicked.connect(Dialog.close)
        self.pushButton_2.clicked.connect(Dialog.showMinimized)
        self.save.clicked.connect(self.lineEdit_2.clear)
        self.save.clicked.connect(self.progressBar.show)
        self.quit.clicked.connect(Dialog.close)
        self.save.clicked.connect(self.lineEdit_1.clear)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.add_word.setText(_translate("Dialog", "Browse"))
        self.add_excel.setText(_translate("Dialog", "Browse"))
        self.result.setText(_translate("Dialog", "Welcome!\
                                       \
                                    Enter e-mail\'s title below"))
        self.label_2.setText(_translate("Dialog", "v1.0"))
        self.heading.setText(_translate("Dialog", "Customized Mail Sender by DALab"))
        self.save.setText(_translate("Dialog", "Run"))
        self.quit.setText(_translate("Dialog", "Quit"))
        self.lineEdit_3.setPlaceholderText(_translate("Dialog", " Title"))
        self.lineEdit_1.setPlaceholderText(_translate("Dialog", " Word Template"))
        self.lineEdit_2.setPlaceholderText(_translate("Dialog", " Excel"))
import resource



from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QWidget
from CES import Ui_Dialog


        
class Mywindow(QMainWindow, Ui_Dialog, QWidget):
    def __init__(self, parent=None):
        super(Mywindow, self).__init__(parent)
        self.setupUi(self)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.setWindowTitle("Customized Mail Sender v1.0")
        self.add_word.clicked.connect(self.read_word)
        self.add_excel.clicked.connect(self.read_excel)  
        self.save.clicked.connect(self.process)



    
    def read_word(self):
        global word
        word = QFileDialog.getOpenFileName(self,'选择文件','','word files(*.doc , *.docx)')
        self.lineEdit_1.setText(word[0])
        
    def read_excel(self):
        global excel
        excel = QFileDialog.getOpenFileName(self,'选择文件','','Excel files(*.xlsx , *.xls)')
        self.lineEdit_2.setText(excel[0])
        Info = pd.read_excel(excel[0])
        excel_field = Info.columns
        self.comboBox.clear()
        self.comboBox.addItems(excel_field)


    def process(self):
        QApplication.processEvents()
        try:
            template = word[0]
            Info = pd.read_excel(excel[0])
            
        except (NameError, FileNotFoundError, AssertionError):
               self.result.setText("Please check your inputs")
               self.progressBar.setProperty("value",100)

        else:
            pythoncom.CoInitialize()
            document = MailMerge(template)
    
            word_field = document.get_merge_fields()
            excel_field = Info.columns
            tmp=list(word_field.difference(excel_field))
        
    
            self.progressBar.setRange(0,len(Info)-1)
    
            if len(tmp)==0:
                for i in range(len(Info)):
                     QApplication.processEvents();
                     document = MailMerge(template)
                     d={}
                     for j in range(len(word_field)):
                         x = list(word_field)[j]
                         y = Info[x][i]
                         y = str(y)
                         d[x] = y
                     document.merge(**d)
                     
                     Path = os.path.join(os.path.expanduser('~'),"Desktop")
                     document.write(Path +'/'+ format(i) + ".docx")
                     document.close()

                     doc = win32com.client.Dispatch("Word.Application")
                     body = doc.Documents.Open(Path +'/'+ format(i) + ".docx")
                     body.Content.Copy()
                     body.Close()
                     
                     title = self.lineEdit_3.text()   
                     if title == "":
                         title = "Default Subject"
                     
                     Address = self.comboBox.currentText()
                     receivers = Info[Address][i]
                     
                     
                     outlook = win32com.client.Dispatch("Outlook.Application")
                     mail = outlook.CreateItem(0)  
                     mail.GetInspector.WordEditor.Range(Start=0, End=0).Paste()
                     mail.To = receivers
                     mail.Subject = title
                     mail.display()
                     mail.Send()
                     
                     os.remove(Path +'/'+ format(i) + ".docx")
                     
                     self.progressBar.setProperty("value",i)
                     self.result.setText("Success!")
                
                doc.Quit()
                

            else:
                self.progressBar.setProperty("value",len(Info)-1)
                self.result.setText('No field(s) called '+'"'+'","'.join(tmp)+'"'+' !')
            
            pythoncom.CoUninitialize()    
            
            del globals()['word']
            del globals()['excel']
            
    def mouseMoveEvent(self, e: QtGui.QMouseEvent): 
        if e.y()<120:
            self._endPos = e.pos() - self._startPos
            self.move(self.pos() + self._endPos)


    def mousePressEvent(self, e: QtGui.QMouseEvent):
            if e.button() == QtCore.Qt.LeftButton:
                self._isTracking = True
                self._startPos = QtCore.QPoint(e.x(), e.y())

    def mouseReleaseEvent(self, e: QtGui.QMouseEvent):
            if e.button() == QtCore.Qt.LeftButton:
                self._isTracking = False
                self._startPos = None
                self._endPos = None
   
if __name__ == '__main__':
    app = QApplication(sys.argv) 
    ui = Mywindow()
    ui.show()
    sys.exit(app.exec_())    
    
