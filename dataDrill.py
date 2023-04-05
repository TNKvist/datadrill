# -*- coding: utf-8 -*-
"""
by: Taylor Kvist Mar-2023
"""

import pandas as pd
import sys,os,csv
import time
from PyQt5 import QtWidgets
from PyQt5 import QtCore
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QFont,QIcon,QPixmap
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, 
                             QFileDialog, QWidget, QLabel, QMessageBox, 
                             QComboBox, QTableWidget, QGridLayout, 
                             QTableWidgetItem,QSplashScreen)


#Main Window to select Files, Pick Location and which sheet to parse
class App(QMainWindow):

    def __init__(self):
        super().__init__()
        self.sheets = None
        self.location = None
        self.chosen_sheet = None
        self.df = None
        self.w = None

        self.left = 300
        self.top = 400
        self.width = 500
        self.height = 200
        self.setMinimumSize(self.width, self.height)
        self.setMaximumHeight(self.height)

        self.title = 'Data Drill v1.0'
        self.setWindowIcon(QIcon('image.png'))
        with open("style.qss", "r") as f:
            _style = f.read()
            app.setStyleSheet(_style)

        self.initUI()

    def initUI(self):
        self.flashSplash()
        time.sleep(1.5)

        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        #Build the Location Combo Box
        self.combo = QComboBox(self)
        self.create_locations()

        #Initialize the sheet selector combo box to be populated later
        self.combo1 = QComboBox(self)
        self.combo1.addItem("tbd")
        self.combo1.move(320, 120)

        #build the "instructions" label
        self.label = QLabel(self)
        self.label.setFont(QFont("Roboto", 10))
        self.label.setText("Select File. Then Location/Sheet. Then GO!")
        self.label.adjustSize()
        self.label.move(20, 20)

        #Build the label to show which file is selected. to be populated later.
        self.label1 = QLabel(self)
        self.label1.setText("No file selected yet.")
        self.label1.setFont(QFont("Roboto", 10))
        self.label1.adjustSize()
        self.label1.move(20, 85)

        #Build the label to show what file type was selected. Populated later.
        self.label2 = QLabel(self)
        self.label2.setText("TBD")
        self.label2.setFont(QFont("Roboto", 10))
        self.label2.adjustSize()
        self.label2.move(170, 120)
        self.label2.hide()

        #Create the button which triggers the file selector.
        button = QPushButton('File', self)
        button.move(20, 50)
        button.clicked.connect(self.select_file)
        #go to def "select file"

        #build the button which creates the next "filtered view" window.
        self.button1 = QPushButton('GO!', self)
        self.button1.move(170, 165)
        self.button1.hide()
        self.button1.clicked.connect(self.show_new_window)
        #go to def "show new window"

        self.show()

    def show_new_window(self):
        #Read the selected values from combo boxes
        self.chosen_sheet = self.combo1.currentText()
        self.location = self.combo.currentText()

        #create a datafarme from the chosen sheet
        data = self.df[self.chosen_sheet]

        #return cleaned data and the headers from the dataframe in order
        #to build the "filtered" view window.
        data, hdr = self.clean_data(data)

        #check to make sure items meeting that criteria were found
        if data.shape[0] < 1:
            msg = QMessageBox()
            msg.setWindowTitle("Complete")
            msg.setText("No Records found!")
            msg.adjustSize()
            msg.setWindowIcon(QIcon('image.png'))
            msg.exec_()
        else:
            #Createthe "filtered" view window.
            self.w = None
            self.w = FilteredWindow(data, hdr, self.location)
            self.w.show()

    def clean_data(self, data):
        """
        data : pandas dataframe
        -------
        data : pandas dataframe
        hdr : list
        """
        data = data.loc[:, ~data.columns.str.contains("Do Not Modify")]

        if self.chosen_sheet == 'B':
            data.rename(columns={"Req. Item #": "Item"}, inplace=True)

        ref = pd.read_csv("equipment.csv")
        ref['Material No'] = ref['Material No'].astype(
            str).apply(lambda x: x.split(".")[0])

        temp = ref.loc[(ref['Location'] == self.location)]

        data['Item'] = data['Item'].astype(
            str).apply(lambda x: x.split(".")[0])

        data = data.loc[(data['Item'].isin(temp['Material No']))]
        hdr = list(data.columns)

        return data, hdr

    def create_locations(self):
        self.combo.addItem("Unit1")
        self.combo.addItem("Unit2")
        self.combo.addItem("Unit3")
        self.combo.addItem("Unit4")
        self.combo.addItem("Unit5")
        self.combo.addItem("Unit6")
        self.combo.addItem("Unit7")
        self.combo.move(20, 120)

    def select_file(self):
        file, check = QFileDialog.getOpenFileName(None, "QFileDialog.getOpenFileName()",
                                                  "", "All Files (*);;Python Files (*.py);;Text Files (*.txt)")

        if check and ("MRP" in str(file) or "Client" in str(file)):
            self.df = pd.read_excel(file, None)
            msg = QMessageBox()
            msg.setWindowTitle("Success")
            msg.setText("File read succesfully!")
            msg.setWindowIcon(QIcon('image.png'))
            msg.exec_()
            self.label1.setText("File: " + str(file))
            self.label1.adjustSize()
            self.get_sheets()

            if "MRP" in str(file):
                self.df[list(self.df.keys())[0]].rename(
                    columns={"Material": "Item"}, inplace=True)
                self.label2.setText("MRP")
            else:
                self.label2.setText("Client Open")

            self.label2.show()
            self.label2.adjustSize()
            self.button1.show()
        else:
            msg = QMessageBox()
            msg.setWindowTitle("Uh Oh")
            msg.setText("Choose MRP or Client Open File!")
            msg.setWindowIcon(QIcon('image.png'))
            msg.exec_()

    def get_sheets(self):
        self.combo1.clear()
        self.sheets = self.df.keys()
        for sheet in self.sheets:
            self.combo1.addItem(sheet)

    def flashSplash(self):
        self.splash = QSplashScreen(QPixmap('logo.jpg'))
        self.splash.show()
        QTimer.singleShot(1500, self.splash.close)


class FilteredWindow(QWidget):

    def __init__(self, df, hdr, title):
        """
        df : pandas dataframe - filtered file
        hdr : list - headers for df
        title : string - location chosen from App main window combo box.
        """
        super().__init__()
        self.df = df
        self.hdr = hdr

        self.colwidth = 500
        self.left = 50
        self.top = 50
        self.width = 1500
        self.height = 500
        self.setWindowTitle(title)
        self.setWindowIcon(QIcon('image.png'))

        self.rownums = self.df.shape[0]
        self.colnums = self.df.shape[1]

        self.initUI()

    def initUI(self):

        self.setGeometry(self.left, self.top, self.width, self.height)
        self.createTable()
        self.layout = QGridLayout()
        self.layout.addWidget(self.tableWidget, 0, 0)
        self.setLayout(self.layout)
        self.show()

    def createTable(self):

       # Create table
        self.tableWidget = QTableWidget()
        self.tableWidget.setEditTriggers(QtWidgets.QTreeView.NoEditTriggers)

        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget.horizontalHeader().setVisible(True)

        #row and column count
        self.tableWidget.setRowCount(self.rownums)
        self.tableWidget.setColumnCount(self.colnums)
        self.tableWidget.setSizeAdjustPolicy(
            QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.tableWidget.setHorizontalHeaderLabels(self.hdr)

        #set column width
        for i in range(self.colnums):
            self.tableWidget.setColumnWidth(i, self.colwidth)

        #fill the table with co-ordinates
        for i in range(self.rownums):
            for j in range(self.colnums):
                cell = QTableWidgetItem(str([self.df.iloc[i, j]][0]))
                self.tableWidget.resizeColumnsToContents()
                self.tableWidget.setItem(i, j, cell)

        self.tableWidget.move(0, 30)
        self.tableWidget.viewport().installEventFilter(self)

    def eventFilter(self, source, event):
        if self.tableWidget.selectedIndexes() != []:

            if event.type() == QtCore.QEvent.MouseButtonRelease:
                if event.button() == QtCore.Qt.RightButton:
                    row = self.tableWidget.currentRow()
                    col = self.tableWidget.currentColumn()
                    self.rightClick(row, col)

        return QtCore.QObject.event(source, event)

    def rightClick(self, row, column):
        self.t = None
        self.ref1 = pd.read_csv("equipment.csv")
        self.ref1['Material No'] = self.ref1['Material No'].astype(
            str).apply(lambda x: x.split(".")[0])

        key = self.df.iloc[row, column]

        if column == self.df.columns.get_loc('Item'):
            ref1 = self.ref1.loc[self.ref1['Material No'] == key]
            ref1.drop(columns='Material No', inplace=True)
            hdr = list(ref1.columns)
            self.t = EquipmentWindow(ref1, hdr, str(key))
            self.t.show()


class EquipmentWindow(FilteredWindow):

    def __init__(self, df, hdr, title):
        """
        df : pandas dataframe - equipment matching selected equipment #
        hdr : list - header for df^
        title : string - equipment # selected
        """
        super().__init__(df, hdr, title)

    def rightClick(self):
        path, ok = QtWidgets.QFileDialog.getSaveFileName(
            self, 'Save CSV', os.getenv('HOME'), 'CSV(*.csv)')
        if ok:
            columns = range(self.tableWidget.columnCount())
            header = [self.tableWidget.horizontalHeaderItem(column).text()
                      for column in columns]
            with open(path, 'w') as csvfile:
                writer = csv.writer(
                    csvfile, dialect='excel', lineterminator='\n')
                writer.writerow(header)
                for row in range(self.tableWidget.rowCount()):
                    writer.writerow(
                        self.tableWidget.item(row, column).text()
                        for column in columns)

    def eventFilter(self, source, event):
        if self.tableWidget.selectedIndexes() != []:

            if event.type() == QtCore.QEvent.MouseButtonRelease:
                if event.button() == QtCore.Qt.RightButton:
                    self.rightClick()
        return QtCore.QObject.event(source, event)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
