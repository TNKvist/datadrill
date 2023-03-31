"""
@author: tkvist1
"""
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, 
                             QFileDialog, QWidget, QLabel, QMessageBox, 
                             QComboBox, QTableWidget, QGridLayout, 
                             QTableWidgetItem,QSplashScreen)
import pandas as pd
import sys,os,csv
import time
from PyQt5 import QtWidgets
from PyQt5 import QtCore
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QFont,QIcon,QPixmap


class test(object):
    def __init__(self):
        self.t = None
   
    def rightClick(self, row, column,df):
        key = df.iloc[row,column]
        if column == df.columns.get_loc('Item'):
            final = ref1.loc[ref1['Material No'] == key]
            final.drop(columns='Material No',inplace=True)
            hdr = list(final.columns) 
            self.t = FinalWindow(final,hdr,str(key))
            self.t.show()
 
class AnotherWindow(QWidget):
 
    def __init__(self,df,hdr,title):
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
        self.test = test()
           
    def initUI(self):
        
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.createTable()
        self.layout = QGridLayout()
        self.layout.addWidget(self.tableWidget,0,0)
        self.setLayout(self.layout) 
        # Show widget
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
        self.tableWidget.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
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
                
        self.tableWidget.move(0,30)
        self.tableWidget.viewport().installEventFilter(self)
        
    def eventFilter(self, source, event):
        if self.tableWidget.selectedIndexes() != []:
            
            if event.type() == QtCore.QEvent.MouseButtonRelease:
                if event.button() == QtCore.Qt.RightButton:
                    row = self.tableWidget.currentRow()
                    col = self.tableWidget.currentColumn()
                    self.test.rightClick(row, col,self.df)
       
        return QtCore.QObject.event(source, event)
    
class FinalWindow(QWidget):
 
    def __init__(self,df,hdr,title):
        super().__init__()
        self.df = df
        self.hdr = hdr
        self.colwidth = 500
        self.left = 300
        self.top = 250
        self.setWindowIcon(QIcon('image.png'))
        self.width = 250
        self.height = 600
        self.setWindowTitle(title)
        self.rownums = self.df.shape[0]
        self.colnums = self.df.shape[1]
        self.initUI()
           
    def initUI(self):
        
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.createTable()
        self.layout = QGridLayout()
        
        self.layout.addWidget(self.tableWidget,0,0)
        self.setLayout(self.layout) 
        # Show widget
        self.show()
        
    def eventFilter(self, source, event):
        if self.tableWidget.selectedIndexes() != []:
            
            if event.type() == QtCore.QEvent.MouseButtonRelease:
                if event.button() == QtCore.Qt.RightButton:
                    self.handleSave()       
        return QtCore.QObject.event(source, event)     
    
    def handleSave(self):
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

    def createTable(self):

       # Create table
        self.tableWidget = QTableWidget()
        self.tableWidget.setEditTriggers(QtWidgets.QTreeView.NoEditTriggers) 

        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget.horizontalHeader().setVisible(True)

        #row and column count
        self.tableWidget.setRowCount(self.rownums)
        self.tableWidget.setColumnCount(self.colnums)
        self.tableWidget.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.tableWidget.setHorizontalHeaderLabels(self.hdr)
        

        #set column width
        for i in range(self.colnums):
            self.tableWidget.setColumnWidth(i, self.colwidth)
            
        #fill the table with co-ordinates
        for i in range(self.rownums):
            for j in range(self.colnums):
                st = str([self.df.iloc[i, j]][0])
                if j == 1:
                    st = '_' + st
                cell = QTableWidgetItem(st)
                self.tableWidget.resizeColumnsToContents()
                self.tableWidget.setItem(i, j, cell)
                
        self.tableWidget.move(0,30)
        self.tableWidget.viewport().installEventFilter(self)
         
    def eventFilter(self, source, event):
        if self.tableWidget.selectedIndexes() != []:
            
            if event.type() == QtCore.QEvent.MouseButtonRelease:
                if event.button() == QtCore.Qt.RightButton:
                    self.handleSave() 
        return QtCore.QObject.event(source, event)       
        
class App(QMainWindow):

    def __init__(self):
        super().__init__()
        self.sheets = None
        self.location = None
        self.chosen_sheet = None
        self.df = None
        self.title = 'Data Drill beta'
        self.left = 300
        self.top = 400
        self.width = 500
        self.height = 200
        self.setWindowIcon(QIcon('image.png'))
        self.w = None
        self.setMinimumSize(self.width, self.height)
        self.setMaximumHeight(self.height)
        with open("style.qss", "r") as f:
            _style = f.read()
            app.setStyleSheet(_style)
        self.initUI()

    
    def initUI(self):
        self.flashSplash()
        time.sleep(1.5)
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        
        self.combo = QComboBox(self)
        self.combo1 = QComboBox(self)
        self.combo1.addItem("tbd")
        self.combo1.move(320,120)
        self.create_locations()
        self.label = QLabel(self)
        self.label.setFont(QFont("Roboto", 10))
        self.label.setText("Select File. Then Location/Sheet. Then GO!")
        self.label.adjustSize()
        self.label.move(20,20)
        self.label1 = QLabel(self)
        self.label1.setText("No file selected yet.")
        self.label1.setFont(QFont("Roboto", 10))
        self.label1.adjustSize()
        self.label1.move(20,85)
        #button.setToolTip('This is an example button')
        button = QPushButton('File', self)
        button.move(20,50)
        button.clicked.connect(self.select_file)
        self.button1 = QPushButton('GO!', self)
        self.button1.move(170,165)
        self.button1.hide()
        self.button1.clicked.connect(self.show_new_window)
        
        self.show()

    def show_new_window(self):
        self.chosen_sheet = self.combo1.currentText()
        self.location = self.combo.currentText()
        data = self.df[self.chosen_sheet]
        data,hdr = self.clean_data(data) 
        if data.shape[0] < 1:
            msg = QMessageBox()
            msg.setWindowTitle("Complete")
            msg.setText("No Records found!")
            msg.adjustSize()
            msg.setWindowIcon(QIcon('image.png'))
            msg.exec_()           
        else:        
            self.w = None
            self.w = AnotherWindow(data,hdr,self.location)
            self.w.show()
        
    def clean_data(self,data):
        data = data.loc[:, ~data.columns.str.contains("Do Not Modify")]
        if self.chosen_sheet == 'MRs':
            data.rename(columns={"Req. Item #": "Item"},inplace=True)
        temp = ref.loc[(ref['Location'] == self.location)]
        data['Item'] = data['Item'].astype(str).apply(lambda x: x.split(".")[0])
        data = data.loc[(data['Item'].isin(temp['Material No']))]
        hdr = list(data.columns)
        return data,hdr
    
    def create_locations(self):
        self.combo.addItem("Zinc")
        self.combo.addItem("Calcium Phenate")
        self.combo.addItem("VISJ")
        self.combo.addItem("Blending")
        self.combo.addItem("Dispersants")
        self.combo.addItem("VisJ")
        self.combo.addItem("Thermal")
        self.combo.move(20,120)
    
    def select_file(self):
        file , check = QFileDialog.getOpenFileName(None, "QFileDialog.getOpenFileName()",
                                                   "", "All Files (*);;Python Files (*.py);;Text Files (*.txt)")
        if check:
            self.df = pd.read_excel(file, None)
            msg = QMessageBox()
            msg.setWindowTitle("Success")
            msg.setText("File read succesfully!")
            msg.setWindowIcon(QIcon('image.png'))
            msg.exec_()
            self.label1.setText("File: " + str(file))
            self.label1.adjustSize()
            self.get_sheets()
            self.button1.show()
    
    def get_sheets(self):
        self.sheets = self.df.keys()
        for sheet in self.sheets:
            self.combo1.addItem(sheet)
        self.combo1.removeItem(0)
        #self.combo1.adjustSize()
        
    def flashSplash(self):
        self.splash = QSplashScreen(QPixmap('logo.jpg'))

        # By default, SplashScreen will be in the center of the screen.
        # You can move it to a specific location if you want:
        # self.splash.move(10,10)

        self.splash.show()

        # Close SplashScreen after 2 seconds (2000 ms)
        QTimer.singleShot(1500, self.splash.close)
           
if __name__ == '__main__':
    global ref
    global ref1
    ref = pd.read_csv("test.csv")
    ref1 = pd.read_csv("output.csv")
    ref['Material No']= ref['Material No'].astype(str).apply(lambda x: x.split(".")[0])
    ref1['Material No']= ref1['Material No'].astype(str).apply(lambda x: x.split(".")[0])
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_()) 
