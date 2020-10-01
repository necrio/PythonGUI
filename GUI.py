import sys
from PySide2.QtCore import Qt, Slot
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import (QAction, QApplication, QHeaderView, QHBoxLayout, QLabel, QLineEdit,
                               QMainWindow, QPushButton, QTableWidget, QTableWidgetItem,
                               QVBoxLayout, QWidget, QTableView, QGridLayout, QGroupBox)
from PyQt5 import  uic
import pandas as pd
from openpyxl import load_workbook
import pyperclip as pc


class Widget(QtWidgets.QMainWindow):
    def __init__(self):        
        super(Widget, self).__init__() # Call the inherited classes __init__ method
        uic.loadUi('F:\\PULLDATA\Main.ui', self) # Load the .ui file
       
        
        ## Items 
        self.items = 0
        
        #Reading Data with Pandas. 
        self._data = pd.read_excel('F:\\PULLDATA')
        #self._data = pd.read_csv('data.csv')
        ## Writing with openpyxl to edit.
        self.wb = load_workbook(filename = 'F:\\GPULLDATA')
        self.ws = self.wb['Sheet1']     
       
        #######UI Build #######
        self.table = self.findChild(QtWidgets.QTableView, 'tableView')
        
        
        
         ###### Left
        # self.table = QTableView(self)       
        
        self.model = PandasModel(self._data)
        
        self.table.setModel(self.model)
        self.table.resizeColumnsToContents()
         

        ######### Signals and Slots
        self.addNewBtn.clicked.connect(self.add_element)
        self.copyBtn.clicked.connect(self.copyToClipboard)
     
    
        ## Filter area
        self.filter_proxy_model = QtCore.QSortFilterProxyModel()
        self.filter_proxy_model.setSourceModel(self.model)
        self.filter_proxy_model.setFilterKeyColumn(0) # First column
        self.searchLine.textChanged.connect(self.filter_proxy_model.setFilterRegExp)
        self.table.setModel(self.filter_proxy_model)
        
        ### Select Whole row if click on item.
        self.selected = self.table.setSelectionBehavior(QTableView.SelectRows)
        
        #### TEST AREA FOR SELECTION ####
        self.table.clicked.connect(self.sendToAlarmGen)
        
        
        
    @Slot()
    def sendToAlarmGen(self, index):
        
        self.datas = str(self._data.loc[index.row()].at['Alarm Note'])        
        self.alarmGen.setText(self.datas)
        # print(self.datas)
    
        
    
        
    @Slot()
    def add_element(self):
        ##TODO: Add Diaglog to add new stuff.
        pass

    @Slot()
    def copyToClipboard(self):
        pc.copy(self.datas)
        # print(self.datas)
        
        
class MainWindow(QMainWindow):
    def __init__(self, widget):
        QMainWindow.__init__(self)
        self.setWindowTitle("Version 0.0.1 Testing")

        # Menu
        # self.menu = self.menuBar()
        # self.file_menu = self.menu.addMenu("File")

        # Exit QAction
        # exit_action = QAction("Exit", self)
        # exit_action.setShortcut("Ctrl+Q")
        # exit_action.triggered.connect(self.exit_app)

        # self.file_menu.addAction(exit_action)
        self.setCentralWidget(widget)

    
        
class PandasModel(QtCore.QAbstractTableModel): 
    def __init__(self, df = pd.DataFrame(), parent=None): 
        QtCore.QAbstractTableModel.__init__(self, parent=parent)
        self._df = df

    def headerData(self, section, orientation, role=QtCore.Qt.DisplayRole):
        if role != QtCore.Qt.DisplayRole:
            return QtCore.QVariant()

        if orientation == QtCore.Qt.Horizontal:
            try:
                return self._df.columns.tolist()[section]
            except (IndexError, ):
                return QtCore.QVariant()
        elif orientation == QtCore.Qt.Vertical:
            try:
                # return self.df.index.tolist()
                return self._df.index.tolist()[section]
            except (IndexError, ):
                return QtCore.QVariant()

    def data(self, index, role=QtCore.Qt.DisplayRole):
        if role != QtCore.Qt.DisplayRole:
            return QtCore.QVariant()

        if not index.isValid():
            return QtCore.QVariant()

        return QtCore.QVariant(str(self._df.iloc[index.row(), index.column()]))

    def setData(self, index, value, role):
        row = self._df.index[index.row()]
        col = self._df.columns[index.column()]
        if hasattr(value, 'toPyObject'):
            # PyQt4 gets a QVariant
            value = value.toPyObject()
        else:
            # PySide gets an unicode
            dtype = self._df[col].dtype
            if dtype != object:
                value = None if value == '' else dtype.type(value)
        self._df.set_value(row, col, value)
        return True

    def rowCount(self, parent=QtCore.QModelIndex()): 
        return len(self._df.index)

    def columnCount(self, parent=QtCore.QModelIndex()): 
        return len(self._df.columns)

    def sort(self, column, order):
        colname = self._df.columns.tolist()[column]
        self.layoutAboutToBeChanged.emit()
        self._df.sort_values(colname, ascending= order == QtCore.Qt.AscendingOrder, inplace=True)
        self._df.reset_index(inplace=True, drop=True)
        self.layoutChanged.emit()


        
if __name__ == "__main__":
    # Qt Application
    app = QApplication(sys.argv)
    # QWidget
    widget = Widget()
    # QMainWindow using QWidget as central widget
    window = MainWindow(widget)
    #window = uic.loadUi("Test1.ui")
    window.resize(1100, 600)
    window.show()

    # Execute application
    sys.exit(app.exec_())        
