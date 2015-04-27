import sys, ctypes
from data import DataBase
from ui import Ui_MainReport
from report_viewer import ReportViewer
from PyQt5.QtWidgets import QMainWindow, QTableWidgetItem, QApplication, QDialog,\
    QMessageBox
from PyQt5.QtGui import QFont, QPalette
from PyQt5.QtCore import Qt, QUrl, QThread, pyqtSignal
from PyQt5 import QtNetwork, QtWebKit, QtPrintSupport #Needed to run after compiling with cx_freeze.

class MainReport(QMainWindow, Ui_MainReport):
    tblState = 0
    def __init__(self):
        super(MainReport, self).__init__()
        self.db = DataBase()
        
        self.setupUi(self)
        
        self.create_table(self.tblState)
        self.tblSummary.resizeColumnsToContents()
        self.tblSummary.cellPressed.connect(self.item_clicked)  
        self.tblSummary.setSortingEnabled(True)
        self.tblSummary.show()          
        
        self.btnToggleTable.clicked.connect(self.btnToggleTable_Clicked)
        self.btnOpen.clicked.connect(self.btnOpen_Clicked)

    def create_table(self, tblType):
        
        data = self.db.get_table_data(tblType)
            
        font = QFont('Veranda', 12, QFont.Bold)
        blk = QPalette()
        blk.setColor(blk.Foreground, Qt.black)      
   
        # Check to make sure the list is there and has data, then we go through it and add data to the table.
        if data:
            self.tblSummary.setRowCount(len(data))
            for i, row in enumerate(data):
                for j, col in enumerate(row):
                    item = QTableWidgetItem(str(col))
                    item.setFont(font)
                    #item.setForeground(QColor.fr
                    #item.setFlags(Qt.ItemIsEditable)
                    if item.text() == "None":
                        item.setText("")
                    self.tblSummary.setItem(i, j, item)      
        else:
            self.tblSummary.setRowCount(1)
            item = QTableWidgetItem()
            item.setText("Nothing Found")
            self.tblSummary.setItem(0, 0, item)

    def btnToggleTable_Clicked(self):
        btn = self.sender()

        if btn.objectName() == 'btnToggleTable':
            self.tblState = 1
            self.create_table(1)
            
            btn.setText('View Open')
            btn.setObjectName('btnOpen')
            
            lstHeader = ["ID", "Company", "Type", "Order\Invoice #", "Date Completed", "Priority?", "Completed By"]     
                        
        elif btn.objectName() == 'btnOpen':
            self.create_table(0)
            self.tblState = 0
            
            btn.setText('View Completed')
            btn.setObjectName('btnToggleTable')
            
            lstHeader = ["ID", "Company", "Type", "Order\Invoice #", "Call Date", "Priority?", "Status"]
            
        self.tblSummary.setHorizontalHeaderLabels(lstHeader)
        self.tblSummary.resizeColumnsToContents()
            
        
    def item_clicked(self, row, col):
        ID = self.tblSummary.item(row, 0).text()
        
        self.rv = ReportViewer(ID)
        self.rv.reportClosed.connect(lambda: self.create_table(self.tblState))        
        self.rv.show()
        
    def btnOpen_Clicked(self):
        if self.leID.text():
            searchNumber = self.leID.text().replace("-", "")

            try: 
                ID = self.db.get_sheet_id(searchNumber)
                self.rv = ReportViewer(ID)
                self.rv.show()
                self.leID.setText(None)
            except BaseException:
                QMessageBox.information(self, 'Doh!', "Could not find problem sheet for the number entered, please make sure it is typed correctly."  \
                    " If you feel you reached this in error, please contact IT.")
                
        else:
            QMessageBox.information(self, 'Enter number', 'Please enter a value to search.')
            
if __name__ == '__main__':
    myappid = 'Main Viewer' # arbitrary string
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)     
    
    app = QApplication(sys.argv)    
    
    mr = MainReport()
    mr.show()
    
    sys.exit(app.exec_())