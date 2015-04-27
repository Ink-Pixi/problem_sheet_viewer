import sys 
import ctypes
import os
import winreg
import subprocess
import time
from data import DataBase
from ui import Ui_ReportViewer
from PyQt5.QtWidgets import QApplication, QMessageBox, QMainWindow,\
    QInputDialog
from PyQt5.QtCore import QUrl, pyqtSignal
from PyQt5.QtPrintSupport import QPrinter, QPrintPreviewDialog, QPrinterInfo
from PyQt5.QtGui import QPagedPaintDevice

from PyQt5.Qt import QKeySequence


class ReportViewer(QMainWindow, Ui_ReportViewer):
    reportClosed = pyqtSignal()
    
    def __init__(self, ID):
        self.ID = ID
        
        super(ReportViewer, self).__init__()
        
        self.setupUi(self)
        self.db = DataBase()        
        
        self.btnMarkStatus.clicked.connect(self.btnMarkStatus_Clicked)
        
        self.actionPrint.triggered.connect(self.print_sheet)
        self.actionPrint.setShortcut(QKeySequence.Print)
        
        self.actionOpen.triggered.connect(self.open_sheet)
        self.actionOpen.setShortcut(QKeySequence.Open)
        
        self.actionExit.triggered.connect(self.exit_sheet)
        self.actionExit.setShortcut(QKeySequence.Close)
        
        self.load_report()
        

    def load_report(self):
        self.setWindowTitle('Problem Sheet:  #' + str(self.ID))
        self.lblProbNum.setText(str(self.ID))
        
        rptType = self.db.get_report_type(self.ID)
        self.set_button_txt(rptType)
        self.set_lblFlag_txt()
        
        self.set_report(rptType)
                
        
    def closeEvent(self, evnt):
        #Built in QT override method to catch close event.
        self.reportClosed.emit()
        self.close()
        
    def set_report(self, rptType):
       
        self.webRptView.load(QUrl('http://sqlrptserver/ReportServer_SQLREPORTs?%2fPython+Problem+Sheets%2f'+rptType+'&ID='+str(self.ID)+'&rs:Command=Render&rc:Parameters=Collapsed&rc:Toolbar=false'))
        
        if self.db.check_credit_info(self.ID):
            self.lblCreditInfo.setText('Credit Card Available')
        
        if rptType == "IP Expedited":
            if not self.db.check_expedited_approval(self.ID):
                self.lblAprove.setText('Needs Approval by shift leader or accounting')

        
    def set_button_txt(self, rptType):
        inProgress = self.db.check_progress(self.ID)  
        isComplete = self.db.check_complete(self.ID)
        
        if inProgress:
            self.btnMarkStatus.setText('Mark Complete')
            self.lblState.setText(rptType + ' - In Progress')
            self.state = 'inProgress'
        elif isComplete:
            self.btnMarkStatus.setText('Re-Open')
            self.lblState.setText(rptType + ' - Closed')
            self.state = 'complete'
        else:
            self.btnMarkStatus.setText('Mark In Progress')
            self.lblState.setText(rptType + ' - Open')
            self.state = 'open'
            
    def btnMarkStatus_Clicked(self):
        if self.state == 'open':
            try:
                self.db.mark_in_progress(self.ID)
            except BaseException as e:
                self.error_message(e)
        elif self.state == 'inProgress':
            try:
                self.db.mark_complete(self.ID)
            except:
                self.error_message(e)
        elif self.state == 'complete':
            try:
                self.db.mark_open(self.ID)
            except BaseException as e:
                self.error_message(e)
         
        self.set_button_txt(self.db.get_report_type(self.ID))

    def print_sheet(self):
        printer = QPrinter(QPrinterInfo.defaultPrinter())
        printer.setOutputFileName('foo.pdf')
        printer.setResolution(72)
        printer.setFullPage(True)

        self.webRptView.print(printer)

        pdf = 'foo.pdf'
            
        AcroRD32Path = winreg.QueryValue(winreg.HKEY_CLASSES_ROOT, 'Software\\Adobe\\Acrobat\Exe')
            
        acroread = AcroRD32Path
            
        cmd= '{0} /N /T "{1}" ""'.format(acroread,pdf)  
        proc = subprocess.Popen(cmd)
        time.sleep(5)
           
        os.system('TASKKILL /F /IM AcroRD32.exe')
        os.remove('foo.pdf')   
         
    def set_lblFlag_txt(self):
        if self.db.check_priority(self.ID) == True and self.db.check_trace(self.ID):
            self.lblFlag.setText('*** Priority Trace ***')
        elif self.db.check_priority(self.ID) == True:
            self.lblFlag.setText('*** Priority ***')
        elif self.db.check_trace(self.ID):
            self.lblFlag.setText('*** For Trace ***')            
        else: self.lblFlag.setText('')
    
    def error_message(self, error):
        mBox = QMessageBox()
        mBox.critical(self, 'Error', 'An Error has occurred please contact IT. \n Error Msg: ' + str(error), mBox.Ok)
        
    def exit_sheet(self):
        self.close()
        
    def open_sheet(self):
        inNumber, ok = QInputDialog.getText(self, 'Please Enter Number', 'Please enter either long number or short number.')
        if ok: 
            if inNumber:
                txtNumber = inNumber.replace('-', '')
                try:
                    self.ID = self.db.get_sheet_id(txtNumber)
                    self.load_report()
                except TypeError:
                    QMessageBox.information(self, 'Error', 'Could not find problem sheet for number enter, please make sure it is typed correctly.')
            else:
                print('nothing entered')
        else:
            pass
    
            
        
if __name__ == '__main__':
    myappid = 'Main Viewer' # arbitrary string
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)     
    
    app = QApplication(sys.argv)    
    
    if len(sys.argv) == 1:
        argID = '1'
    else: argID = sys.argv[1]
    
    mr = ReportViewer(argID)
    mr.show()
    
    sys.exit(app.exec_())    