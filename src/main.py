import sys, os
from PyQt5.QtWidgets import *
from PyQt5 import uic
import nltk
import openpyxl
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
import itertools

form_class = uic.loadUiType("./application.ui")[0]
input_path_list = []

class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.findpath.clicked.connect(self.findpathBtnListener)
        self.process_start_btn.clicked.connect(self.processBtnListener)
        
    def findpathBtnListener(self):
        global input_path_list
        fnames = list(QFileDialog.getOpenFileNames(self))
        if(fnames[0] == []):
            return
        input_path_list = fnames[0]
        path_text = ""
        for path in fnames[0]:
            path_text += path.split("/")[-1] + "\n"
        self.pathText.setText(path_text)
        
    def processBtnListener(self):
        global input_path_list
        if(input_path_list == []):
            return
        for path in input_path_list:
            fine_text = self.read_txt(path)
            fine_text  = fine_text.split('"')
            fine_text=[nltk.sent_tokenize(fine_text) for fine_text in fine_text]
            fine_text=[element for array in fine_text for element in fine_text]
            if(len(fine_text) > 1):
                fine_text = list(itertools.chain(*fine_text))
            
            self.create_xlsx(path, fine_text)
        
        self.msg = QMessageBox()
        self.msg.setIcon(QMessageBox.Information)
        self.msg.setWindowTitle('Success')
        self.msg.setText('작업 완료.')
        self.msg.setStandardButtons(QMessageBox.Ok)
        retval = self.msg.exec_()
            
    def read_txt(self, path):
        with open(path, "rt", encoding='UTF8') as file:
            text=file.read().replace('\n',' ').replace('  ',' ')
        return text;
    
    def create_xlsx(self, path, data):
        global input_path_list
        wb = openpyxl.Workbook()
        ws = wb.active
        for row in data:
            try:
                ws.append([row])
            except:
                ws.append([ILLEGAL_CHARACTERS_RE.sub(r'', row)])
        save_path = path.split(".")[0] + "_result.xlsx"
        wb.save(save_path)
        wb.close()
        
        input_path_list = []
        self.pathText.setText('')
        return True

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.setWindowTitle("NAN ROW TXT TO EXCEL")
    myWindow.show()
    app.exec_()