import sys, os
from PyQt5.QtWidgets import *
from PyQt5 import uic
import nltk

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
            print(fine_text[0])
            
    def read_txt(self, path):
        with open(path, "rt", encoding='UTF8') as file:
            text=file.read().replace('\n',' ').replace('  ',' ')
        return text;

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.setWindowTitle("TXT TO EXCEL")
    myWindow.show()
    app.exec_()