import xlrd 
from output import Ui_MainWindow
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import sys
import os

class UI(QMainWindow,Ui_MainWindow):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
    
        self.chooseIn_btn.clicked.connect(self.openfile)
        self.chooseIn_btn_2.clicked.connect(self.openfile2)
        self.chooseOut_btn.clicked.connect(self.savefile)
        self.start_btn.clicked.connect(self.start)

        iconPath = "favicon.png"
        icon = QIcon(iconPath)
        icon.addPixmap(QPixmap(iconPath), QIcon.Disabled)
        self.setWindowIcon(icon)

        self.setWindowTitle("Excel to HTML")

        self.logo = QLabel(self)
        pixmap = QPixmap('logo.png')
        pixmap=pixmap.scaledToWidth(360)
        self.logo.setPixmap(pixmap)
        self.logo.setGeometry(420,80,351,71)

        self.web_path=None
        self.app_path=None
        self.html_path=None

    def openfile(self):
        filename = QFileDialog.getOpenFileName()
        path = filename[0]
        self.input_addr.setText(path)
        self.web_path=path

    def openfile2(self):
        filename = QFileDialog.getOpenFileName()
        path = filename[0]
        self.input_addr_2.setText(path)
        self.app_path=path
    
    def savefile(self):
        filename = QFileDialog.getSaveFileName()
        path = filename[0]
        self.output_addr.setText(path)
        self.html_path=path

    def start(self):
        if (not self.web_path) or (not self.app_path):
            return
        web = xlrd.open_workbook(self.web_path) 
        sheet_web = web.sheet_by_index(0)
        
        app = xlrd.open_workbook(self.app_path) 
        sheet_app = app.sheet_by_index(0)

        group=None
        iapp=0
        html=""
        for i in range (0, sheet_web.nrows):
            row=sheet_web.row_values(i)
            if group!= row[0]:
                while (sheet_app.row_values(iapp)[0]==group):
                    row_app=sheet_app.row_values(iapp)
                    a_tag='<a href="%s">%s</a><br>\n'%(row_app[2],row_app[1])
                    html+=a_tag
                    iapp+=1
                group= row[0]
                text_group='<br>%s<br>\n'%(row[0])
                html+=text_group
            a_tag='<a href="%s">%s %s</a><br>\n'%(row[2],row[1],row[2])
            html+=a_tag
        self.code.setPlainText(html)
        self.html.setHtml(html)

        if self.html_path:
            f = open(self.html_path, "w",encoding='utf-8')
            f.write(html)
            f.close()

app = QApplication(sys.argv)
ui = UI()
ui.show()
app.exec_()