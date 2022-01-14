from lib2to3.pytree import convert
import sys
import os
import time
import shutil
from tkinter.tix import Tree

from PyQt5.QtCore import Qt, QSize, center
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import QApplication, QListWidget, QMainWindow, QVBoxLayout, QWidget, \
                                QHBoxLayout, QPushButton, QLineEdit, QLabel, QFileDialog, QMessageBox, QDialog

from pdf2docx import Converter

class MainWindow(QMainWindow):
    def __init__(self):
        self.pdf_list = []
        self.filedir = ""
        self.working_dir = ""
        #self.currentdir = ""
        super().__init__()
        self.setWindowTitle("PDF TO WORD Convertor")
        self.setWindowIcon(QIcon("inpng.ico"))
        #self.resize(QSize(1000,600))
        self.setFixedSize(QSize(1200, 600))

        pixmap = QPixmap("thelogo.png")

        pixmap2 = pixmap.scaled(300, 300, Qt.KeepAspectRatio)
        ##trying to add a picture
        label1 = QLabel("Hello")
        label1.setPixmap(pixmap2)
        
        button = QPushButton("Open Folder")
        button.setStyleSheet('''
            font-size: 25px;
            width: 200px;
            height: 50;
        ''')
        button.clicked.connect(self.openFile)
        self.button2 = QPushButton("Convert!")
        self.button2.setStyleSheet('''
            font-size: 25px;
            width: 150px;
            height: 50;
        ''')
        #self.button2.clicked.connect(self.changelabel)
        self.button2.clicked.connect(self.convertfile)
        self.button3 = QPushButton("Delete")
        self.button3.setStyleSheet('''
            font-size: 25px;
            width: 150px;
            height: 50;
        ''')
        self.button3.clicked.connect(self.deleteSelected)
        self.button4 = QPushButton("Clear")
        self.button4.setStyleSheet('''
            font-size: 25px;
            width: 150px;
            height: 50;
        ''')
        self.button4.clicked.connect(self.clearSelected)
        button5 = QPushButton("Save to")
        button5.setStyleSheet('''
            font-size: 25px;
            width: 200px;
            height: 50;
        ''')
        button5.clicked.connect(self.saveto)
        self.listwid = QListWidget(self)
        self.listwid.setStyleSheet('''
                border: 2px solid black;
                font-size: 18px;
        ''')

        self.label_file_save_to = QLineEdit("Save File to", self)
        self.label_file_save_to.setReadOnly(True)
        self.label_file_save_to.setStyleSheet('''
                border: 2px solid black;
                font-size: 20px;
        ''')

        self.label_currentprogress = QLineEdit("Select File Path and File Save Location")
        self.label_currentprogress.setReadOnly(True)
        self.label_currentprogress.setStyleSheet('''
                font-size: 20px;
                border: 2px solid black;
        ''')
        ##widgets in a vertical direction
        layoutvertleft = QVBoxLayout()
        self.layoutverticalbottom = QVBoxLayout()

        ##widgets in a horizontal direction
        layouthorizontaltop = QHBoxLayout()
        
        layoutvertleft.addWidget(label1)
        layoutvertleft.addWidget(button)
        layoutvertleft.addWidget(self.button2)
        layoutvertleft.addWidget(self.button3)
        layoutvertleft.addWidget(self.button4)
        layoutvertleft.addWidget(button5)

        layouthorizontaltop.addLayout(layoutvertleft)
        layouthorizontaltop.addWidget(self.listwid)

        self.layoutverticalbottom.addLayout(layouthorizontaltop)
        self.layoutverticalbottom.addWidget(self.label_file_save_to)
        self.layoutverticalbottom.addWidget(self.label_currentprogress)
        #self.inital()
        if self.pdf_list == []:
            self.button2.setEnabled(False)
            self.button3.setEnabled(False)
            self.button4.setEnabled(False)

        self.openvalue = 0
        self.savevalue = 0
        #print(self.pdf_list)
        widget = QWidget()
        widget.setLayout(self.layoutverticalbottom)
        self.setCentralWidget(widget)

    def openFile(self):
            self.file_path, s = QFileDialog.getOpenFileNames(self, 'Open PDF file', os.getcwd(), 'PDF file(*.pdf)')
            self.working_dir = os.getcwd()
            self.currentdir = self.working_dir
            print(self.currentdir)
            for i in range(len(self.file_path)):
                #print(self.file_path[i])
                self.pdf_list.append(self.file_path[i])

            #print(pdf_list)
            
            self.listwid.addItems(self.pdf_list)
            
            if len(self.pdf_list) > 0:
                #print(i)
                self.openvalue = 2

            else: 
                self.openvalue = 0
                #self.label_currentprogress.setText("Please Choose File Location")

            if self.openvalue + self.savevalue >= 4:
                self.label_currentprogress.setText("Ready To Convert")
                self.button2.setEnabled(True)
                self.button3.setEnabled(True)
                self.button4.setEnabled(True)

            #print(self.value)
            #print(self.pdf_list)
            #print(file_path)
            return print(self.pdf_list)

    def deleteSelected(self):
            for item in self.listwid.selectedItems():
                #print(self.listwid.currentItem().text())
                self.pdf_list.remove(self.listwid.currentItem().text())
                self.listwid.takeItem(self.listwid.row(item))
            #print(self.pdf_list)

            return self.pdf_list

    def clearSelected(self):
        self.listwid.clear()
        self.pdf_list.clear()
        #print(self.pdf_list)
        return self.pdf_list

    def saveto(self):
        self.filedir = QFileDialog.getExistingDirectory(self, 'Save PDF file to')
        
        #print(self.filedir)

        if self.filedir  != "":
            self.savevalue = 2
        else:
            self.savevalue = 0
            self.button2.setEnabled(False)
            self.button3.setEnabled(False)
            self.button4.setEnabled(False)
        app.processEvents()
        if (self.savevalue + self.openvalue) >= 4:
            self.label_currentprogress.setText("Ready To Convert")
            self.button2.setEnabled(True)
            self.button3.setEnabled(True)
            self.button4.setEnabled(True)

        #print(self.currentdir)
        print(self.filedir)
        #print(self.value)
        self.label_file_save_to.setText("Files Saved at " + str(self.filedir))

        #print(self.label_file_save_to.text())
        return self.filedir
    
    def convertfile(self):
        app.processEvents()
        self.button2.setEnabled(False)
        self.button3.setEnabled(False)
        self.button4.setEnabled(False)        
        self.label_currentprogress.setText("Converting...")
        app.processEvents()
        self.workingConversion()

        self.button2.setEnabled(False)
        self.button3.setEnabled(False)
        self.button4.setEnabled(False)
        self.openvalue = 0
        #self.value = 1
        dlg2 = QMessageBox(self)
        dlg2.setWindowTitle("Conversion Complete!")
        dlg2.setText("<center>The conversion is complete! \n</center>" + "<center>Open Converted Files Folder?</center>")
        dlg2.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        dlg2.setIcon(QMessageBox.Question)
        button = dlg2.exec_()

        self.label_currentprogress.setText("Select File Path and File Save Location")
        if button == QMessageBox.Yes:
            data = 'start ' + self.filedir
            print(data)
            os.system(data)
            print("Yes!")
            self.clearSelected()
        else:
            print("No!")
            dlg2.close()
            self.clearSelected()


    def workingConversion(self):
        for i in range(len(self.pdf_list)):
            app.processEvents()
            pdf_file = self.pdf_list[i]
            name = self.pdf_list[i]
            print(os.path.dirname(pdf_file))
            #if os.path.abspath(pdf_file) == self.filedir
            #shutil.copy(pdf_file,self.filedir)
            cv = Converter(name)
            #print(pdf_file)
            newname = name.replace(".pdf", ".docx")
            cv.convert(newname)
            #print(self.currentdir)
            if os.path.dirname(pdf_file) == self.filedir:
                print("A")
                #shutil.copy(newname,self.filedir)
            else:
                shutil.copy(newname,self.filedir)
                os.chdir(self.working_dir)
                os.remove(newname)
            time.sleep(1)
            cv.close()


app = QApplication(sys.argv)

window = MainWindow()
window.show()

app.exec_()
