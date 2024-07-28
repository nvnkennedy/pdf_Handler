import sys
import os
from PyQt6 import QtWidgets, QtTest
from files.pdf_UI import *
from pdf2docx import Converter
from PyQt6.QtWidgets import QFileDialog
from PyQt6.QtGui import QDesktopServices, QIcon
from PyQt6.QtCore import QUrl
from qt_thread_updater import get_updater
from threading import Thread
from pathlib import Path
import subprocess
import pikepdf
import pypdf
from pikepdf import Pdf
class pdf_Handler(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super(pdf_Handler, self).__init__(parent)
        self.ui = Ui_pdf_Handler()
        self.ui.setupUi(self)
        self.setWindowIcon(QIcon(stCwd+"\\files\\pdficon.ico"))
        self.setFixedSize(599, 294)
        self.ui.textBrowser.anchorClicked.connect(QtGui.QDesktopServices.openUrl)
        self.ui.textBrowser.setOpenLinks(False)
        self.ui.pdf_Doc_Button.clicked.connect(self.convertPdfToDoc)
        self.ui.pdf_Powerpoint_Button.clicked.connect(self.convertPdfToPowerPoint)
        self.ui.pdf_Unlock_Button.clicked.connect(self.unlockPdf)
        self.ui.pdf_Merge_Button.clicked.connect(self.mergePdf)
    def convertPdfToDoc(self):
        self.pdffileName = self.openPDFFileNameDialog()
        self.disableObject(self.ui.pdf_Doc_Button)
        self.Thread = Thread(target=self.pdf_Doc, daemon=True)
        self.Thread.start()
    def convertPdfToPowerPoint(self):
        self.pdffileName = self.openPDFFileNameDialog()
        self.disableObject(self.ui.pdf_Powerpoint_Button)
        self.Thread = Thread(target=self.pdf_PowerPoint, daemon=True)
        self.Thread.start()
    def unlockPdf(self):
        self.pdffileName = self.openPDFFileNameDialog()
        if pypdf.PdfReader(self.pdffileName).is_encrypted:
            self.infoUpdate(self.pdffileName +" is password protected")
            self.placeholder()
            self.stPassword, self.pwdstatus = QtWidgets.QInputDialog.getText(
                                    self, 'Password Input Dialog', 'Enter the password to unlock the pdf:') 
            self.disableObject(self.ui.pdf_Unlock_Button)
            self.Thread = Thread(target=self.pdf_Unlock, daemon=True)
            self.Thread.start()
        else:
            self.infoUpdate(self.pdffileName +" is not password protected. No need to unlock")
            self.placeholder()
    def mergePdf(self):
        self.pdffileNames = self.openPDFFileNamesDialog() 
        self.disableObject(self.ui.pdf_Merge_Button)
        self.Thread = Thread(target=self.merge_Pdf, daemon=True)
        self.Thread.start()
    def sleepTime(self, iTime):
        QtTest.QTest.qWait(iTime)
    def handleLinks(self,url):
        if not url.scheme():
            url = QUrl.fromLocalFile(url.toString())
        QDesktopServices.openUrl(url)
    def updateHyperlink(self, stLinkPath, stLinkText):
        stfileLink = Path(stLinkPath).as_uri()
        stfileLink = "<a href="+stfileLink+">"+stLinkText+"</a>"
        self.infoUpdate(stfileLink)
    def openPDFFileNameDialog(self):
        self.pdffileName = ""
        self.pdffileName, self.extension = QFileDialog.getOpenFileName(self,"Choose PDF File", "","All Files (*);;PDF Files (*.pdf)")
        if self.pdffileName:
            if self.pdffileName.endswith(".pdf"):
                self.infoUpdate("Choosen PDF File is: "+self.pdffileName)
                self.placeholder()
                return self.pdffileName
            else:
                self.errorUpdate("choose a pdf file. You have Choosen a "+self.pdffileName)
                self.placeholder()
                return self.pdffileName
        else:
            self.errorUpdate("Choose PDF File. Without choosing PDF file, the operation cannot be performed")
            self.placeholder()
            return self.pdffileName
    def openPDFFileNamesDialog(self):
        self.pdffileNames = []
        self.pdffileNames, self.extensions = QFileDialog.getOpenFileNames(self,"QFileDialog.getOpenFileNames()", "","All Files (*);;PDF Files (*.pdf)")
        print(str(self.extensions))
        if self.pdffileNames:
            pdfresults = [ststring for ststring in self.pdffileNames if ststring.endswith(".pdf")]
            if pdfresults == self.pdffileNames:
                self.infoUpdate("Choosen PDF File is: "+str(self.pdffileNames))
                self.placeholder()
                return self.pdffileNames
            else:
                self.errorUpdate("choose a pdf file. You have Choosen a "+str(self.pdffileNames))
                self.placeholder()
                return self.pdffileNames
        else:
            self.errorUpdate("Choose PDF File. Without choosing PDF file, the operation cannot be performed")
            self.placeholder()
            return self.pdffileNames
    def pdf_Doc(self):
        try:
            if self.pdffileName.endswith('.pdf'):
                self.docx_file = self.pdffileName.replace('.pdf', '.docx')
            else:
                raise Exception("Please choose a pdf file")
            self.infoUpdate("Converting PDF to DOC")
            self.placeholder()
            # convert pdf to docx
            cv = Converter(self.pdffileName)
            cv.convert(self.docx_file, start=0, end=None)      # all pages by default
            cv.close()
            if os.stat(self.docx_file).st_size > 0:
                self.validUpdate(self.pdffileName+" is converted to "+
                                 self.docx_file)
                self.placeholder()
                self.updateHyperlink(self.docx_file, "Click to View the converted Docx File")
                self.placeholder()
            else:
                self.errorUpdate(self.pdffileName+" is not converted. Please retry")
                self.placeholder()
        except Exception as e:
            self.errorUpdate(str(e))
            self.placeholder()
        finally:
            self.enableObject(self.ui.pdf_Doc_Button)
    def pdf_PowerPoint(self):
        try:
            if self.pdffileName.endswith('.pdf'):
                self.powerpoint_file = self.pdffileName.replace('.pdf', '.pptx')
            else:
                raise Exception("Please choose a pdf file")
            # convert pdf to ppt
            self.infoUpdate("Converting PDF to PPT")
            self.placeholder()
            proc = subprocess.Popen('pdf2pptx '+self.pdffileName+' -o '+self.powerpoint_file, stdin = subprocess.PIPE, stdout = subprocess.PIPE)
            stdout, stderr = proc.communicate()
            print(stdout)
            if os.stat(self.powerpoint_file).st_size > 0:
                self.validUpdate(self.pdffileName+" is converted to "+
                                 self.powerpoint_file)
                self.placeholder()
                self.updateHyperlink(self.powerpoint_file, "Click to View the converted PPT File")
                self.placeholder()
            else:
                self.errorUpdate(self.pdffileName+" is not converted. Please retry")
                self.placeholder()
        except Exception as e:
            self.errorUpdate(str(e))
            self.placeholder()
        finally:
            self.enableObject(self.ui.pdf_Powerpoint_Button)
    
    def pdf_Unlock(self):
        try:
            if self.pdffileName.endswith('.pdf'):
                self.unlock_file = self.pdffileName.replace('.pdf', '_unlock.pdf')
            else:
                raise Exception("Please choose a pdf file")
            # unlock pdf
            self.infoUpdate("Unlocking PDF")
            self.placeholder()
            if self.pwdstatus:
                pdf = pikepdf.open(self.pdffileName, password=self.stPassword)
                pdf.save(self.unlock_file)
            else:
                self.errorUpdate("Please enter the correct password")
                self.placeholder()
            if os.stat(self.unlock_file).st_size > 0:
                self.validUpdate(self.pdffileName+" is unlocked to "+
                                 self.unlock_file)
                self.placeholder()
                self.updateHyperlink(self.unlock_file, "Click to View the Unlocked File")
                self.placeholder()
            else:
                self.errorUpdate(self.pdffileName+" is not unlocked. Please retry")
                self.placeholder()
        except Exception as e:
            self.errorUpdate(str(e))
            self.placeholder()
        finally:
            self.enableObject(self.ui.pdf_Unlock_Button)
    def merge_Pdf(self):
        try:
            pdfresults = [ststring for ststring in self.pdffileNames if ststring.endswith(".pdf")]
            if pdfresults == self.pdffileNames:
                self.merge_file = self.pdffileNames[0].replace('.pdf', '_merged.pdf')
            else:
                raise Exception("Please choose all files as PDF files")
            self.infoUpdate("Merging PDF")
            self.placeholder()
            pdfDoc = Pdf.new()
            for file in self.pdffileNames:  # you can change this to browse directories recursively
                with Pdf.open(file) as src:
                    pdfDoc.pages.extend(src.pages)
            pdfDoc.save(self.merge_file)
            pdfDoc.close()
            if os.stat(self.merge_file).st_size > 0:
                self.validUpdate(str(self.pdffileNames)+" is merged to "+
                                 self.merge_file)
                self.placeholder()
                self.updateHyperlink(self.merge_file, "Click to View the Merged File")
                self.placeholder()
            else:
                self.errorUpdate(self.pdffileNames+" is not merged. Please retry")
                self.placeholder()
        except Exception as e:
            self.errorUpdate(str(e))
            self.placeholder()
        finally:
            self.enableObject(self.ui.pdf_Merge_Button)
    def disableObject(self, stobject):
        get_updater().call_in_main(stobject.setEnabled, False)
    def enableObject(self, stobject):
        get_updater().call_in_main(stobject.setEnabled, True)
    def validUpdate(self, stText):
        get_updater().call_in_main(self.ui.textBrowser.append, validFormat.format(stText))
    def errorUpdate(self, stText):
        get_updater().call_in_main(self.ui.textBrowser.append, errorFormat.format(stText))
    def infoUpdate(self, stText):
        get_updater().call_in_main(self.ui.textBrowser.append, infoFormat.format(stText))
    def placeholder(self):
        get_updater().call_in_main(self.ui.textBrowser.append, infoFormat.format("==============================================="))
def getApplicationPath():
    if getattr(sys, 'frozen', False):
        applicationpath = os.path.dirname(sys.executable)
    elif __file__:
        applicationpath = os.path.dirname(__file__)
    return applicationpath

#================================================================================================================================
#Format of output strings
errorFormat = '<span style="color:red;">{}</span>'
warningFormat = '<span style="color:orange;">{}</span>'
infoFormat = '<span style="color:blue;">{}</span>'
validFormat = '<span style="color:green;">{}</span>'

def main():
    global stCwd
    stCwd = getApplicationPath()
    app = QtWidgets.QApplication(sys.argv)
    pdf_HandlerMain = pdf_Handler()
    pdf_HandlerMain.show()
    app.exec()

if __name__ == "__main__":
    print("Hello, World!")
    main()