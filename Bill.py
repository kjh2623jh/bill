import sys
import os
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QCoreApplication
from PyQt5.QtGui import QIcon, QFont
import win32com.client
from PIL import ImageGrab
import json


class BillApp(QWidget):
    def __init__(self):
        super().__init__()
        self.message_log = ''
        try:
            with open('Path.json') as f:
                self.save_path = json.load(f)
        except:
            self.save_path = {"path":''}

        self.initUI()

    def initUI(self):
        QToolTip.setFont(QFont('SansSerif', 10))


        self.table = QTextBrowser()
        getTextButton = QPushButton('내역 넣기', self)
        getTextButton.clicked.connect(self.getMessageLog)

        self.line = QLineEdit(self)
        self.line.setText(self.save_path["path"])
        folderButton = QPushButton("찾아보기..")
        folderButton.clicked.connect(self.getPath)

        okButton = QPushButton('완료')
        okButton.clicked.connect(self.accept)
        cancelButton = QPushButton('취소')
        cancelButton.clicked.connect(QCoreApplication.instance().quit)


        getButton_hbox = QHBoxLayout()
        getButton_hbox.addStretch(1)
        getButton_hbox.addWidget(getTextButton)
        getButton_hbox.addStretch(1)

        path_hbox = QHBoxLayout()
        path_hbox.addStretch(1)
        path_hbox.addWidget(self.line)
        path_hbox.addWidget(folderButton)
        path_hbox.addStretch(1)

        button_hbox = QHBoxLayout()
        button_hbox.addStretch(1)
        button_hbox.addWidget(okButton)
        button_hbox.addWidget(cancelButton)
        button_hbox.addStretch(1)

        vbox = QVBoxLayout()
        vbox.addWidget(self.table)
        vbox.addLayout(getButton_hbox)
        vbox.addStretch(1)
        vbox.addLayout(path_hbox)
        vbox.addStretch(1)
        vbox.addLayout(button_hbox)
        vbox.addStretch(1)

        self.setLayout(vbox)


        self.setWindowTitle('금액 청구')
        self.setWindowIcon(QIcon('ico/bill.ico'))
        self.resize(500, 350)
        self.center()

        self.show()

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
    
    def getMessageLog(self):
        text, ok = QInputDialog.getMultiLineText(self, '내역', '메시지 붙여넣기:')

        if ok:
            self.message_log=str(text)
            self.table.setText(self.message_log)
    
    def getPath(self):
        path=QFileDialog.getExistingDirectory(self, "Select Directory")
        if path: self.line.setText(path)

    def accept(self):
        save_path = {"path": self.line.text()}
        with open("Path.json", 'w') as f:
            json.dump(save_path, f, indent=4)

        log = self.message_log.split('[Web발신]\n')
        log_list=[]
        for idx in range(1,len(log)):
            txt = log[idx].split()
            if txt[0] == '농협':
                if txt[1][:2]=="입금": continue
                log_list.append(f"{txt[2]} | {txt[5]} | {txt[1][2:-1]}")
            else:
                txt2 = txt[3].split('(')
                log_list.append(f"{txt[0][2:]} | {txt2[1][:-1]} | {txt2[0][4:-1]}")
        
        self.ca = checkApp(self, log_list)
        self.ca.exec()

        if self.ca.check_idx:
            BillApp.makeFile(log_list, self.ca.check_idx, self.line.text())
            self.close()
    

    def makeFile(log, arr, path):
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible=True
        workbook = excel.Workbooks.Add()
        sheet = workbook.Sheets(1)
        sheet.Name = "Bill"

        sheet.Range("A1").Value = "날짜"
        sheet.Range("C1").Value = "가격"
        for idx in range(len(arr)):
            txt = log[arr[idx]].split(" | ")
            sheet.Cells(idx+2,1).Value = txt[0]
            sheet.Cells(idx+2,2).Value = txt[1]
            sheet.Range(f"C{idx+2}").Value = int(txt[2].replace(',',''))
        sheet.Cells(idx+3,1).Value = "합계"
        sheet.Cells(idx+3,2).Value = "-"
        sheet.Cells(idx+3,3).Value = f"=SUM(C2:C{idx+2})"

        sheet.Columns(2).ColumnWidth = 20
        sheet.Range("A1:C1").Interior.ColorIndex = 24
        sheet.Range("A{0}:C{0}".format(idx+3)).Interior.ColorIndex = 15
        sheet.Range(f"A1:B{idx+3}").HorizontalAlignment  = 3
        sheet.Range("C1").HorizontalAlignment  = 3        
        all = sheet.Range(f"A1:C{idx+3}")
        all.Borders.ColorIndex = 1
        all.Borders.Weight = 2
        all.Borders.LineStyle = 1

        try:
            if not os.path.exists(path+"/build2/pdf"):
                os.makedirs(path+"/build2/pdf")
            if not os.path.exists(path+"/build2/png"):
                os.makedirs(path+"/build2/png")
        except OSError:
            print("Error: Creating diractory. "+path+"/build2/...")
        else:
            year= 2022
            month= txt[0].split('/')[0]
            path+=f"/build2"

            sheet.Range(f"A1:C{idx+3}").Copy()
            img = ImageGrab.grabclipboard()
            img.save(path+f"/png/{year}-{month}.png")
            sheet.Cells(1,1).Copy()
            workbook.ActiveSheet.ExportAsFixedFormat(0, path+f"/pdf/{year}-{month}")
            workbook.Close(False)
            excel.Quit()



class checkApp(QDialog):
    def __init__(self, parent, loglist):
        super(checkApp, self).__init__(parent)
        self.len_ = len(loglist)
        self.check_list = [None]*self.len_
        self.check_idx = []
        self.log_list = loglist

        self.initUI()

    def initUI(self):
        for i in range(self.len_):
            self.check_list[i] = QCheckBox(self.log_list[i], self)
            self.check_list[i].toggle()
            self.check_list[i].move(20, 20*i+1)

        okButton = QPushButton('완료')
        okButton.clicked.connect(self.accept)

        cancelButton = QPushButton('취소')
        cancelButton.clicked.connect(self.back)

        hbox = QHBoxLayout()
        hbox.addStretch(1)
        hbox.addWidget(okButton)
        hbox.addWidget(cancelButton)
        hbox.addStretch(1)

        vbox = QVBoxLayout()
        for widget in self.check_list:
            vbox.addWidget(widget)
        vbox.addStretch(1)
        vbox.addLayout(hbox)

        self.setLayout(vbox)

        self.setWindowTitle('확인')
        self.setWindowIcon(QIcon('ico/checkbox.ico'))
        self.resize(400, 20*i+40)
        self.center()

        self.show()
    
    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
    
    def accept(self):
        self.check_idx = [i for i in range(self.len_) if self.check_list[i].isChecked()]
        self.close()

    def back(self):
        self.close()


if __name__ == '__main__':
   app = QApplication(sys.argv)
   program = BillApp()
   sys.exit(app.exec_())