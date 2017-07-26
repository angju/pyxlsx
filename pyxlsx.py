# -*- coding: utf-8 -*-

"""
Module implementing main.
"""
import sys
from PyQt5.QtCore import pyqtSlot,  QThread, pyqtSignal
from PyQt5.QtWidgets import QDialog,  QApplication, QFileDialog, QMessageBox
import xlwings as xw
from pathlib import Path

from Ui_pyxlsx import Ui_Dialog

class WorkThread(QThread):
    trigger = pyqtSignal("QString")
    info = pyqtSignal(int,  "QString")
    def __init__(self):
        super(WorkThread, self).__init__()
    def setData(self,  files):
        self.rowKey = 0
        self.colName = 0
        self.colItem1 = 0
        self.colItem2 = 0
        self.files = files
    def EvaluateMostRight(self,  cell):
        savedCell = cell
        col = cell.column
        cell = cell.end('right')
        while (cell.value):
            col = cell.column
            cell = cell.end('right')
            if (col == 1 and savedCell.value == None):
                col = 0
        return col
    def EvaluateMostDown(self,  cell):
        savedCell = cell
        row = cell.row
        cell = cell.end('down')
        while (cell.value):
            row = cell.row
            cell = cell.end('down')
            if (row == 1 and savedCell.value == None):
                row = 0
        return row
    def EvaluateColBoundary(self,  sheet):
        mostRight = 0
        for row in range(1,  10):
            cell = sheet.range(row,  1)
            mostRight = max(self.EvaluateMostRight(cell),  mostRight)
        return mostRight
    def FindKeyLocation(self,  sheet):
        boundaryCol = self.EvaluateColBoundary(sheet)
        row = 1
        colName = 0
        colItem1 = 0
        colItem2 = 0
        while True:
            itemCount = 1
            for col in range(1,  boundaryCol+1):
                value = sheet.range(row,  col).value
                if (value == None):
                    continue
                value = ''.join(str(value).split())
                if (value == '姓名'):
                    colName = col
                if (value == '项目名称'):
                    if itemCount == 1:
                        colItem1 = col
                    if itemCount == 2:
                        colItem2 = col
                    itemCount = itemCount + 1
            if colName > 0 or row > 20:
                break
            else:
                row = row + 1
        self.rowKey = row
        self.colName = colName
        self.colItem1 = colItem1
        self.colItem2 = colItem2
        
    def run(self):
        class Data:
            def __init__(self):
                self.gradeClass = None
                self.teacher = None
                self.book = None
        
        try:
            app = xw.App(visible=False,  add_book=False)
            datas = []
            outDatas1 = []
            outDatas2 = []
            
            outFileName = "pyxlsx.output.xlsx"
            outPath = Path(outFileName)
            if outPath.exists():
                if outPath.is_dir():
                    outPath.rmdir()
                elif outPath.is_file():
                    outPath.unlink()    
                    
            self.trigger.emit("收集有效表格。。。")
            for filename in self.files:
                file = Path(filename)
                gradePos = file.stem.find("年")
                classPos = file.stem.find("班")
                if gradePos != -1 and classPos != -1:
                    gradeStr = file.stem[0:gradePos]
                    classStr = file.stem[gradePos+1:classPos]
                    teacher = file.stem[classPos+1:]
                    self.trigger.emit(gradeStr + "年级" + classStr + "班" + teacher)
                    data = Data()
                    data.gradeClass = gradeStr + classStr
                    data.teacher = teacher
                    data.book = app.books.open(filename)
                    datas.append(data)
                    
            self.trigger.emit("读取表格数据。。。")
            for data in datas:
                self.trigger.emit("处理文件" + str(data.book.name))
                self.FindKeyLocation(data.book.sheets[0])
                if (self.colName == 0):
                    self.trigger.emit("表格缺少'姓名', 跳过此文件")
                    continue
                if (self.colItem1 == 0):
                    self.trigger.emit("表格缺少'第一段（项目）', 跳过此文件")
                    continue
                if (self.colItem2 == 0):
                    self.trigger.emit("表格缺少'第二段（项目）', 跳过此文件")
                    continue
                
                nameKey = data.book.sheets[0].range(self.rowKey,  self.colName)
                nameStartRow = nameKey.row + 1
                #nameEndRow = nameKey.end("down").row
                nameEndRow = self.EvaluateMostDown(nameKey)
                item1EndRow = nameEndRow
                item2EndRow = nameEndRow
                names = data.book.sheets[0].range((self.rowKey,  self.colName),  (nameEndRow,  self.colName)).value
                items1 = data.book.sheets[0].range((self.rowKey,  self.colItem1),  (item1EndRow,  self.colItem1)).value
                items2 = data.book.sheets[0].range((self.rowKey,  self.colItem2),  (item2EndRow,  self.colItem2)).value
                if nameStartRow == nameEndRow:
                    items1 = [(names,  items1)]
                    items2 = [(names,  items2)]
                else:
                    items1 = [*zip(names,  items1)]
                    items2 = [*zip(names,  items2)]
                outDatas1 += [(item[0],  data.gradeClass,  data.teacher,  item[1]) for item in items1 if (item[1] is not None) and (str(item[1]).strip() != '')]
                outDatas2 += [(item[0],  data.gradeClass,  data.teacher,  item[1]) for item in items2 if (item[1] is not None) and (str(item[1]).strip() != '')]
            
            self.trigger.emit("数据排序。。。")
            outDatas1 = sorted(outDatas1,  key=lambda outdata: (outdata[3],  outdata[1]))
            outDatas2 = sorted(outDatas2,  key=lambda outdata: (outdata[3],  outdata[1]))
            
            self.trigger.emit("生成数据。。。")
            outBook = app.books.add()
            if (len(outBook.sheets) < 2):
                outBook.sheets.add()
            outBook.sheets[0].name = "第一段"
            outBook.sheets[1].name = "第二段"
            outBook.sheets[0].range("A1").value = ["姓名",  "班级",  "班主任",  "第一段（项目）"]
            outBook.sheets[1].range("A1").value = ["姓名",  "班级",  "班主任",  "第二段（项目）"]
            outBook.sheets[0].range("A2").value = outDatas1
            outBook.sheets[1].range("A2").value = outDatas2
            
            self.trigger.emit("保存数据。。。")
            outBook.save(outFileName)
            
            self.trigger.emit("输出文件：" + outFileName)
        
        except Exception as err:
            self.info.emit(1, str(err))
        
        else:
            self.info.emit(0, "运行成功！")
        
        finally:
            for book in app.books:
                book.close()
            app.quit()

class main(QDialog, Ui_Dialog):
    """
    Class documentation goes here.
    """
    def __init__(self, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget
        @type QWidget
        """
        super(main, self).__init__(parent)
        self.setupUi(self)
        self.pushButtonRun.setEnabled(False)
        self.files = []
        self.work = WorkThread()
        self.work.trigger.connect(self.textBrowserOutputAppend)
        self.work.info.connect(self.outputRunInfo)
        
    def textBrowserOutputAppend(self,  info):
        self.textBrowserOutput.append(info)
    def outputRunInfo(self,  infoNo,  infoStr): 
        if (infoNo == 0):
            QMessageBox.information(self,  "info",  infoStr,  QMessageBox.Yes)
        else:
            QMessageBox.critical(self,  "错误",  infoStr,  QMessageBox.Yes)
        self.enableAllButtons()
    def disableAllButtons(self):
        self.pushButtonSelectFile.setEnabled(False)
        self.pushButtonClose.setEnabled(False)
        self.pushButtonRun.setEnabled(False)
    def enableAllButtons(self):
        self.pushButtonSelectFile.setEnabled(True)
        self.pushButtonClose.setEnabled(True)
        self.pushButtonRun.setEnabled(True)
    
    @pyqtSlot()
    def on_pushButtonSelectFile_clicked(self):
        """
        Slot documentation goes here.
        """
        cwdPath = Path().cwd()
        filenames, type = QFileDialog.getOpenFileNames(self, "选择Excel文件", str(cwdPath), "*.xls;*.xlsx")
        for filename in filenames:
            if filename in self.files:
                continue
            self.textBrowser.append(filename)
            self.files.append(filename)
        if len(self.files) > 0:
            self.pushButtonRun.setEnabled(True)
        else:
            self.pushButtonRun.setEnabled(False)
    
    @pyqtSlot()
    def on_pushButtonClear_clicked(self):
        """
        Slot documentation goes here.
        """
        self.files = []
        self.textBrowser.setText('')
        self.pushButtonRun.setEnabled(False)
            
    
    @pyqtSlot()
    def on_pushButtonRun_clicked(self):
        """
        Slot documentation goes here.
        """
        self.disableAllButtons()
        self.work.setData(self.files)
        self.work.start()

    

    

        
if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainDialog = main()
    mainDialog.show()
    sys.exit(app.exec_())
