import sys
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QPushButton, QLabel, QLineEdit, QComboBox, QGridLayout, QFileDialog
from PyQt5.Qt import QThread, QMutex
from PyQt5.QtCore import pyqtSignal, pyqtSlot
from docx import Document
import copy
import xlrd
import time


class WordU(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super(WordU, self).__init__(parent)
        self.setFixedSize(400, 200)
        self.setWindowTitle("自动填写")
        self.file_path = None
        self.save_word = None
        self.file_in = None
        self.save_in = None
        self.style_in = None
        self.yes_b = None
        
        # 布局初始化
        self.glayout = QGridLayout()
        self.glayout.setSpacing(10)
        self.setLayout(self.glayout)
        # 函数初始化
        self.set_prom()
        self.activity()
    
    def activity(self):
        self.file_in.clicked.connect(self.choose_x_file)
        self.save_in.clicked.connect(self.choose_w_file)
        self.yes_b.clicked.connect(self.start_W)
    
    def set_prom(self):
        self.file_path = QLineEdit(self)
        self.file_in = QPushButton("选择Excel", self)
        self.file_path.setReadOnly(True)
        self.glayout.addWidget(self.file_path, 1, 1, 1, 10)
        self.glayout.addWidget(self.file_in, 1, 11, 1, 4)
        
        self.save_word = QLineEdit(self)
        self.save_in = QPushButton("选择Word", self)
        self.save_word.setReadOnly(True)
        self.glayout.addWidget(self.save_word, 2, 1, 1, 10)
        self.glayout.addWidget(self.save_in, 2, 11, 1, 4)
        
        self.style_in = QComboBox(self)
        self.style_in.addItems(["测试说明", "测试记录"])
        self.glayout.addWidget(self.style_in, 3, 4, 1, 2)
        self.yes_b = QPushButton("开始填写", self)
        self.glayout.addWidget(self.yes_b, 3, 8, 1, 4)
        self.prompt = QLabel(self)
        self.glayout.addWidget(self.prompt, 3, 1, 1, 2)
    
    def choose_x_file(self):
        filename, i = QFileDialog.getOpenFileNames(None, "请选择要添加的文件", "./",
                                                   "Text Files (*.xlsx);;Text Files (*.xls);;All Files (*)")
        if filename:
            self.file_path.setText(filename[0])
    
    def choose_w_file(self):
        filename, i = QFileDialog.getOpenFileNames(None, "请选择要添加的文件", "./",
                                                   "Word Files (*.docx);;Word Files (*.doc);;All Files (*)")
        if filename:
            self.save_word.setText(filename[0])
    
    def start_W(self):
        self.prompt.setText("执行中...")
        in_item = self.style_in.currentIndex()
        # if in_item == 1:
        #     in_item = 2
        self.write_word = ThreadRW(self.file_path.text(), self.save_word.text(), in_item)
        self.write_word.str_out.connect(self.prompt_out)
        self.write_word.start()
        # time.sleep(1)
    
    @pyqtSlot(str)
    def prompt_out(self, i):
        self.prompt.setText(i)


lock = QMutex()


class ThreadRW(QThread):
    str_out = pyqtSignal(str)
    
    def __init__(self, xfile=None, wfile=None, item=None, parent=None):
        super().__init__(parent)
        self.R_file = xfile
        self.W_file = wfile
        self.Item = item
    
    def move_table_after(self, table, paragraph):
        tbl, p = table._tbl, paragraph._p
        p.addnext(tbl)
    
    def run(self):
        lock.lock()
        myWorkbook = xlrd.open_workbook(self.R_file)
        mySheets = myWorkbook.sheets()
        mySheet = mySheets[0]
        rows = mySheet.nrows
        cols = mySheet.ncols
        temp = [[] * 2 for row in range(rows)]
        for row in range(rows):
            for col in range(cols):
                row_data = mySheet.cell_value(row, col)
                temp[row].append(row_data)
        
        doc = Document(self.W_file)
        tb = doc.tables
        num = 0
        i = 3
        if self.Item == 1:
            self.number = 2
        while True:
            a = tb[i].cell(0, 0).text
            b = tb[i].cell(0, 4 + self.Item).text
            if "用例名称" in a and "用例标识" in b:
                copy_tb = copy.deepcopy(tb[i])
                tb[i].cell(0, 1).text = temp[num][0]
                run = tb[i].cell(0, 5 + self.number).paragraphs[0].add_run(temp[num][1])
                run.font.name = 'Times New Roman'
                try:
                    num += 1
                    if temp[num]:
                        pg = doc.paragraphs
                        self.move_table_after(copy_tb, pg[len(pg) - 1])
                        doc.add_paragraph(" ")
                        doc.add_paragraph(" ")
                        tb = doc.tables
                except BaseException as e:
                    print(e)
                    break
            i += 1
        doc.save(self.W_file)
        i = "执行完成"
        self.str_out.emit(str(i))
        lock.unlock()


if __name__ == '__main__':
    App = QApplication(sys.argv)
    ex = WordU()
    ex.show()
    sys.exit(App.exec_())
