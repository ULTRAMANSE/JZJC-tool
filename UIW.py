import sys
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QPushButton, QLabel, QLineEdit, QComboBox, QGridLayout, QFileDialog
from PyQt5.Qt import QThread, QMutex
from PyQt5.QtCore import pyqtSignal, pyqtSlot
from  PyQt5.QtGui import QIcon
from docx import Document
import copy


class WordU(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super(WordU, self).__init__(parent)
        self.setFixedSize(400, 150)
        self.setWindowTitle("自动填写")
        self.setWindowIcon(QIcon("i.ico"))
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
        self.save_in.clicked.connect(self.choose_w_file)
        self.yes_b.clicked.connect(self.start_W)
    
    def set_prom(self):
        
        self.save_word = QLineEdit(self)
        self.save_in = QPushButton("选择Word", self)
        self.save_word.setReadOnly(True)
        self.glayout.addWidget(self.save_word, 1, 1, 1, 10)
        self.glayout.addWidget(self.save_in, 1, 11, 1, 4)
        
        self.style_in = QComboBox(self)
        self.style_in.addItems(["测试说明", "测试记录"])
        self.glayout.addWidget(self.style_in, 2, 4, 1, 2)
        self.yes_b = QPushButton("开始填写", self)
        self.glayout.addWidget(self.yes_b, 2, 8, 1, 4)
        self.prompt = QLabel(self)
        self.glayout.addWidget(self.prompt, 2, 1, 1, 2)

    # def choose_x_file(self):
        # filename, i = QFileDialog.getOpenFileNames(None, "请选择要添加的文件", "./",
        #                                            "Text Files (*.xlsx);;Text Files (*.xls);;All Files (*)")
        # if filename:
        #     self.file_path.setText(filename[0])
    
    def choose_w_file(self):
        filename, i = QFileDialog.getOpenFileNames(None, "请选择要添加的文件", "./",
                                                   "Word Files (*.docx);;Word Files (*.doc);;All Files (*)")
        if filename:
            self.save_word.setText(filename[0])
    
    def start_W(self):
        self.prompt.setText("执行中...")
        in_item = self.style_in.currentIndex()
        self.write_word = ThreadRW(self.save_word.text(), in_item)
        self.write_word.str_out.connect(self.prompt_out)
        self.write_word.start()
    
    @pyqtSlot(str)
    def prompt_out(self, i):
        self.prompt.setText(i)


lock = QMutex()


class ThreadRW(QThread):
    str_out = pyqtSignal(str)
    
    def __init__(self, wfile=None, item=None, parent=None):
        super().__init__(parent)
        self.W_file = wfile # word文件名
        self.Item = item # 类型选项
        self.number = 0 # 记录差
        self.num = 0 # 列表下表值
        self.t_number = 5 # 表格index
        self.t_number_copy = 0  # 复制表格index
        self.t_rows = 1 # 表格的行index
    
    def move_table_after(self, table, paragraph):
        tbl, p = table._tbl, paragraph._p
        p.addnext(tbl)
    
    def run(self):
        lock.lock()
        # myWorkbook = xlrd.open_workbook(self.R_file)
        # mySheets = myWorkbook.sheets()
        # mySheet = mySheets[0]
        # rows = mySheet.nrows
        # cols = mySheet.ncols
        # temp = [[] * 2 for row in range(rows)]
        # for row in range(rows):
        #     for col in range(cols):
        #         row_data = mySheet.cell_value(row, col)
        #         temp[row].append(row_data)
        #

        if self.Item == 1:
            self.number = 2
        example_name = []
        identity = []
        doc = Document(self.W_file)
        tb = doc.tables

        while True:
            a = tb[self.t_number].cell(0, 2).text
            b = tb[self.t_number].cell(0, 4).text
            if "用例标识" in a and "用例名称" in b:
                self.t_number_copy = self.t_number
                break
            self.t_number += 1

        while True:
            if self.t_number_copy == 0:
                break
            try:
                c = tb[self.t_number_copy].cell(self.t_rows, 2).text
                d = tb[self.t_number_copy].cell(self.t_rows, 4).text
            except BaseException as e:
                print(e)
                break
            example_name.append(d)
            identity.append(c)
            self.t_rows += 1

        while True:
            a = tb[self.t_number].cell(0, 0).text
            b = tb[self.t_number].cell(0, 4 + self.Item).text
            if "用例名称" in a and "用例标识" in b:
                self.t_number_copy = self.t_number
                break
            self.t_number += 1

        while True:
            if self.t_number_copy == 0:
                break
            copy_tb = copy.deepcopy(tb[self.t_number_copy])
            tb[self.t_number_copy].cell(0, 1).text = example_name[self.num]
            run = tb[self.t_number_copy].cell(0, 5 + self.number).paragraphs[0].add_run(identity[self.num])
            run.font.name = 'Times New Roman'
            self.t_number_copy += 1
            try:
                self.num += 1
                if example_name[self.num]:
                    pg = doc.paragraphs
                    self.move_table_after(copy_tb, pg[len(pg) - 1])
                    doc.add_paragraph(" ")
                    doc.add_paragraph(" ")
                    tb = doc.tables
            except BaseException as e:
                print(e)
                break

        doc.save(self.W_file)
        re_str = "执行完成"
        self.str_out.emit(str(re_str))
        lock.unlock()


if __name__ == '__main__':
    App = QApplication(sys.argv)
    ex = WordU()
    ex.show()
    sys.exit(App.exec_())
