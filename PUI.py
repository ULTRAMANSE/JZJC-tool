import sys
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QPushButton, QLineEdit, QComboBox, QGridLayout, QFileDialog
from PyQt5.QtGui import QIcon
from pypinyin import Style, lazy_pinyin
import xlwt, xlrd
import resource


class PinYin(QtWidgets.QWidget):
	def __init__(self, parent=None):
		super(PinYin, self).__init__(parent)
		self.setFixedSize(400, 250)
		self.setWindowTitle("标识转换")
		self.setWindowIcon(QIcon(":/favicon.ico"))
		self.file_path = None  # 要转化的文件路径
		self.save_path = None  # 转化后的文件保存路径
		self.file_in = None  # 路径选择按钮
		self.save_in = None
		self.style_in = None  # 标识类型
		self.style2_in = None  # 标识2类型
		self.yes_b = None  # 确定按钮

		# 布局初始化
		self.glayout = QGridLayout()
		self.glayout.setSpacing(10)
		self.setLayout(self.glayout)
		# 函数初始化
		self.set_prom()
		self.activity()

	def activity(self):
		self.file_in.clicked.connect(self.choose_file)
		self.save_in.clicked.connect(self.choose_path)
		self.yes_b.clicked.connect(self.transfor)

	def set_prom(self):
		self.file_path = QLineEdit(self)
		self.file_in = QPushButton("选择文件", self)
		self.file_path.setReadOnly(True)
		self.glayout.addWidget(self.file_path, 1, 1, 1, 10)
		self.glayout.addWidget(self.file_in, 1, 11, 1, 4)

		self.save_path = QLineEdit(self)
		self.save_in = QPushButton("保存路径", self)
		self.save_path.setReadOnly(True)
		self.glayout.addWidget(self.save_path, 2, 1, 1, 10)
		self.glayout.addWidget(self.save_in, 2, 11, 1, 4)

		self.style_in = QComboBox(self)
		self.style_in.addItems(["功能性 F", "性能 P", "安全性 IS", "兼容性 Sc", "易用性 Su", "可靠性 Sr", "可维护 Sm", "可移植 Sp"])
		self.glayout.addWidget(self.style_in, 3, 1, 1, 3)
		self.style2_in = QComboBox(self)
		self.style2_in.addItems(["测试项", "测试用例"])
		self.glayout.addWidget(self.style2_in, 3, 4, 1, 2)
		self.yes_b = QPushButton("开始转化", self)
		self.glayout.addWidget(self.yes_b, 3, 8, 1, 4)

	def choose_file(self):
		filename, i = QFileDialog.getOpenFileNames(None, "请选择要添加的文件", "./",
												   "Text Files (*.xlsx);;Text Files (*.xls);;All Files (*)")
		self.file_path.setText(filename[0])

	def choose_path(self):
		pathname = QFileDialog.getExistingDirectory(None, "请选择文件夹路径", "./")
		self.save_path.setText(pathname + "/print.xls")

	def choose_imge_path(self):
		pathname = QFileDialog.getExistingDirectory(None, "请选择文件夹路径", "./")
		self.imge_path.setText(pathname)

	def cell_real_value(self, row, col):
		for merged in self.mySheet.merged_cells:  # 判断合并的单元格
			if (row >= merged[0] and row < merged[1]
					and col >= merged[2] and col < merged[3]):
				return self.mySheet.cell_value(merged[0], merged[2])
		return self.mySheet.cell_value(row, col)

	def transfor(self):
		if self.file_path.text() is " " or self.save_path.text() is " ":
			pass
		else:
			in_style = self.style_in.currentText().split(" ")[1]  # 获取转换类型
			in_item = self.style2_in.currentIndex()
			self.myWorkbook = xlrd.open_workbook(self.file_path.text())
			self.mySheets = self.myWorkbook.sheets()
			self.mySheet = self.mySheets[0]
			temp = None
			rows = self.mySheet.nrows
			cols = self.mySheet.ncols
			if in_item == 0:
				temp = []
				for row in range(rows):
					row_data = self.mySheet.row_values(row)
					if row_data[0]:
						temp.append("".join(lazy_pinyin(row_data[0], style=Style.FIRST_LETTER)).upper())  # 转化添加
			if in_item == 1:
				temp = [[] * 2 for row in range(rows)]
				for row in range(rows):
					for col in range(cols):
						row_data = self.cell_real_value(row, col)
						if row_data is not "":
							temp[row].append("".join(lazy_pinyin(row_data, style=Style.FIRST_LETTER)).upper())

			workbook = xlwt.Workbook(encoding='utf-8')
			worksheet = workbook.add_sheet('pinyin')
			style = xlwt.XFStyle()  # 初始化样式
			font = xlwt.Font()  # 为样式创建字体
			font.name = 'Times New Roman'
			style.font = font  # 设定样式

			for i in range(len(temp)):
				print(1)
				if in_item == 0:
					print(2)
					worksheet.write(i, 0, temp[i], style)
				else:
					print(3)
					if i < 9:
						tstr = "00" + str(i + 1)
					elif i < 99:
						tstr = "0" + str(i + 1)
					else:
						tstr = str(i + 1)
					worksheet.write(i, 0, temp[i][0] + "-" + in_style + "-" + temp[i][1] + "-" + tstr, style)

			workbook.save(self.save_path.text())


if __name__ == '__main__':
	App = QApplication(sys.argv)
	ex = PinYin()
	ex.show()
	sys.exit(App.exec_())
