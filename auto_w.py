# coding=utf-8
from docx import Document
from pypinyin import Style, lazy_pinyin
from read_word import read_head


def read(docx_in):
	# print(docx_in)
	test_style = {"功能性": "F", "性能效率": "P", "信息安全性": "IS", "兼容性": "Sc", "易用性": "Su", "可靠性": "Sr", "可维护": "Sm",
				  "可移植": "Sp"}
	re_style = read_head(docx_in)
	print(re_style)

	doc = Document(docx_in)
	tb = doc.tables
	# example_name = []
	# test_item = []
	temp_out = []
	t_number = 5
	temp_map = None  # 零时保存转换后的字母
	while t_number < 5 + len(re_style):
		temp_a = tb[t_number].cell(0, 2).text
		temp_b = tb[t_number].cell(0, 4).text
		if "用例标识" in temp_a and "用例名称" in temp_b:
			for index, tb_row in enumerate(tb[t_number].column_cells(4)):
				if index is 0:
					continue
				# example_name.append("".join(lazy_pinyin(tb_row.text, style=Style.FIRST_LETTER)).upper())
				# test_item.append(
				# 	"".join(lazy_pinyin(tb[t_number].column_cells(0)[index].text, style=Style.FIRST_LETTER)).upper())
				# temp_map = list(map(list, zip(test_item, example_name)))
				temp_out.append(
					"".join(lazy_pinyin(tb_row.text, style=Style.FIRST_LETTER)).upper() + "-" + test_style[
						re_style[t_number - 5]] + "-" + "".join(
						lazy_pinyin(tb[t_number].column_cells(0)[index].text,
									style=Style.FIRST_LETTER)).upper() + "-" + str(index).zfill(3))
			print(temp_out)
		t_number += 1
	temp_num = 0
	# 填写编号
	t_number = 5
	while t_number < 5 + len(re_style):
		temp_a = tb[t_number].cell(0, 2).text
		temp_b = tb[t_number].cell(0, 4).text
		if "用例标识" in temp_a and "用例名称" in temp_b:
			for index, tb_row in enumerate(tb[t_number].column_cells(4)):
				if index is 0:
					continue
				print(temp_out[temp_num])

				tb[t_number].column_cells(2)[index].text = temp_out[temp_num]
				temp_num += 1
		t_number += 1

	# doc.save(docx_in)
	return "执行完成"


"""
	temp_out = []
	# 生成编号
	for i in range(len(temp_map)):
		tstr = str(i + 1).zfill(3)
		temp_out.append(temp_map[i][0] + "-" + "F" + "-" + temp_map[i][1] + "-" + tstr)
"""

if __name__ == '__main__':
	read("C:\\Users\\ULTRAMA\\Desktop\\工具\\test.docx")
# read(input("请复制文件完整路径："))
