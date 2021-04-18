# coding=utf-8
from docxtpl import DocxTemplate
import xlrd
from read_word import word, read_head


def auto(record_docx):
	workbook = xlrd.open_workbook("./template.xlsx")
	sheet = workbook.sheet_by_index(0)
	text = sheet.col_values(1)
	tpl = DocxTemplate("./模板/JZJC-TR-000X-2021 软件测试报告.docx")
	w = word(record_docx)
	temp = read_head(record_docx)
	c = {}
	for i in temp:
		c[i] = next(w)

	context = {
		"number": text[0], "con_number": text[1], "sample_number": text[2], "rev_staff": text[3],
		"soft_name": text[4], "version": text[5], "requester": text[6], "deve_unit": text[7],
		"u_address": text[8], "item_n": text[9], "ph_number": text[10], "email": text[11],
		"postcode": text[12], "contact_per": text[13], "t_address": text[14], "accept_date": text[15],
		"b_date": text[16], "r_date": text[17], "manual": text[18], "test_type": text[19],
		"test_item_type": text[20], "description": text[21], "test_item": c
	}

	tpl.render(context)
	tpl.save(".\JZJC-TR-" + text[0] + " 软件测试报告.docx")


if __name__ == '__main__':
	auto("C:\\Users\\ULTRAMA\\Desktop\\123.docx")
