# coding=utf-8
from docxtpl import DocxTemplate
import xlrd
from read_word import word, read_head


def auto(template_xlsx, record_docx):
	workbook = xlrd.open_workbook(template_xlsx)
	sheet = workbook.sheet_by_index(0)
	text = sheet.col_values(1)

	tpl = DocxTemplate("./template.docx")

	# features, performance, safety, compatible = [], [], [], []
	# easy, safeguard, reliable, transplant = [], [], [], []
	w = word(record_docx)
	temp = read_head(record_docx)
	c = {}
	for i in temp:
		c[i] = next(w)
	# print(c)
	context = {
		"number": text[0], "con_number": text[1], "sample_number": text[2], "rev_staff": text[3],
		"soft_name": text[4], "version": text[5], "requester": text[6], "deve_unit": text[7],
		"u_address": text[8], "item_n": text[9], "ph_number": text[10], "email": text[11],
		"postcode": text[12], "contact_per": text[13], "t_address": text[14], "accept_date": text[15],
		"b_date": text[16], "r_date": text[17], "manual": text[18], "test_type": text[19],
		"test_item_type": text[20], "description": text[21], "test_item": c
	}

	tpl.render(context)
	tpl.save(".\out.docx")


if __name__ == '__main__':
	auto("./template.xlsx", "H:\\WK存\\test\\JZJC.docx")
