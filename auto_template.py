from docxtpl import DocxTemplate
import xlrd


def auto(in_path):
	workbook = xlrd.open_workbook(".\\template.xlsx")
	sheet = workbook.sheet_by_index(0)
	text = sheet.col_values(1)

	Outline = DocxTemplate(".\\模板\JZJC-O-000X-2021 软件测试大纲.docx")
	Description = DocxTemplate(".\\模板\\JZJC-C-000X-2021 软件测试说明.docx")
	logs = DocxTemplate(".\\模板\\JZJC-R-000X-2021 软件测试记录.docx")

	context = {
		"number": text[0], "con_number": text[1], "sample_number": text[2], "rev_staff": text[3],
		"soft_name": text[4], "version": text[5], "requester": text[6], "deve_unit": text[7],
		"u_address": text[8], "item_n": text[9], "ph_number": text[10], "email": text[11],
		"postcode": text[12], "contact_per": text[13], "t_address": text[14], "accept_date": text[15],
		"b_date": text[16], "r_date": text[17], "manual": text[18], "test_type": text[19],
		"test_item_type": text[20], "description": text[21]
	}

	Outline.render(context)
	Description.render(context)
	logs.render(context)
	Outline.save(in_path + "\\JZJC-O-" + text[0] + " 测试大纲.docx")
	Description.save(in_path + "\\JZJC-C-" + text[0] + " 测试说明.docx")
	logs.save(in_path + "\\JZJC-R-" + text[0] + " 测试记录.docx")


if __name__ == '__main__':
	auto(input("输入文件目录"))
