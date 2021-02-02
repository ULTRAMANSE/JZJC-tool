from docxtpl import DocxTemplate
import xlrd
from read_word import word


def auto(a, b):
	workbook = xlrd.open_workbook(a)
	sheet = workbook.sheet_by_index(0)
	text = sheet.col_values(1)

	tpl = DocxTemplate("./template.docx")

	features, performance, safety, compatible = [], [], [], []
	easy, safeguard, reliable, transplant = [], [], [], []
	w = word(b)
	try: # 有待改进
		if "功能" in text[19]:
			features = next(w)
		if "性能" in text[19]:
			performance = next(w)
		if "安全" in text[19]:
			safety = next(w)
		if "兼容" in text[19]:
			compatible = next(w)
		if "易用" in text[19]:
			easy = next(w)
		if "维护" in text[19]:
			safeguard = next(w)
		if "可靠" in text[19]:
			reliable = next(w)
		if "移植" in text[19]:
			transplant = next(w)
	except StopIteration:
		pass

	x = {"features": features, "performance": performance, "safety": safety, "compatible": compatible, "easy": easy,
		 "safeguard": safeguard, "reliable": reliable, "transplant": transplant}
	context = {
		"number": text[0], "con_number": text[1], "sample_number": text[2], "rev_staff": text[3],
		"soft_name": text[4], "version": text[5], "requester": text[6], "deve_unit": text[7],
		"u_address": text[8], "item_n": text[9], "ph_number": text[10], "email": text[11],
		"postcode": text[12], "contact_per": text[13], "t_address": text[14], "accept_date": text[15],
		"b_date": text[16], "r_date": text[17], "manual": text[18], "test_type": text[19],
		"test_item_type": text[20], "description": text[21], "test_item": x
	}
	tpl.render(context)
	tpl.save(".\out.docx")


if __name__ == '__main__':
	auto("./template.xlsx", "C:\\Users\\ULTRAMA\\Desktop\\工具\\test.docx")
