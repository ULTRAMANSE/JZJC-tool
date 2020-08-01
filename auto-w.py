from docx import Document

doc = Document("./test1.docx")
tb = doc.tables
i = 5
x = 1
l = 0
while True:
	print(1)
	a = tb[i].cell(0, 2).text
	b = tb[i].cell(0, 4).text
	if "用例标识" in a and "用例名称" in b:
		l = i
		break
	i += 1

if l is not 0:
	for i, t in enumerate(tb[l].column_cells(2)):
		# print(tb[l].columns)
		# c = tb[l].cell(x, 2).text
		print(i, t.text)
		print(tb[l].column_cells(4)[i].text)
	# d = tb[l].cell(x, 4).text
