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

while True:
	if l == 0:
		break
	try:
		c = tb[l].cell(x, 2).text
		d = tb[l].cell(x, 4).text
	except BaseException as e:
		print(e)
		break
	print(c, d)
	x += 1
