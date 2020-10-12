from docxtpl import DocxTemplate
import xlrd

excel = xlrd.open_workbook('2020福布斯中国富豪榜名单列表.xlsx')

sheet = excel.sheets()[0]
print("共", sheet.nrows, "行")

names = []
categories = []

for i in range(sheet.nrows):
    names.append(sheet.cell_value(i, 0))
    categories.append(sheet.cell_value(i, 3))


for name,category in zip(names, categories):
    doc = DocxTemplate("template.docx")
    context = {'name': name,'category':category}
    doc.render(context)
    doc.save(name + ".docx")
    print(name + ".docx finished")