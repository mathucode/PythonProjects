import matplotlib.pyplot as plt
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from random import randint
from docx2pdf import convert
import openpyxl

doc = DocxTemplate("reportTmpl.docx")

salesRows = []
topItems = []

df = pd.read_excel(r"C:/Users/suvesh/Documents/data_items.xlsx")

for index, row in df.iterrows():
   sNo = row['S.No.']
   name = row['Item Name']
   costPu = row['Cost Per Unit']
   nUnits = row['Units Sold']
   salesRows.append({"sNo": sNo, "name": name, "cPu": costPu, "nUnits": nUnits, "revenue": costPu * nUnits})


"""""
for iItr in range(10):
  costPu = randint(1, 15)
  nUnits = randint(100, 500)
  salesRows.append({"sNo": iItr+1, "name": "Item "+str(iItr+1), "cPu": costPu, "nUnits": nUnits, "revenue": costPu*nUnits})
"""

topItems = [x["name"] for x in sorted(salesRows, key=lambda x: x["revenue"], reverse=True)][0:3]

fig, ax = plt.subplots()
ax.bar([x["name"] for x in salesRows], [x["revenue"] for x in salesRows])
fig.tight_layout()
fig.savefig("trendImg.png")

context = {
   "reportDtStr": "02-07-2023",
   "salesTblRows": salesRows,
   "topItemsRows": topItems,
   "trendImg": InlineImage(doc, "trendImg.png")
}

doc.render(context)

doc.save("report.docx")
convert("report.docx", "reports.pdf")
