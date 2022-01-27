

import docx

from docx.shared import Pt

doc = docx.Document()

#style = doc.styles['Normal']
#font = style.font
#font.name = 'Courier New'
#font.size = Pt(10)


# CHANGE THIS TEXT FILE FOR THE DIVISION ----------------------------

f = open("Barrie employee abstracts from MTO - Jan 2020.txt", "r")

Lines = f.readlines()

#currline = f.readline()
#doc.add_paragraph(currline)

#count = 0

for line in Lines:
	doc.add_paragraph(line.strip())
	if "END OF RECORD" in line:
		doc.add_page_break()



f.close()


# CHANGE THIS DOCX FILE FOR THE DIVISION ----------------------------

doc.save('Barrie employee abstracts from MTO - Jan 2020test.docx')




