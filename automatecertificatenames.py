from docx import Document
from docx2pdf import convert
import csv


# Load .csv to List
replaceNames = []
with open('nameList.csv', newline='') as inputfile:
    for row in csv.reader(inputfile):
        replaceNames.append(row[0])


def automateIt(replaceNames):
    i = 0
    while i < len(replaceNames):
        doc = Document('certifcate.docx')
        findandReplace(doc, toFind_Name, replaceNames[i])
        doc.save('Done\{name}.docx'.format(name=replaceNames[i]))
        print("Done for " + replaceNames[i])
        i += 1


def findandReplace(doc_obj, toFind_Name, toReplace):
    for p in doc_obj.paragraphs:
        if toFind_Name in p.text:
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                if toFind_Name in inline[i].text:
                    text = inline[i].text.replace(toFind_Name, toReplace)
                    inline[i].text = text
                else:
                    print("else")


toFind_Name = "SomeName"
automateIt(replaceNames)
convert("Done\\")
