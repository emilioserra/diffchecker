# -*- coding: utf-8 -*-
"""
Script for detection of copies in docx files.
It generates ordenated list with match ratios between pairs of docx documents. 
"""
from difflib import SequenceMatcher
import os
import docx #pip install python-docx 
from operator import itemgetter


#get couples of files
def getListOfDocx():
    filesList = []
    for file in os.listdir("./"):
        if file.endswith(".docx"):
            filesList.append(os.path.join("./", file))
    return filesList

#from word to text
def docx2Text(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

 
filesList = getListOfDocx()


matchList =[]

#iterate over all possible pair of different documents
for i in range(len(filesList)-1):
    for j in range(i+1, len(filesList)):

        x=filesList[i]
        y=filesList[j]
        #match between two documents
        m = SequenceMatcher(None, docx2Text(x), docx2Text(y))
        percentage = round(m.ratio(),3) 
        matchList.append([percentage, x, y])
          
 
#order the list of list using the field 0, the match percentage
matchList= sorted(matchList, key=itemgetter(0))

print("DIFF REPORT")
print("=======================")
for s in matchList:
    print(*s)
