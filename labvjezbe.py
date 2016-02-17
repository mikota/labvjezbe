try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile
from os import listdir
from os.path import isfile,join
import os
"""
Module that extract text from MS XML Word document (.docx).
(Inspired by python-docx <https://github.com/mikemaccana/python-docx>)
"""

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'


def get_docx_text(path):
    """
    Take the path of a docx file as argument, return the text in unicode.
    """
    document = zipfile.ZipFile(path)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = XML(xml_content)

    paragraphs = []
    for paragraph in tree.getiterator(PARA):
        texts = [node.text
                 for node in paragraph.getiterator(TEXT)
                 if node.text]
        if texts:
            paragraphs.append(''.join(texts))

    return '\n\n'.join(paragraphs)

#POCETAK
imena=["lovro","ivona","isabela","lucija","paula","kaltrina","antonio","stela","tena","tomislav","magdalena","lorena","mirta","adrian","elizabeta","karlo","ilan","dorian","marko","martina","bernard","bruno","luka","morena","matej"]
onlyfiles = [f for f in listdir(os.getcwd()) if isfile(join(os.getcwd(), f))]
docs=[]
for j in onlyfiles:
    if ".doc" in j and not "~$" in j:
        docs.append(j)
for i in range(len(imena)):
    imena[i]=imena[i].lower()
    #za svaki slucaj
predali,nisu,warning=[],[],[]
for k in docs:
    a=get_docx_text(k)
    a=a.lower()
    warn=True
    for i in range(len(imena)):
        if imena[i] in a and imena[i] not in predali:
            predali.append(imena[i])
            i=len(imena)+3
            warn=False
    if warn==True:
        warning.append(k)
if warning!=[]:
    print " Upozorenje! Sljedeci dokumenti postoje, ali nisu potpisani:"
    if len(warning)==1:
        print warning[0][0].upper()+warning[0][1::],"\n"
    else:
        for i in warning[0:-1]:
            zarez=","
            if i==warning[-2]:
                zarez=""
            print i[0].upper()+i[1::]+zarez,
        print "i",nisu[-1][0].upper()+nisu[-1][1::],"\n"
print "- Ukupan broj predalih:",len(predali),"\n"

nitko=""
if predali==[]:
    nitko="nitko\n\n"
print "- Predali su:",nitko,
if nitko=="":
    if len(predali)==1:
        print predali[0][0].upper()+predali[0][1::],"\n"
    else:
        for i in predali[0:-1]:
            zarez=","
            if i==predali[-2]:
                zarez=""
            print i[0].upper()+i[1::]+zarez,
        print "i",predali[-1][0].upper()+predali[-1][1::],"\n"
for i in imena:
    if i not in predali:
        nisu.append(i)
nitko=""
if nisu==[]:
    nitko="nitko\n\n"
print "- Nisu predali:",nitko,
if nitko=="":
    if len(nisu)==1:
        print nisu[0][0].upper()+nisu[0][1::],"\n"
    else:
        for i in nisu[0:-1]:
            zarez=","
            if i==nisu[-2]:
                zarez=""
            print i[0].upper()+i[1::]+zarez,
        print "i",nisu[-1][0].upper()+nisu[-1][1::],"\n"
