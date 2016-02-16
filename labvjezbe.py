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
predali=[]
for k in docs:
    a=get_docx_text(k)
    a=a.lower()
    for i in range(len(imena)):
        if imena[i] in a and imena[i] not in predali:
            predali.append(imena[i])
            i=len(imena)+1
print "Predali su:",       
for i in predali:
    print i+",",
