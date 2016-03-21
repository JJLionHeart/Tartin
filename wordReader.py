"""
Experiment code no 1: wordReader

Description: This piece of code is supposed to take a word file and read its contents

Entry no. 1
03/20/2016
I'll begin with finding out how to open a word file and read the first sentence or paragraph

Thanks to etienned for his piece of code at https://gist.github.com/etienned/7539105

the code that reads text from a word file will be inspired on his work

"""


from xml.etree.cElementTree import XML
import zipfile

def get_raw_text(pthFile):
	"""
	gets a path to a file as an argument and returns a list containing
	the paragraphs of the word document file
	"""
	
	"""
	Constants used to iterate over the XML tree
	"""
	WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
	PARA = WORD_NAMESPACE + 'p'
	TEXT = WORD_NAMESPACE + 't'

	docWordDoc = zipfile.ZipFile(pthFile) #gets the documents of the word
	xmlContent = docWordDoc.read('word/document.xml') #access the xml file
	docWordDoc.close()
	treeXML = XML(xmlContent) #parses the xml content into a tree that will be further used to access the text

	lstParagraphs = [] #output list with the paragraphs of the text
	#now we proceed to extract the text from the tree
	#the idea is to iterate over the tree and 
	#for each node that contains text, substract it and add it to
	#the output
	for parParagraph in treeXML.getiterator(PARA):
		lstTexts = [nodElement.text
			    for nodElement in parParagraph.getiterator(TEXT)
			    if nodElement.text]
		if lstTexts:
			print lstTexts
			lstParagraphs.append(''.join(lstTexts))
		
	return lstParagraphs


print get_raw_text("prueba.docx")
