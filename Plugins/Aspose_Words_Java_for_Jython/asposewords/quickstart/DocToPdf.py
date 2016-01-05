from asposewords import Settings
from com.aspose.words import Document

class DocToPdf:

    def __init__(self):
        dataDir = Settings.dataDir + 'quickstart/'
        
        doc = Document(dataDir + 'Document.doc')
        
        doc.save(dataDir + 'Document.pdf')
        
        print "Converted document to PDF."

if __name__ == '__main__':           
    DocToPdf()