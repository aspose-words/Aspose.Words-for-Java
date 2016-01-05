from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import SaveFormat

class CloneDocument:

    def __init__(self):
        dataDir = Settings.dataDir + 'programming_documents/'
        
        doc = Document(dataDir + "document.doc")

        clone = doc.deepClone()
        
        clone.save(dataDir + "CloneDocument.doc", SaveFormat.DOC)

        print "Done."

if __name__ == '__main__':
    CloneDocument()