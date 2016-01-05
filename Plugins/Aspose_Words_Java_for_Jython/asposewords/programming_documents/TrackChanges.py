from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import SaveFormat

class TrackChanges:

    def __init__(self):
        dataDir = Settings.dataDir + 'programming_documents/'
        
        doc = Document(dataDir + "trackDoc.doc")

        doc.acceptAllRevisions()
        
        doc.save(dataDir + "AcceptChanges.doc", SaveFormat.DOC)

        print "Done."

if __name__ == '__main__':
    TrackChanges()