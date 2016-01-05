from asposewords import Settings
from com.aspose.words import Document

class LoadAndSaveToDisk:

    def __init__(self):
        dataDir = Settings.dataDir + 'quickstart/'
        
        doc = Document(dataDir + 'Document.doc')
        
        doc.save(dataDir + 'Document_Out.doc')
        
        print "Document saved."

if __name__ == '__main__':           
    LoadAndSaveToDisk()