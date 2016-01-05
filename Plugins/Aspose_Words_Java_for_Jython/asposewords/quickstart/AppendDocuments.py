from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import ImportFormatMode

class AppendDocuments:

    def __init__(self):
        dataDir = Settings.dataDir + 'quickstart/'
        
        dstDoc = Document(dataDir + 'TestFile.Destination.doc')
        srcDoc = Document(dataDir + 'TestFile.Source.doc')

        dstDoc.appendDocument(srcDoc,ImportFormatMode.KEEP_SOURCE_FORMATTING)
        
        dstDoc.save(dataDir + 'AppendDocuments.doc')
        
        print ("Documents appended successfully.")
        
if __name__ == '__main__':        
    AppendDocuments()