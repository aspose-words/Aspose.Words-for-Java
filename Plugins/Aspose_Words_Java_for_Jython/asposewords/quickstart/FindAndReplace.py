from asposewords import Settings
from com.aspose.words import Document

class FindAndReplace:

    def __init__(self):
        dataDir = Settings.dataDir + 'quickstart/'
        
        doc = Document(dataDir + 'ReplaceSimple.doc')
        
        # Check the text of the document.
        print "Original document text: " + doc.getRange().getText()
        
        # Replace the text in the document.
        doc.getRange().replace("_CustomerName_", "James Bond", False, False)
        
        # Check the replacement was made.
        print "Document text after replace: " + doc.getRange().getText()
        
        doc.save(dataDir + 'ReplaceSimple_Out.doc')

if __name__ == '__main__':           
    FindAndReplace()