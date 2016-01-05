from asposewords import Settings
from com.aspose.words import Document

class RemoveField:

    def __init__(self):
        dataDir = Settings.dataDir + 'programming_documents/'
        
        doc = Document(dataDir + "RemoveField.doc")

        field = doc.getRange().getFields().get(0)
        
        # Calling this method completely removes the field from the document.
        field.remove()
        
        doc.save(dataDir + "RemoveField Out.docx")

        print "Field removed from the document successfully."

if __name__ == '__main__':        
    RemoveField()