from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import DocumentBuilder

class MovingCursor:

    def __init__(self):
        dataDir = Settings.dataDir + 'programming_documents/'
        
        doc = Document(dataDir + "document.doc")

        builder = DocumentBuilder(doc);

        #Shows how to access the current node in a document builder.
        curNode = builder.getCurrentNode()
        curParagraph = builder.getCurrentParagraph()

        # Shows how to move a cursor position to a specified node.
        builder.moveTo(doc.getFirstSection().getBody().getLastParagraph())

        # Shows how to move a cursor position to the beginning or end of a document.
        builder.moveToDocumentEnd()
        builder.writeln("This is the end of the document.")

        builder.moveToDocumentStart()
        builder.writeln("This is the beginning of the document.")

        doc.save(dataDir + "MovingCursor.doc");

        print "Done."

if __name__ == '__main__':
    MovingCursor()