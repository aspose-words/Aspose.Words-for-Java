from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import NodeType

class JoiningTables:

    def __init__(self):
        dataDir = Settings.dataDir + 'programming_documents/'
        
        # Load the document.
        doc = Document(dataDir + "tableDoc.doc")
        
        # Get the first and second table in the document.
        # The rows from the second table will be appended to the end of the first table.
        firstTable = doc.getChild(NodeType.TABLE, 0, True)
        secondTable = doc.getChild(NodeType.TABLE, 1, True)

        # Append all rows from the current table to the next.
        # Due to the design of tables even tables with different cell count and widths can be joined into one table.
        while (secondTable.hasChildNodes()) :
            firstTable.getRows().add(secondTable.getFirstRow())

        # Remove the empty table container.
        secondTable.remove()

        doc.save(dataDir + "JoinTables.doc")

        print "Done."

if __name__ == '__main__':
    JoiningTables()