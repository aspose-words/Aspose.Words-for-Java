from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import Node
from com.aspose.words import Paragraph

class WorkingWithNodes:

    def __init__(self):
        # Create a new document.
        doc = Document()

        # Creates and adds a paragraph node to the document.
        para = Paragraph(doc)

        # Typed access to the last section of the document.
        section = doc.getLastSection()
        section.getBody().appendChild(para)

        # Next print the node type of one of the nodes in the document.
        nodeType = doc.getFirstSection().getBody().getNodeType()

        print "NodeType: " + Node.nodeTypeToString(nodeType)

if __name__ == '__main__':       
    WorkingWithNodes()