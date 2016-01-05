from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import NodeType
from com.aspose.words import NodeImporter
from com.aspose.words import ImportFormatMode
from com.aspose.words import SaveFormat
from com.aspose.words import DocumentBuilder

from java.util import Collections

class ExtractContent:

    def __init__(self):
        self.dataDir = Settings.dataDir + 'programming_documents/'
        
        # Open the document.
        doc = Document(self.dataDir + "TestFile.doc")
        
        self.extract_content_between_paragraphs(doc)
        self.extract_content_between_block_level_nodes(doc)
        self.extract_content_between_runs(doc)
        self.extract_content_using_field(doc)
        self.extract_content_between_bookmark(doc)
        self.extract_content_between_comment_range(doc)
        self.extract_content_between_paragraph_styles(doc)
    
    def extract_content_between_paragraphs(self, doc):
        """
            Shows how to extract the content between specific paragraphs using the ExtractContent method above.
        """
        
        # Gather the nodes. The GetChild method uses 0-based index
        startPara = doc.getFirstSection().getChild(NodeType.PARAGRAPH, 6, True)
        endPara = doc.getFirstSection().getChild(NodeType.PARAGRAPH, 10, True)
        # Extract the content between these nodes in the document. Include these markers in the extraction.
        extractedNodes = self.extract_contents(startPara, endPara, True)

        # Insert the content into a new separate document and save it to disk.
        dstDoc = self.generate_document(doc, extractedNodes)
        
        dstDoc.save(self.dataDir + "TestFile.Paragraphs Out.doc")
        
    def extract_content_between_block_level_nodes(self, doc):
        """
            Shows how to extract the content between a paragraph and table using the ExtractContent method.
        """

        startPara = doc.getLastSection().getChild(NodeType.PARAGRAPH, 2, True)
        endTable = doc.getLastSection().getChild(NodeType.TABLE, 0, True)

        # Extract the content between these nodes in the document. Include these markers in the extraction.
        extractedNodes = self.extract_contents(startPara, endTable, True)

        # Lets reverse the array to make inserting the content back into the document easier.
        Collections.reverse(extractedNodes[::-1])

        while (len(extractedNodes) > 0):
            # Insert the last node from the reversed list
            endTable.getParentNode().insertAfter(extractedNodes[0], endTable)
            # Remove this node from the list after insertion.
            del extractedNodes[0]

        # Save the generated document to disk.
        doc.save(self.dataDir + "TestFile.DuplicatedContent Out.doc")
        
    def extract_content_between_runs(self, doc):
        """
            Shows how to extract content between specific runs of the same paragraph using the ExtractContent method.
        """
        
        # Retrieve a paragraph from the first section.
        para = doc.getChild(NodeType.PARAGRAPH, 7, True)

        # Use some runs for extraction.
        startRun = para.getRuns().get(1)
        endRun = para.getRuns().get(4)

        # Extract the content between these nodes in the document. Include these markers in the extraction.
        extractedNodes = self.extract_contents(startRun, endRun, True)

        # Get the node from the list. There should only be one paragraph returned in the list.
        node = extractedNodes[0]
        
        # Print the text of this node to the console.
        print node.toString(SaveFormat.TEXT)
        
    def extract_content_using_field(self, doc):
        """
            Shows how to extract content between a specific field and paragraph in the document using the ExtractContent method.
        """

        # Use a document builder to retrieve the field start of a merge field.
        builder = DocumentBuilder(doc)

        # Pass the first boolean parameter to get the DocumentBuilder to move to the FieldStart of the field.
        # We could also get FieldStarts of a field using GetChildNode method as in the other examples.
        builder.moveToMergeField("Fullname", False, False)

        # The builder cursor should be positioned at the start of the field.
        startField = builder.getCurrentNode()
        endPara = doc.getFirstSection().getChild(NodeType.PARAGRAPH, 5, True)

        # Extract the content between these nodes in the document. Don't include these markers in the extraction.
        extractedNodes = self.extract_contents(startField, endPara, False)

        # Insert the content into a new separate document and save it to disk.
        dstDoc = self.generate_document(doc, extractedNodes)
        
        dstDoc.save(self.dataDir + "TestFile.Fields Out.pdf")
        
    def extract_content_between_bookmark(self, doc):
        """
            Shows how to extract the content referenced a bookmark using the ExtractContent method.
        """

        # Retrieve the bookmark from the document.
        bookmark = doc.getRange().getBookmarks().get("Bookmark1")

        # We use the BookmarkStart and BookmarkEnd nodes as markers.
        bookmarkStart = bookmark.getBookmarkStart()
        bookmarkEnd = bookmark.getBookmarkEnd()

        # Firstly extract the content between these nodes including the bookmark.
        extractedNodesInclusive = self.extract_contents(bookmarkStart, bookmarkEnd, True)
        dstDoc = self.generate_document(doc, extractedNodesInclusive)
        dstDoc.save(self.dataDir + "TestFile.BookmarkInclusive Out.doc")

        # Secondly extract the content between these nodes this time without including the bookmark.
        extractedNodesExclusive = self.extract_contents(bookmarkStart, bookmarkEnd, False)
        dstDoc = self.generate_document(doc, extractedNodesExclusive)
        dstDoc.save(self.dataDir + "TestFile.BookmarkExclusive Out.doc")
        
    def extract_content_between_comment_range(self, doc):
        """
            Shows how to extract content referenced by a comment using the ExtractContent method.
        """

        # This is a quick way of getting both comment nodes.
        # Your code should have a proper method of retrieving each corresponding start and end node.
        commentStart = doc.getChild(NodeType.COMMENT_RANGE_START, 0, True)
        commentEnd = doc.getChild(NodeType.COMMENT_RANGE_END, 0, True)

        # Firstly extract the content between these nodes including the comment as well.
        extractedNodesInclusive = self.extract_contents(commentStart, commentEnd, True)
        dstDoc = self.generate_document(doc, extractedNodesInclusive)
        dstDoc.save(self.dataDir + "TestFile.CommentInclusive Out.doc")

        # Secondly extract the content between these nodes without the comment.
        extractedNodesExclusive = self.extract_contents(commentStart, commentEnd, False)
        dstDoc = self.generate_document(doc, extractedNodesExclusive)
        dstDoc.save(self.dataDir + "TestFile.CommentExclusive Out.doc")
        
    def extract_content_between_paragraph_styles(self, doc):
        """
            Shows how to extract content between paragraphs with specific styles using the ExtractContent method.
        """

        # Gather a list of the paragraphs using the respective heading styles.
        parasStyleHeading1 = self.paragraphs_by_style_name(doc, "Heading 1")
        parasStyleHeading3 = self.paragraphs_by_style_name(doc, "Heading 3")

        # Use the first instance of the paragraphs with those styles.
        startPara1 = parasStyleHeading1[0]
        endPara1 = parasStyleHeading3[0]

        # Extract the content between these nodes in the document. Don't include these markers in the extraction.
        extractedNodes = self.extract_contents(startPara1, endPara1, False)

        # Insert the content into a new separate document and save it to disk.
        dstDoc = self.generate_document(doc, extractedNodes)
        
        dstDoc.save(self.dataDir + "TestFile.Styles Out.doc")

    def paragraphs_by_style_name(self, doc, styleName):
        # Create an array to collect paragraphs of the specified style.
        paragraphsWithStyle = []
        
        # Get all paragraphs from the document.
        paragraphs = doc.getChildNodes(NodeType.PARAGRAPH, True)
        
        # Look through all paragraphs to find those with the specified style.
        paragraphs_count = paragraphs.getCount()

        i = 0
        while(i < paragraphs_count) :
            paragraph = paragraphs.get(i)
            if (paragraph.getParagraphFormat().getStyle().getName() == styleName):
                paragraphsWithStyle.append(paragraph)
            i = i + 1

        return paragraphsWithStyle
    
    def extract_contents(self,startNode, endNode, isInclusive):
        """
            Extracts a range of nodes from a document found between specified markers and returns a copy of those nodes. Content can be extracted
            between inline nodes, block level nodes, and also special nodes such as Comment or Boomarks. Any combination of different marker types can used.

            @param startNode The node which defines where to start the extraction from the document. This node can be block or inline level of a body.
            @param endNode The node which defines where to stop the extraction from the document. This node can be block or inline level of body.
            @param isInclusive Should the marker nodes be included.
        """
        
        # First check that the nodes passed to this method are valid for use.
        self.verify_parameter_nodes(startNode, endNode)

        # Create a list to store the extracted nodes.
        nodes = []

        # Keep a record of the original nodes passed to this method so we can split marker nodes if needed.
        originalStartNode = startNode
        originalEndNode = endNode

        # Extract content based on block level nodes (paragraphs and tables). Traverse through parent nodes to find them.
        # We will split the content of first and last nodes depending if the marker nodes are inline
        while startNode.getParentNode().getNodeType() != NodeType.BODY :
            startNode = startNode.getParentNode()

        while (endNode.getParentNode().getNodeType() != NodeType.BODY):
            endNode = endNode.getParentNode()

        print str(originalStartNode) + " = " + str(startNode)
        print str(originalEndNode) + " = " + str(endNode)

        isExtracting = True
        isStartingNode = True

        # The current node we are extracting from the document.
        currNode = startNode

        # Begin extracting content. Process all block level nodes and specifically split the first and last nodes when needed so paragraph formatting is retained.
        # Method is little more complex than a regular extractor as we need to factor in extracting using inline nodes, fields, bookmarks etc as to make it really useful.
        while (isExtracting):

            # Clone the current node and its children to obtain a copy.
            cloneNode = currNode.deepClone(True)
            isEndingNode = currNode.equals(endNode)

            if(isStartingNode or isEndingNode):

                # We need to process each marker separately so pass it off to a separate method instead.
                if (isStartingNode):
                    self.process_marker(cloneNode, nodes, originalStartNode, isInclusive, isStartingNode, isEndingNode)
                    isStartingNode = False
                # Conditional needs to be separate as the block level start and end markers maybe the same node.
                if (isEndingNode):
                    self.process_marker(cloneNode, nodes, originalEndNode, isInclusive, isStartingNode, isEndingNode)
                    isExtracting = False
            else:
                # Node is not a start or end marker, simply add the copy to the list.
                nodes.append(cloneNode)

            # Move to the next node and extract it. If next node is null that means the rest of the content is found in a different section.
            if (currNode.getNextSibling() is None and isExtracting):
                # Move to the next section.
                nextSection = currNode.getAncestor(NodeType.SECTION).getNextSibling()
                currNode = nextSection.getBody().getFirstChild()
            else:
                # Move to the next node in the body.
                currNode = currNode.getNextSibling()


        # Return the nodes between the node markers.
        return nodes
    # ExEnd

    def verify_parameter_nodes(self,startNode,endNode):
        """
            Checks the input parameters are correct and can be used. Throws an exception if there is any problem.
        """
    
        # The order in which these checks are done is important.
        if (startNode is None):
            raise ValueError('Start node cannot be null')
        if (endNode is None):
            raise ValueError('End node cannot be null')
        if (startNode.getDocument() != endNode.getDocument()):
            raise ValueError('Start node and end node must belong to the same document')
        if (startNode.getAncestor(NodeType.BODY) is None or endNode.getAncestor(NodeType.BODY) is None):
            raise ValueError('Start node and end node must be a child or descendant of a body')

        # Check the end node is after the start node in the DOM tree
        # First check if they are in different sections, then if they're not check their position in the body of the same section they are in.
        startSection = startNode.getAncestor(NodeType.SECTION)
        endSection = endNode.getAncestor(NodeType.SECTION)

        startIndex = startSection.getParentNode().indexOf(startSection)
        endIndex = endSection.getParentNode().indexOf(endSection)

        if (startIndex == endIndex):

            if (startSection.getBody().indexOf(startNode) > endSection.getBody().indexOf(endNode)):
                raise ValueError('The end node must be after the start node in the body')

        elif (startIndex > endIndex):
            raise ValueError('The section of end node must be after the section start node')

    def is_inline(self,node):
        # Test if the node is desendant of a Paragraph or Table node and also is not a paragraph or a table a paragraph inside a comment class which is decesant of a pararaph is possible.
        return ((node.getAncestor(NodeType.PARAGRAPH) is not None or node.getAncestor(NodeType.TABLE) is not None) and not(node.getNodeType() == NodeType.PARAGRAPH or node.getNodeType() == NodeType.TABLE))

    def process_marker(self,cloneNode,nodes,node, isInclusive, isStartMarker, isEndMarker):
        
        # If we are dealing with a block level node just see if it should be included and add it to the list.
        if(not(self.is_inline(node))):
            # Don't add the node twice if the markers are the same node
            if(not(isStartMarker and isEndMarker)):
                if (isInclusive):
                    nodes.append(cloneNode)
            return

        # If a marker is a FieldStart node check if it's to be included or not.
        # We assume for simplicity that the FieldStart and FieldEnd appear in the same paragraph.
        if (node.getNodeType() == NodeType.FIELD_START):
            # If the marker is a start node and is not be included then skip to the end of the field.
            # If the marker is an end node and it is to be included then move to the end field so the field will not be removed.
            if ((isStartMarker and not(isInclusive)) or (not(isStartMarker) and isInclusive)):
                while ((node.getNextSibling() is not None) and (node.getNodeType() != NodeType.FIELD_END)):
                    node = node.getNextSibling()


        # If either marker is part of a comment then to include the comment itself we need to move the pointer forward to the Comment
        # node found after the CommentRangeEnd node.
        if (node.getNodeType() == NodeType.COMMENT_RANGE_END):
            while (node.getNextSibling() is not None and node.getNodeType() != NodeType.COMMENT):
                node = node.getNextSibling()


        # Find the corresponding node in our cloned node by index and return it.
        # If the start and end node are the same some child nodes might already have been removed. Subtract the
        # difference to get the right index.
        indexDiff = node.getParentNode().getChildNodes().getCount() - cloneNode.getChildNodes().getCount()

        # Child node count identical.
        if (indexDiff == 0):
            node = cloneNode.getChildNodes().get(node.getParentNode().indexOf(node))
        else:
            node = cloneNode.getChildNodes().get(node.getParentNode().indexOf(node) - indexDiff)

        # Remove the nodes up to/from the marker.
        isProcessing = True
        isRemoving = isStartMarker
        nextNode = cloneNode.getFirstChild()

        while (isProcessing and nextNode is not None):

            currentNode = nextNode
            isSkip = False

            if (currentNode.equals(node)):
                if (isStartMarker):
                    isProcessing = False
                    if (isInclusive):
                        isRemoving = False
                else:
                    isRemoving = True
                    if (isInclusive):
                        isSkip = True
            nextNode = nextNode.getNextSibling()
            if (isRemoving and not(isSkip)):
                currentNode.remove()

        # After processing the composite node may become empty. If it has don't include it.
        if (not(isStartMarker and isEndMarker)):
            if (cloneNode.hasChildNodes()):
                nodes.append(cloneNode)

    def generate_document(self,srcDoc,nodes):

        # Create a blank document.
        dstDoc = Document()
        # Remove the first paragraph from the empty document.
        dstDoc.getFirstSection().getBody().removeAllChildren()

        # Import each node from the list into the new document. Keep the original formatting of the node.
        importer = NodeImporter(srcDoc, dstDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING)

        for node in nodes:
            importNode = importer.importNode(node, True)
            dstDoc.getFirstSection().getBody().appendChild(importNode)

        # Return the generated document.
        return dstDoc

if __name__ == '__main__':        
    ExtractContent()