__author__ = 'fahadadeel'
import jpype

class AppendDocument:

    def __init__(self,gDataDir):

        self.gDataDir = gDataDir
        self.Document = jpype.JClass("com.aspose.words.Document")
        self.DocumentBuilder = jpype.JClass("com.aspose.words.DocumentBuilder")
        self.ImportFormatMode = jpype.JClass("com.aspose.words.ImportFormatMode")
        self.SectionStart = jpype.JClass("com.aspose.words.SectionStart")
        self.MessageFormat = jpype.JClass("java.text.MessageFormat")

    def main(self):

        self.appendDocument_SimpleAppendDocument()
        self.appendDocument_KeepSourceFormatting()
        self.appendDocument_UseDestinationStyles()
        self.appendDocument_JoinContinuous()
        self.appendDocument_JoinNewPage()
        self.appendDocument_RestartPageNumbering()
        self.appendDocument_LinkHeadersFooters()
        self.appendDocument_UnlinkHeadersFooters()
        self.appendDocument_RemoveSourceHeadersFooters()
        self.appendDocument_DifferentPageSetup()


    def appendDocument_SimpleAppendDocument(self):

        dstDoc = self.Document(self.gDataDir + "TestFile.Destination.doc")
        srcDoc = self.Document(self.gDataDir + "TestFile.Source.doc")

        #ExStart
        #ExId:AppendDocument_SimpleAppend
        #ExSummary:Shows how to append a document to the end of another document using no additional options.

        # Append the source document to the destination document using no extra options.
        dstDoc.appendDocument(srcDoc, self.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        #ExEnd

        dstDoc.save(self.gDataDir + "TestFile.SimpleAppendDocument Out.docx")

    def appendDocument_KeepSourceFormatting(self):

        dstDoc = self.Document(self.gDataDir + "TestFile.Destination.doc")
        srcDoc = self.Document(self.gDataDir + "TestFile.Source.doc")

        #ExStart
        #ExId:AppendDocument_SimpleAppend
        #ExSummary:Shows how to append a document to the end of another document using no additional options.

        # Append the source document to the destination document using no extra options.
        dstDoc.appendDocument(srcDoc, self.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        #ExEnd

        dstDoc.save(self.gDataDir + "TestFile.KeepSourceFormatting Out.docx")

    def appendDocument_UseDestinationStyles(self):

        #ExStart
        #ExId:AppendDocument_UseDestinationStyles
        #ExSummary:Shows how to append a document to another document using the formatting of the destination document.
        # Load the documents to join.
        dstDoc = self.Document(self.gDataDir + "TestFile.Destination.doc")
        srcDoc = self.Document(self.gDataDir + "TestFile.Source.doc")

        # Append the source document using the styles of the destination document.
        dstDoc.appendDocument(srcDoc, self.ImportFormatMode.USE_DESTINATION_STYLES)

        # Save the joined document to disk.
        dstDoc.save(self.gDataDir + "TestFile.UseDestinationStyles Out.doc")
        #ExEnd

    def appendDocument_JoinContinuous(self):

        #ExStart
        #ExId:AppendDocument_JoinContinuous
        #ExSummary:Shows how to append a document to another document so the content flows continuously.
        dstDoc = self.Document(self.gDataDir + "TestFile.Destination.doc")
        srcDoc = self.Document(self.gDataDir + "TestFile.Source.doc")

        # Make the document appear straight after the destination documents content.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(self.SectionStart.CONTINUOUS)

        # Append the source document using the original styles found in the source document.
        dstDoc.appendDocument(srcDoc, self.ImportFormatMode.KEEP_SOURCE_FORMATTING)
        dstDoc.save(self.gDataDir + "TestFile.JoinContinuous Out.doc")
        #ExEnd

    def appendDocument_JoinNewPage(self):

        #ExStart
        #ExId:AppendDocument_JoinNewPage
        #ExSummary:Shows how to append a document to another document so it starts on a new page.
        dstDoc = self.Document(self.gDataDir + "TestFile.Destination.doc")
        srcDoc = self.Document(self.gDataDir + "TestFile.Source.doc")

        # Set the appended document to start on a new page.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(self.SectionStart.NEW_PAGE)

        # Append the source document using the original styles found in the source document.
        dstDoc.appendDocument(srcDoc, self.ImportFormatMode.KEEP_SOURCE_FORMATTING)
        dstDoc.save(self.gDataDir + "TestFile.JoinNewPage Out.doc")
        # ExEnd

    def appendDocument_RestartPageNumbering(self):

        #ExStart
        #ExId:AppendDocument_RestartPageNumbering
        #ExSummary:Shows how to append a document to another document with page numbering restarted.
        dstDoc = self.Document(self.gDataDir + "TestFile.Destination.doc")
        srcDoc = self.Document(self.gDataDir + "TestFile.Source.doc")

        # Set the appended document to appear on the next page.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(self.SectionStart.NEW_PAGE)
        # Restart the page numbering for the document to be appended.
        srcDoc.getFirstSection().getPageSetup().setRestartPageNumbering(1)

        dstDoc.appendDocument(srcDoc, self.ImportFormatMode.KEEP_SOURCE_FORMATTING)
        dstDoc.save(self.gDataDir + "TestFile.RestartPageNumbering Out.doc")
        #ExEnd

    def appendDocument_LinkHeadersFooters(self):

        #ExStart
        #ExFor:HeaderFooterCollection.LinkToPrevious(Boolean)
        #ExId:AppendDocument_LinkHeadersFooters
        #ExSummary:Shows how to append a document to another document and continue headers and footers from the destination document.
        dstDoc = self.Document(self.gDataDir + "TestFile.Destination.doc")
        srcDoc = self.Document(self.gDataDir + "TestFile.Source.doc")

        # Set the appended document to appear on a new page.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(self.SectionStart.NEW_PAGE)

        # Link the headers and footers in the source document to the previous section.
        # This will override any headers or footers already found in the source document.
        srcDoc.getFirstSection().getHeadersFooters().linkToPrevious(1)

        dstDoc.appendDocument(srcDoc, self.ImportFormatMode.KEEP_SOURCE_FORMATTING)
        dstDoc.save(self.gDataDir + "TestFile.LinkHeadersFooters Out.doc")
        #ExEnd

    def appendDocument_UnlinkHeadersFooters(self):

        #ExStart
        #ExId:AppendDocument_UnlinkHeadersFooters
        #ExSummary:Shows how to append a document to another document so headers and footers do not continue from the destination document.
        dstDoc = self.Document(self.gDataDir + "TestFile.Destination.doc")
        srcDoc = self.Document(self.gDataDir + "TestFile.Source.doc")

        # Even a document with no headers or footers can still have the LinkToPrevious setting set to True.
        # Unlink the headers and footers in the source document to stop this from continuing the headers and footers
        # from the destination document.
        srcDoc.getFirstSection().getHeadersFooters().linkToPrevious(0)

        dstDoc.appendDocument(srcDoc, self.ImportFormatMode.KEEP_SOURCE_FORMATTING)
        dstDoc.save(self.gDataDir + "TestFile.UnlinkHeadersFooters Out.doc")
        #ExEnd

    def appendDocument_RemoveSourceHeadersFooters(self):

        #ExStart
        #ExId:AppendDocument_RemoveSourceHeadersFooters
        #ExSummary:Shows how to remove headers and footers from a document before appending it to another document.
        dstDoc = self.Document(self.gDataDir + "TestFile.Destination.doc")
        srcDoc = self.Document(self.gDataDir + "TestFile.Source.doc")

        # Remove the headers and footers from each of the sections in the source document.
        for section in srcDoc.getSections().toArray():
            section.clearHeadersFooters()

        # Even after the headers and footers are cleared from the source document, the "LinkToPrevious" setting
        # for HeadersFooters can still be set. This will cause the headers and footers to continue from the destination
        # document. This should set to false to avoid this behaviour.
        srcDoc.getFirstSection().getHeadersFooters().linkToPrevious(0)

        dstDoc.appendDocument(srcDoc, self.ImportFormatMode.KEEP_SOURCE_FORMATTING)
        dstDoc.save(self.gDataDir + "TestFile.RemoveSourceHeadersFooters Out.doc")
        #ExEnd

    def appendDocument_DifferentPageSetup(self):

        #ExStart
        #ExId:AppendDocument_DifferentPageSetup
        #ExSummary:Shows how to append a document to another document continuously which has different page settings.
        dstDoc = self.Document(self.gDataDir + "TestFile.Destination.doc")
        srcDoc = self.Document(self.gDataDir + "TestFile.SourcePageSetup.doc")

        # Set the source document to continue straight after the end of the destination document.
        # If some page setup settings are different then this may not work and the source document will appear
        # on a new page.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(self.SectionStart.CONTINUOUS)

        # To ensure this does not happen when the source document has different page setup settings make sure the
        # settings are identical between the last section of the destination document.
        # If there are further continuous sections that follow on in the source document then this will need to be
        # repeated for those sections as well.
        srcDoc.getFirstSection().getPageSetup().setPageWidth(dstDoc.getLastSection().getPageSetup().getPageWidth())
        srcDoc.getFirstSection().getPageSetup().setPageHeight(dstDoc.getLastSection().getPageSetup().getPageHeight())
        srcDoc.getFirstSection().getPageSetup().setOrientation(dstDoc.getLastSection().getPageSetup().getOrientation())

        dstDoc.appendDocument(srcDoc, self.ImportFormatMode.KEEP_SOURCE_FORMATTING)
        dstDoc.save(self.gDataDir + "TestFile.DifferentPageSetup Out.doc")
        #ExEnd

class CopyBookmarkedText:

    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Document = jpype.JClass("com.aspose.words.Document")
        self.NodeImporter = jpype.JClass("com.aspose.words.NodeImporter")
        self.ImportFormatMode = jpype.JClass("com.aspose.words.ImportFormatMode")

    def main(self):

        # Load the source document.
        srcDoc = self.Document(self.dataDir + "Template.doc")
        # This is the bookmark whose content we want to copy.
        srcBookmark = srcDoc.getRange().getBookmarks().get("ntf010145060")
        # We will be adding to this document.
        dstDoc = self.Document()
        # Let's say we will be appending to the end of the body of the last section.
        dstNode = dstDoc.getLastSection().getBody()
        # It is a good idea to use this import context object because multiple nodes are being imported.
        # If you import multiple times without a single context, it will result in many styles created.
        importer = self.NodeImporter(srcDoc, dstDoc, self.ImportFormatMode.KEEP_SOURCE_FORMATTING)
        # Do it once.
        self.appendBookmarkedText(importer, srcBookmark, dstNode)
        # Do it one more time for fun.
        self.appendBookmarkedText(importer, srcBookmark, dstNode)
        # Save the finished document.
        dstDoc.save(self.dataDir + "Template Out.doc")

    def appendBookmarkedText(self,importer,srcBookmark,dstNode):

        # This is the paragraph that contains the beginning of the bookmark.
        startPara = srcBookmark.getBookmarkStart().getParentNode()

        # This is the paragraph that contains the end of the bookmark.
        endPara = srcBookmark.getBookmarkEnd().getParentNode()

        if startPara is None or endPara is None :
            raise ValueError('Parent of the bookmark start or end is not a paragraph, cannot handle this scenario yet.')

        # Limit ourselves to a reasonably simple scenario.
        if startPara.getParentNode() != endPara.getParentNode() :
            raise ValueError('Start and end paragraphs have different parents, cannot handle this scenario yet.')

        # We want to copy all paragraphs from the start paragraph up to (and including) the end paragraph,
        # therefore the node at which we stop is one after the end paragraph.
        endNode = endPara.getNextSibling()

        # This is the loop to go through all paragraph-level nodes in the bookmark.
        curNode = startPara
        while curNode != endNode :
            # This creates a copy of the current node and imports it (makes it valid) in the context
            # of the destination document. Importing means adjusting styles and list identifiers correctly.
            newNode = importer.importNode(curNode, True)

            # Now we simply append the new node to the destination.
            dstNode.appendChild(newNode)
            curNode = curNode.getNextSibling()

class UntangleRowBookmarks:

    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Document = jpype.JClass("com.aspose.words.Document")
        self.Row = jpype.JClass("com.aspose.words.Row")
        self.NodeImporter = jpype.JClass("com.aspose.words.NodeImporter")
        self.ImportFormatMode = jpype.JClass("com.aspose.words.ImportFormatMode")

    def main(self):

        # Load a document.
        doc = self.Document(self.dataDir + "TestDefect1352.doc")

        # This perform the custom task of putting the row bookmark ends into the same row with the bookmark starts.
        self.untangleRowBookmarks(doc)

        # Now we can easily delete rows by a bookmark without damaging any other row's bookmarks.
        self.deleteRowByBookmark(doc, "ROW2")

        # This is just to check that the other bookmark was not damaged.
        if doc.getRange().getBookmarks().get("ROW1").getBookmarkEnd() is None:
            raise ValueError('Wrong, the end of the bookmark was deleted.')

        # Save the finished document.
        doc.save(self.dataDir + "TestDefect1352 Out.doc")

    def untangleRowBookmarks(self,doc):

        bookmarks = doc.getRange().getBookmarks()
        bookmarks_count = bookmarks.getCount()

        x = 0

        while x < bookmarks_count:

            bookmark = bookmarks.get(x)
            # Get the parent row of both the bookmark and bookmark end node.
            row1 = bookmark.getBookmarkStart().getAncestor(self.Row)
            row2 = bookmark.getBookmarkEnd().getAncestor(self.Row)

            # If both rows are found okay and the bookmark start and end are contained
            # in adjacent rows, then just move the bookmark end node to the end
            # of the last paragraph in the last cell of the top row.
            if row1 is not None and row2 is not None and row1.getNextSibling() == row2:
                row1.getLastCell().getLastParagraph().appendChild(bookmark.getBookmarkEnd())
            x = x + 1

    def deleteRowByBookmark(self,doc,bookmarkName):

        # Find the bookmark in the document. Exit if cannot find it.
        bookmark = doc.getRange().getBookmarks().get(bookmarkName)
        if bookmark is None:
            return

        # Get the parent row of the bookmark. Exit if the bookmark is not in a row.
        row = bookmark.getBookmarkStart().getAncestor(self.Row)
        if row is None:
            return

        # Remove the row.
        row.remove()

class ProcessComments:

    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Document = jpype.JClass("com.aspose.words.Document")
        self.nodeType = jpype.JClass("com.aspose.words.NodeType")
        self.SaveFormat = jpype.JClass("com.aspose.words.SaveFormat")

    def main(self):

        # Open the document.
        doc = self.Document(self.dataDir + "TestFile.doc")

        #ExStart
        #ExId:ProcessComments_Main
        #ExSummary: The demo-code that illustrates the methods for the comments extraction and removal.
        # Extract the information about the comments of all the authors.

        comments = self.extractComments(doc)

        for comment in comments:
            print (comment)

        # Remove comments by the "pm" author.
        self.removeComments(doc, "pm")
        print ("Comments from \"pm\" are removed!")

        # Extract the information about the comments of the "ks" author.
        comments = self.extractComments(doc, "ks")
        for comment in comments:
            print (comment)

        # Remove all comments.
        self.removeComments(doc)
        print ("All comments are removed!")

        # Save the document.
        doc.save(self.dataDir + "Test File Out.doc")
        #ExEnd

    def extractComments(self,*args):

        doc = args[0]
        collectedComments = []
        # Collect all comments in the document
        comments = doc.getChildNodes(self.nodeType.COMMENT, True).toArray()

        # Look through all comments and gather information about them.

        for comment in comments :

            if 1 < len(args) and args[1] is not None :
                authorName = args[1]
                if str(comment.getAuthor()) == authorName:
                    collectedComments.append(str(comment.getAuthor()) + " " + str(comment.getDateTime()) + " " + str(comment.toString()))
            else:
                collectedComments.append(str(comment.getAuthor()) + " " + str(comment.getDateTime()) + " " + str(comment.toString()))

        return collectedComments

    def removeComments(self,*args):

        doc = args[0]
        if 1 < len(args) and args[1] is not None :
                authorName = args[1]
        # Collect all comments in the document

        comments = doc.getChildNodes(self.nodeType.COMMENT, True)
        comments_count = comments.getCount()
        #/ Look through all comments and remove those written by the authorName author.
        i = comments_count
        i = i - 1
        while i >= 0 :
            comment = comments.get(i)

            if 1 < len(args) and args[1] is not None :
                authorName = args[1]
                if str(comment.getAuthor()) == authorName:
                    comment.remove()
            else:
                comment.remove()
            i = i - 1

class ExtractContent:

    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Document = jpype.JClass("com.aspose.words.Document")
        self.DocumentBuilder = jpype.JClass("com.aspose.words.DocumentBuilder")
        self.NodeType = jpype.JClass("com.aspose.words.NodeType")
        self.NodeImporter = jpype.JClass("com.aspose.words.NodeImporter")
        self.ImportFormatMode = jpype.JClass("com.aspose.words.ImportFormatMode")
        self.Collections = jpype.JClass("java.util.Collections")
        self.SaveFormat = jpype.JClass("com.aspose.words.SaveFormat")

    def main(self):

        # Call methods to test extraction of different types from the document.
        self.extractContentBetweenParagraphs()
        self.extractContentBetweenBlockLevelNodes()
        self.extractContentBetweenParagraphStyles()
        self.extractContentBetweenRuns()
        self.extractContentUsingField()
        self.extractContentBetweenBookmark()
        self.extractContentBetweenCommentRange()

    def extractContentBetweenParagraphs(self):

        #ExStart
        #ExId:ExtractBetweenNodes_BetweenParagraphs
        #ExSummary:Shows how to extract the content between specific paragraphs using the ExtractContent method above.
        # Load in the document
        doc = self.Document(self.dataDir + "TestFile.doc")

        # Gather the nodes. The GetChild method uses 0-based index
        startPara = doc.getFirstSection().getChild(self.NodeType.PARAGRAPH, 6, True)
        endPara = doc.getFirstSection().getChild(self.NodeType.PARAGRAPH, 10, True)
        # Extract the content between these nodes in the document. Include these markers in the extraction.
        extractedNodes = self.extractContents(startPara, endPara, True)

        # Insert the content into a new separate document and save it to disk.
        dstDoc = self.generateDocument(doc, extractedNodes)
        dstDoc.save(self.dataDir + "TestFile.Paragraphs Out.doc")
        #ExEnd

    def extractContentBetweenBlockLevelNodes(self):

        #ExStart
        #ExId:ExtractBetweenNodes_BetweenNodes
        #ExSummary:Shows how to extract the content between a paragraph and table using the ExtractContent method.
        # Load in the document
        doc = self.Document(self.dataDir + "TestFile.doc")

        startPara = doc.getLastSection().getChild(self.NodeType.PARAGRAPH, 2, True)
        endTable = doc.getLastSection().getChild(self.NodeType.TABLE, 0, True)

        # Extract the content between these nodes in the document. Include these markers in the extraction.
        extractedNodes = self.extractContents(startPara, endTable, True)

        # Lets reverse the array to make inserting the content back into the document easier.
        Collections = extractedNodes[::-1]

        while (len(extractedNodes) > 0):
            # Insert the last node from the reversed list
            endTable.getParentNode().insertAfter(extractedNodes[0], endTable)
            # Remove this node from the list after insertion.
            del extractedNodes[0]

        # Save the generated document to disk.
        doc.save(self.dataDir + "TestFile.DuplicatedContent Out.doc")
        #ExEnd

    def extractContentBetweenRuns(self):

        #ExStart
        #ExId:ExtractBetweenNodes_BetweenRuns
        #ExSummary:Shows how to extract content between specific runs of the same paragraph using the ExtractContent method.
        # Load in the document
        doc = self.Document(self.dataDir + "TestFile.doc")

        # Retrieve a paragraph from the first section.
        para = doc.getChild(self.NodeType.PARAGRAPH, 7, True)

        # Use some runs for extraction.
        startRun = para.getRuns().get(1)
        endRun = para.getRuns().get(4)

        # Extract the content between these nodes in the document. Include these markers in the extraction.
        extractedNodes = self.extractContents(startRun, endRun, True)

        # Get the node from the list. There should only be one paragraph returned in the list.
        node = extractedNodes[0]
        # Print the text of this node to the console.
        print (node.toString())

    #ExEnd

    def extractContentUsingField(self):

        #ExStart
        #ExFor:DocumentBuilder.MoveToMergeField(String, Boolean, Boolean)
        #ExId:ExtractBetweenNodes_UsingField
        #ExSummary:Shows how to extract content between a specific field and paragraph in the document using the ExtractContent method.
        # Load in the document
        doc = self.Document(self.dataDir + "TestFile.doc")

        # Use a document builder to retrieve the field start of a merge field.
        builder = self.DocumentBuilder(doc)

        # Pass the first boolean parameter to get the DocumentBuilder to move to the FieldStart of the field.
        # We could also get FieldStarts of a field using GetChildNode method as in the other examples.
        builder.moveToMergeField("Fullname", False, False)

        # The builder cursor should be positioned at the start of the field.
        startField = builder.getCurrentNode()
        endPara = doc.getFirstSection().getChild(self.NodeType.PARAGRAPH, 5, True)

        # Extract the content between these nodes in the document. Don't include these markers in the extraction.
        extractedNodes = self.extractContents(startField, endPara, False)

        # Insert the content into a new separate document and save it to disk.
        dstDoc = self.generateDocument(doc, extractedNodes)
        dstDoc.save(self.dataDir + "TestFile.Fields Out.pdf")

    #ExEnd

    def extractContentBetweenBookmark(self):
        
        #ExStart
        #ExId:ExtractBetweenNodes_BetweenBookmark
        #ExSummary:Shows how to extract the content referenced a bookmark using the ExtractContent method.
        # Load in the document
        doc = self.Document(self.dataDir + "TestFile.doc")

        # Retrieve the bookmark from the document.
        bookmark = doc.getRange().getBookmarks().get("Bookmark1")

        # We use the BookmarkStart and BookmarkEnd nodes as markers.
        bookmarkStart = bookmark.getBookmarkStart()
        bookmarkEnd = bookmark.getBookmarkEnd()

        # Firstly extract the content between these nodes including the bookmark.
        extractedNodesInclusive = self.extractContents(bookmarkStart, bookmarkEnd, True)
        dstDoc = self.generateDocument(doc, extractedNodesInclusive)
        dstDoc.save(self.dataDir + "TestFile.BookmarkInclusive Out.doc")

        # Secondly extract the content between these nodes this time without including the bookmark.
        extractedNodesExclusive = self.extractContents(bookmarkStart, bookmarkEnd, False)
        dstDoc = self.generateDocument(doc, extractedNodesExclusive)
        dstDoc.save(self.dataDir + "TestFile.BookmarkExclusive Out.doc")

    #ExEnd

    def extractContentBetweenCommentRange(self):

        #ExStart
        #ExId:ExtractBetweenNodes_BetweenComment
        #ExSummary:Shows how to extract content referenced by a comment using the ExtractContent method.
        # Load in the document
        doc = self.Document(self.dataDir + "TestFile.doc")

        # This is a quick way of getting both comment nodes.
        # Your code should have a proper method of retrieving each corresponding start and end node.
        commentStart = doc.getChild(self.NodeType.COMMENT_RANGE_START, 0, True)
        commentEnd = doc.getChild(self.NodeType.COMMENT_RANGE_END, 0, True)

        # Firstly extract the content between these nodes including the comment as well.
        extractedNodesInclusive = self.extractContents(commentStart, commentEnd, True)
        dstDoc = self.generateDocument(doc, extractedNodesInclusive)
        dstDoc.save(self.dataDir + "TestFile.CommentInclusive Out.doc")

        # Secondly extract the content between these nodes without the comment.
        extractedNodesExclusive = self.extractContents(commentStart, commentEnd, False)
        dstDoc = self.generateDocument(doc, extractedNodesExclusive)
        dstDoc.save(self.dataDir + "TestFile.CommentExclusive Out.doc")
        #ExEnd


    def extractContentBetweenParagraphStyles(self):

        #ExStart
        #ExId:ExtractBetweenNodes_BetweenStyles
        #ExSummary:Shows how to extract content between paragraphs with specific styles using the ExtractContent method.
        # Load in the document
        doc = self.Document(self.dataDir + "TestFile.doc")

        # Gather a list of the paragraphs using the respective heading styles.
        parasStyleHeading1 = self.paragraphsByStyleName(doc, "Heading 1")
        parasStyleHeading3 = self.paragraphsByStyleName(doc, "Heading 3")

        # Use the first instance of the paragraphs with those styles.
        startPara1 = parasStyleHeading1[0]
        endPara1 = parasStyleHeading3[0]


        # Extract the content between these nodes in the document. Don't include these markers in the extraction.
        extractedNodes = self.extractContents(startPara1, endPara1, False)


        # Insert the content into a new separate document and save it to disk.
        dstDoc = self.generateDocument(doc, extractedNodes)
        dstDoc.save(self.dataDir + "TestFile.Styles Out.doc")

    #ExEnd

    def paragraphsByStyleName(self,doc,styleName):

        # Create an array to collect paragraphs of the specified style.
        paragraphsWithStyle = []
        # Get all paragraphs from the document.
        paragraphs = doc.getChildNodes(self.NodeType.PARAGRAPH, True)
        # Look through all paragraphs to find those with the specified style.

        paragraphs_count = paragraphs.getCount()

        i = 0
        while(i < paragraphs_count) :
            paragraph = paragraphs.get(i)
            if (paragraph.getParagraphFormat().getStyle().getName() == styleName):
                paragraphsWithStyle.append(paragraph)
            i = i + 1

        return paragraphsWithStyle



    # ExStart
    # ExId:ExtractBetweenNodes_ExtractContent
    # ExSummary:This is a method which extracts blocks of content from a document between specified nodes.
    #
    # Extracts a range of nodes from a document found between specified markers and returns a copy of those nodes. Content can be extracted
    # between inline nodes, block level nodes, and also special nodes such as Comment or Boomarks. Any combination of different marker types can used.
    #
    # @param startNode The node which defines where to start the extraction from the document. This node can be block or inline level of a body.
    # @param endNode The node which defines where to stop the extraction from the document. This node can be block or inline level of body.
    # @param isInclusive Should the marker nodes be included.
    #

    def extractContents(self,startNode, endNode, isInclusive):

        # First check that the nodes passed to this method are valid for use.
        self.verifyParameterNodes(startNode, endNode)

        # Create a list to store the extracted nodes.
        nodes = []

        # Keep a record of the original nodes passed to this method so we can split marker nodes if needed.
        originalStartNode = startNode
        originalEndNode = endNode

        # Extract content based on block level nodes (paragraphs and tables). Traverse through parent nodes to find them.
        # We will split the content of first and last nodes depending if the marker nodes are inline
        while startNode.getParentNode().getNodeType() != self.NodeType.BODY :
            startNode = startNode.getParentNode()

        while (endNode.getParentNode().getNodeType() != self.NodeType.BODY):
            endNode = endNode.getParentNode()

        print (str(originalStartNode) + " = " + str(startNode))
        print (str(originalEndNode) + " = " + str(endNode))

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
                    self.processMarker(cloneNode, nodes, originalStartNode, isInclusive, isStartingNode, isEndingNode)
                    isStartingNode = False
                # Conditional needs to be separate as the block level start and end markers maybe the same node.
                if (isEndingNode):
                    self.processMarker(cloneNode, nodes, originalEndNode, isInclusive, isStartingNode, isEndingNode)
                    isExtracting = False
            else:
                # Node is not a start or end marker, simply add the copy to the list.
                nodes.append(cloneNode)

            # Move to the next node and extract it. If next node is null that means the rest of the content is found in a different section.
            if (currNode.getNextSibling() is None and isExtracting):
                # Move to the next section.
                nextSection = currNode.getAncestor(self.NodeType.SECTION).getNextSibling()
                currNode = nextSection.getBody().getFirstChild()
            else:
                # Move to the next node in the body.
                currNode = currNode.getNextSibling()


        # Return the nodes between the node markers.
        return nodes
    # ExEnd

    # ExStart
    # ExId:ExtractBetweenNodes_Helpers
    # ExSummary:The helper methods used by the ExtractContent method.
    #
    # Checks the input parameters are correct and can be used. Throws an exception if there is any problem.
    #

    def verifyParameterNodes(self,startNode,endNode):
        
        # The order in which these checks are done is important.
        if (startNode is None):
            raise ValueError('Start node cannot be null')
        if (endNode is None):
            raise ValueError('End node cannot be null')
        if (startNode.getDocument() != endNode.getDocument()):
            raise ValueError('Start node and end node must belong to the same document')
        if (startNode.getAncestor(self.NodeType.BODY) is None or endNode.getAncestor(self.NodeType.BODY) is None):
            raise ValueError('Start node and end node must be a child or descendant of a body')

        # Check the end node is after the start node in the DOM tree
        # First check if they are in different sections, then if they're not check their position in the body of the same section they are in.
        startSection = startNode.getAncestor(self.NodeType.SECTION)
        endSection = endNode.getAncestor(self.NodeType.SECTION)

        startIndex = startSection.getParentNode().indexOf(startSection)
        endIndex = endSection.getParentNode().indexOf(endSection)

        if (startIndex == endIndex):

            if (startSection.getBody().indexOf(startNode) > endSection.getBody().indexOf(endNode)):
                raise ValueError('The end node must be after the start node in the body')

        elif (startIndex > endIndex):
            raise ValueError('The section of end node must be after the section start node')

    def isInline(self,node):

        # Test if the node is desendant of a Paragraph or Table node and also is not a paragraph or a table a paragraph inside a comment class which is decesant of a pararaph is possible.
        return ((node.getAncestor(self.NodeType.PARAGRAPH) is not None or node.getAncestor(self.NodeType.TABLE) is not None) and not(node.getNodeType() == self.NodeType.PARAGRAPH or node.getNodeType() == self.NodeType.TABLE))

    def processMarker(self,cloneNode,nodes,node, isInclusive, isStartMarker, isEndMarker):
        
        # If we are dealing with a block level node just see if it should be included and add it to the list.
        if(not(self.isInline(node))):
            # Don't add the node twice if the markers are the same node
            if(not(isStartMarker and isEndMarker)):
                if (isInclusive):
                    nodes.append(cloneNode)
            return

        # If a marker is a FieldStart node check if it's to be included or not.
        # We assume for simplicity that the FieldStart and FieldEnd appear in the same paragraph.
        if (node.getNodeType() == self.NodeType.FIELD_START):
            # If the marker is a start node and is not be included then skip to the end of the field.
            # If the marker is an end node and it is to be included then move to the end field so the field will not be removed.
            if ((isStartMarker and not(isInclusive)) or (not(isStartMarker) and isInclusive)):
                while ((node.getNextSibling() is not None) and (node.getNodeType() != self.NodeType.FIELD_END)):
                    node = node.getNextSibling()


        # If either marker is part of a comment then to include the comment itself we need to move the pointer forward to the Comment
        # node found after the CommentRangeEnd node.
        if (node.getNodeType() == self.NodeType.COMMENT_RANGE_END):
            while (node.getNextSibling() is not None and node.getNodeType() != self.NodeType.COMMENT):
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

    def generateDocument(self,srcDoc,nodes):

        # Create a blank document.
        dstDoc = self.Document()
        # Remove the first paragraph from the empty document.
        dstDoc.getFirstSection().getBody().removeAllChildren()

        # Import each node from the list into the new document. Keep the original formatting of the node.
        importer = self.NodeImporter(srcDoc, dstDoc, self.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        for node in nodes:
            importNode = importer.importNode(node, True)
            dstDoc.getFirstSection().getBody().appendChild(importNode)

        # Return the generated document.
        return dstDoc

class RemoveBreaks:

    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Document = jpype.JClass("com.aspose.words.Document")
        self.NodeType = jpype.JClass("com.aspose.words.NodeType")
        self.ControlChar = jpype.JClass("com.aspose.words.ControlChar")

    def main(self):

        # Open the document.
        doc = self.Document(self.dataDir + "TestFile.doc")

        # Remove the page and section breaks from the document.
        # In Aspose.Words section breaks are represented as separate Section nodes in the document.
        # To remove these separate sections the sections are combined.
        self.removePageBreaks(doc)
        self.removeSectionBreaks(doc)

        # Save the document.
        doc.save(self.dataDir + "TestFile Out.doc")

    def removePageBreaks(self,doc):

        # Retrieve all paragraphs in the document.
        paragraphs = doc.getChildNodes(self.NodeType.PARAGRAPH, True)

        paragraphs_count = paragraphs.getCount()

        i = 0
        while(i < paragraphs_count) :
            para = paragraphs.get(i)
            if (para.getParagraphFormat().getPageBreakBefore()):
                para.getParagraphFormat().setPageBreakBefore(False)

            runs = para.getRuns().toArray()

            for run in runs:
                if (run.getText() in self.ControlChar.PAGE_BREAK):
                    run.setText(run.getText().replace(self.ControlChar.PAGE_BREAK, ""))

            i = i + 1


    #ExStart
    #ExId:RemoveBreaks_Sections
    #ExSummary:Combines all sections in the document into one.
    def removeSectionBreaks(self,doc):

        # Loop through all sections starting from the section that precedes the last one
        # and moving to the first section.
        i = doc.getSections().getCount() - 2
        while ( i >= 0 ):
            # Copy the content of the current section to the beginning of the last section.
            doc.getLastSection().prependContent(doc.getSections().get(i))
            # Remove the copied section.
            doc.getSections().get(i).remove()
            i = i - 1

class InsertNestedFields:

    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Document = jpype.JClass("com.aspose.words.Document")
        self.DocumentBuilder = jpype.JClass("com.aspose.words.DocumentBuilder")
        self.BreakType = jpype.JClass("com.aspose.words.BreakType")
        self.HeaderFooterType = jpype.JClass("com.aspose.words.HeaderFooterType")


    def main(self):

        doc = self.Document()
        builder = self.DocumentBuilder(doc)

        # Insert few page breaks (just for testing)
        i = 0
        while(i < 5):
            builder.insertBreak(self.BreakType.PAGE_BREAK)
            i = i + 1

        # Move DocumentBuilder cursor into the primary footer.
        builder.moveToHeaderFooter(self.HeaderFooterType.FOOTER_PRIMARY)

        # We want to insert a field like this:
        # { IF {PAGE} <> {NUMPAGES} "See Next Page" "Last Page" }
        field = builder.insertField("IF ")
        builder.moveTo(field.getSeparator())
        builder.insertField("PAGE")
        builder.write(" <> ")
        builder.insertField("NUMPAGES")
        builder.write(" \"See Next Page\" \"Last Page\" ")

        # Finally update the outer field to recalcaluate the final value. Doing this will automatically update
        # the inner fields at the same time.
        field.update()

        doc.save(self.dataDir + "InsertNestedFields Out.docx")

class RemoveField:

    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Document = jpype.JClass("com.aspose.words.Document")

    def main(self):
        
        doc = self.Document(self.dataDir + "Field.RemoveField.doc")

        #ExStart
        #ExFor:Field.Remove
        #ExId:DocumentBuilder_RemoveField
        #ExSummary:Removes a field from the document.
        field = doc.getRange().getFields().get(0)
        # Calling this method completely removes the field from the document.
        field.remove()
    #ExEnd

class AddWatermark:

    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Document = jpype.JClass("com.aspose.words.Document")
        self.Shape = jpype.JClass("com.aspose.words.Shape")
        self.ShapeType = jpype.JClass("com.aspose.words.ShapeType")
        self.Color = jpype.JClass("java.awt.Color")
        self.RelativeHorizontalPosition = jpype.JClass("com.aspose.words.RelativeHorizontalPosition")
        self.RelativeVerticalPosition = jpype.JClass("com.aspose.words.RelativeVerticalPosition")
        self.WrapType = jpype.JClass("com.aspose.words.WrapType")
        self.VerticalAlignment = jpype.JClass("com.aspose.words.VerticalAlignment")
        self.HorizontalAlignment = jpype.JClass("com.aspose.words.HorizontalAlignment")
        self.Paragraph = jpype.JClass("com.aspose.words.Paragraph")
        self.HeaderFooterType = jpype.JClass("com.aspose.words.HeaderFooterType")
        self.HeaderFooter = jpype.JClass("com.aspose.words.HeaderFooter")


    def main(self):

        doc = self.Document(self.dataDir + "TestFile.doc")
        self.insertWatermarkText(doc, "CONFIDENTIAL")
        doc.save(self.dataDir + "TestFile Out.doc")

    def insertWatermarkText(self,doc,watermarkText):

        # Create a watermark shape. This will be a WordArt shape.
        # You are free to try other shape types as watermarks.
        watermark = self.Shape(doc, self.ShapeType.TEXT_PLAIN_TEXT)

        # Set up the text of the watermark.
        watermark.getTextPath().setText(watermarkText)
        watermark.getTextPath().setFontFamily("Arial")
        watermark.setWidth(500)
        watermark.setHeight(100)
        # Text will be directed from the bottom-left to the top-right corner.
        watermark.setRotation(-40)
        # Remove the following two lines if you need a solid black text.
        watermark.getFill().setColor(self.Color.GRAY) # Try LightGray to get more Word-style watermark
        watermark.setStrokeColor(self.Color.GRAY) # Try LightGray to get more Word-style watermark

        # Place the watermark in the page center.
        watermark.setRelativeHorizontalPosition(self.RelativeHorizontalPosition.PAGE)
        watermark.setRelativeVerticalPosition(self.RelativeVerticalPosition.PAGE)
        watermark.setWrapType(self.WrapType.NONE)
        watermark.setVerticalAlignment(self.VerticalAlignment.CENTER)
        watermark.setHorizontalAlignment(self.HorizontalAlignment.CENTER)

        # Create a new paragraph and append the watermark to this paragraph.
        watermarkPara = self.Paragraph(doc)
        watermarkPara.appendChild(watermark)

        # Insert the watermark into all headers of each document section.
        for sect in doc.getSections() :
            # There could be up to three different headers in each section, since we want
            # the watermark to appear on all pages, insert into all headers.
            self.insertWatermarkIntoHeader(watermarkPara, sect, self.HeaderFooterType.HEADER_PRIMARY)
            self.insertWatermarkIntoHeader(watermarkPara, sect, self.HeaderFooterType.HEADER_FIRST)
            self.insertWatermarkIntoHeader(watermarkPara, sect, self.HeaderFooterType.HEADER_EVEN)

    def insertWatermarkIntoHeader(self,watermarkPara,sect,headerType):

        header = sect.getHeadersFooters().getByHeaderFooterType(headerType)

        if (header is None):
            # There is no header of the specified type in the current section, create it.
            header = self.HeaderFooter(sect.getDocument(), headerType)
            sect.getHeadersFooters().add(header)

        # Insert a clone of the watermark into the header.
        header.appendChild(watermarkPara.deepClone(True))

class ExtractContentBasedOnStyles:

    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Document = jpype.JClass("com.aspose.words.Document")
        self.SaveFormat = jpype.JClass("com.aspose.words.SaveFormat")
        self.NodeType = jpype.JClass("com.aspose.words.NodeType")

    def main(self):

        #ExStart
        #ExId:ExtractContentBasedOnStyles_Main
        #ExSummary:Run queries and display results.
        # Open the document.
        doc = self.Document(self.dataDir + "TestFile.doc")

        # Define style names as they are specified in the Word document.
        PARA_STYLE = "Heading 1"
        RUN_STYLE = "Intense Emphasis"

        # Collect paragraphs with defined styles.
        # Show the number of collected paragraphs and display the text of this paragraphs.
        paragraphs = self.paragraphsByStyleName(doc, PARA_STYLE)

        print ("abc = " + str(paragraphs[0]))
        print ("Paragraphs with " + PARA_STYLE + " styles " + str(len(paragraphs)) + ":")

        for paragraph in paragraphs :
            print (str(paragraph.toString()))

        # Collect runs with defined styles.
        # Show the number of collected runs and display the text of this runs.
        runs = self.runsByStyleName(doc, RUN_STYLE)

        print ("Runs with " + RUN_STYLE + " styles " + str(len(runs)) + ":")

        for run in runs :
            print (run.getRange().getText())



        #ExEnd

    def paragraphsByStyleName(self,doc,styleName):

        # Create an array to collect paragraphs of the specified style.
        paragraphsWithStyle = []
        # Get all paragraphs from the document.
        paragraphs = doc.getChildNodes(self.NodeType.PARAGRAPH, True)
        # Look through all paragraphs to find those with the specified style.

        paragraphs_count = paragraphs.getCount()

        i = 0
        while(i < paragraphs_count) :
            paragraph = paragraphs.get(i)
            if (paragraph.getParagraphFormat().getStyle().getName() == styleName):
                paragraphsWithStyle.append(paragraph)
            i = i + 1

        return paragraphsWithStyle

    def runsByStyleName(self,doc,styleName):

        # Create an array to collect runs of the specified style.
        runsWithStyle = []

        runs = doc.getChildNodes(self.NodeType.RUN, True)
        # Look through all runs to find those with the specified style.
        runs = runs.toArray()
        for run in runs :
            if (run.getFont().getStyle().getName() == styleName):
                runsWithStyle.append(run)

        return runsWithStyle

class AutoFitTables:

    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Document = jpype.JClass("com.aspose.words.Document")
        self.AutoFitBehavior = jpype.JClass("com.aspose.words.AutoFitBehavior")
        self.NodeType = jpype.JClass("com.aspose.words.NodeType")

    def main(self):

        # Demonstrate autofitting a table to the window.
        self.autoFitTableToWindow()
        # Demonstrate autofitting a table to its contents.
        self.autoFitTableToContents()
        # Demonstrate autofitting a table to fixed column widths.
        self.autoFitTableToFixedColumnWidths()

    def autoFitTableToWindow(self):

        doc = self.Document(self.dataDir + "TestFile.doc")
        table = doc.getChild(self.NodeType.TABLE, 0, True)

        # Autofit the first table to the page width.
        table.autoFit(self.AutoFitBehavior.AUTO_FIT_TO_WINDOW)

        # Save the document to disk.
        doc.save(self.dataDir + "TestFile.AutoFitToWindow Out.doc");
        #ExEnd

    def autoFitTableToContents(self):

        doc = self.Document(self.dataDir + "TestFile.doc")
        table = doc.getChild(self.NodeType.TABLE, 0, True)

        # Auto fit the table to the cell contents
        table.autoFit(self.AutoFitBehavior.AUTO_FIT_TO_CONTENTS)

        # Save the document to disk.
        doc.save(self.dataDir + "TestFile.AutoFitToContents Out.doc")
        #ExEnd

    def autoFitTableToFixedColumnWidths(self):

        doc = self.Document(self.dataDir + "TestFile.doc")
        table = doc.getChild(self.NodeType.TABLE, 0, True)

        # Disable autofitting on this table.
        table.autoFit(self.AutoFitBehavior.FIXED_COLUMN_WIDTHS)

        # Save the document to disk.
        doc.save(self.dataDir + "TestFile.FixedWidth Out.doc")
        #ExEnd