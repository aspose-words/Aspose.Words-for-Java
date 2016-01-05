from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import NodeImporter
from com.aspose.words import ImportFormatMode
from com.aspose.words import Row

class Bookmarks:

    def __init__(self):
        self.dataDir = Settings.dataDir + 'programming_documents/'
        
        self.copy_bookmarked_text()
        
        # Load a document.
        doc = Document(self.dataDir + "TestDefect1352.doc")

        # This perform the custom task of putting the row bookmark ends into the same row with the bookmark starts.
        self.untangle_row_bookmarks(doc)

        # Now we can easily delete rows by a bookmark without damaging any other row's bookmarks.
        self.delete_row_by_bookmark(doc, "ROW2")

        # This is just to check that the other bookmark was not damaged.
        if doc.getRange().getBookmarks().get("ROW1").getBookmarkEnd() is None:
            raise ValueError('Wrong, the end of the bookmark was deleted.')

        # Save the finished document.
        doc.save(self.dataDir + "TestDefect1352 Out.doc")
    
    def copy_bookmarked_text(self):

        # Load the source document.
        srcDoc = Document(self.dataDir + "Template.doc")
        # This is the bookmark whose content we want to copy.
        srcBookmark = srcDoc.getRange().getBookmarks().get("ntf010145060")
        # We will be adding to this document.
        dstDoc = Document()
        # Let's say we will be appending to the end of the body of the last section.
        dstNode = dstDoc.getLastSection().getBody()
        # It is a good idea to use this import context object because multiple nodes are being imported.
        # If you import multiple times without a single context, it will result in many styles created.
        importer = NodeImporter(srcDoc, dstDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING)
        # Do it once.
        self.append_bookmarked_text(importer, srcBookmark, dstNode)
        # Do it one more time for fun.
        self.append_bookmarked_text(importer, srcBookmark, dstNode)
        # Save the finished document.
        dstDoc.save(self.dataDir + "Template Out.doc")

    def append_bookmarked_text(self, importer, srcBookmark, dstNode):

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

    def untangle_row_bookmarks(self, doc):
        bookmarks = doc.getRange().getBookmarks()
        bookmarks_count = bookmarks.getCount()

        x = 0

        while x < bookmarks_count:

            bookmark = bookmarks.get(x)
            # Get the parent row of both the bookmark and bookmark end node.
            row1 = bookmark.getBookmarkStart().getAncestor(Row)
            row2 = bookmark.getBookmarkEnd().getAncestor(Row)

            # If both rows are found okay and the bookmark start and end are contained
            # in adjacent rows, then just move the bookmark end node to the end
            # of the last paragraph in the last cell of the top row.
            if row1 is not None and row2 is not None and row1.getNextSibling() == row2:
                row1.getLastCell().getLastParagraph().appendChild(bookmark.getBookmarkEnd())
            x = x + 1

    def delete_row_by_bookmark(self, doc, bookmarkName):

        # Find the bookmark in the document. Exit if cannot find it.
        bookmark = doc.getRange().getBookmarks().get(bookmarkName)
        if bookmark is None:
            return

        # Get the parent row of the bookmark. Exit if the bookmark is not in a row.
        row = bookmark.getBookmarkStart().getAncestor(Row)
        if row is None:
            return

        # Remove the row.
        row.remove()

if __name__ == '__main__':        
    Bookmarks()