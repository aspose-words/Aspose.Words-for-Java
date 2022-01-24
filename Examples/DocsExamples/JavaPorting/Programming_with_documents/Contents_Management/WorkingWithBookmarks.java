package DocsExamples.Programming_with_Documents.Contents_Management;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.Bookmark;
import com.aspose.words.DocumentBuilder;
import com.aspose.ms.System.msConsole;
import com.aspose.words.NodeType;
import com.aspose.words.Row;
import com.aspose.words.CompositeNode;
import com.aspose.words.NodeImporter;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.Paragraph;
import com.aspose.words.Node;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.Field;
import com.aspose.words.SaveFormat;


class WorkingWithBookmarks extends DocsExamplesBase
{
    @Test
    public void accessBookmarks() throws Exception
    {
        //ExStart:AccessBookmarks
        Document doc = new Document(getMyDir() + "Bookmarks.docx");
        
        // By index:
        Bookmark bookmark1 = doc.getRange().getBookmarks().get(0);
        // By name:
        Bookmark bookmark2 = doc.getRange().getBookmarks().get("MyBookmark3");
        //ExEnd:AccessBookmarks
    }

    @Test
    public void updateBookmarkData() throws Exception
    {
        //ExStart:UpdateBookmarkData
        Document doc = new Document(getMyDir() + "Bookmarks.docx");

        Bookmark bookmark = doc.getRange().getBookmarks().get("MyBookmark1");

        String name = bookmark.getName();
        String text = bookmark.getText();

        bookmark.setName("RenamedBookmark");
        bookmark.setText("This is a new bookmarked text.");
        //ExEnd:UpdateBookmarkData
    }

    @Test
    public void bookmarkTableColumns() throws Exception
    {
        //ExStart:BookmarkTable
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startTable();
        
        builder.insertCell();

        builder.startBookmark("MyBookmark");

        builder.write("This is row 1 cell 1");

        builder.insertCell();
        builder.write("This is row 1 cell 2");

        builder.endRow();

        builder.insertCell();
        builder.writeln("This is row 2 cell 1");

        builder.insertCell();
        builder.writeln("This is row 2 cell 2");

        builder.endRow();
        builder.endTable();
        
        builder.endBookmark("MyBookmark");
        //ExEnd:BookmarkTable

        //ExStart:BookmarkTableColumns
        for (Bookmark bookmark : doc.getRange().getBookmarks())
        {
            System.out.println("Bookmark: {0}{1}",bookmark.getName(),bookmark.isColumn() ? " (Column)" : "");

            if (bookmark.isColumn())
            {
                if (bookmark.getBookmarkStart().getAncestor(NodeType.ROW) instanceof Row row && bookmark.FirstColumn < row.Cells.Count)
                    Console.WriteLine(row.Cells[bookmark.FirstColumn].GetText().TrimEnd(ControlChar.CellChar));
            }
        }
        //ExEnd:BookmarkTableColumns
    }

    @Test
    public void copyBookmarkedText() throws Exception
    {
        Document srcDoc = new Document(getMyDir() + "Bookmarks.docx");

        // This is the bookmark whose content we want to copy.
        Bookmark srcBookmark = srcDoc.getRange().getBookmarks().get("MyBookmark1");

        // We will be adding to this document.
        Document dstDoc = new Document();

        // Let's say we will be appended to the end of the body of the last section.
        CompositeNode dstNode = dstDoc.getLastSection().getBody();

        // If you import multiple times without a single context, it will result in many styles created.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        appendBookmarkedText(importer, srcBookmark, dstNode);
        
        dstDoc.save(getArtifactsDir() + "WorkingWithBookmarks.CopyBookmarkedText.docx");
    }

    /// <summary>
    /// Copies content of the bookmark and adds it to the end of the specified node.
    /// The destination node can be in a different document.
    /// </summary>
    /// <param name="importer">Maintains the import context.</param>
    /// <param name="srcBookmark">The input bookmark.</param>
    /// <param name="dstNode">Must be a node that can contain paragraphs (such as a Story).</param>
    private void appendBookmarkedText(NodeImporter importer, Bookmark srcBookmark, CompositeNode dstNode) throws Exception
    {
        // This is the paragraph that contains the beginning of the bookmark.
        Paragraph startPara = (Paragraph) srcBookmark.getBookmarkStart().getParentNode();

        // This is the paragraph that contains the end of the bookmark.
        Paragraph endPara = (Paragraph) srcBookmark.getBookmarkEnd().getParentNode();

        if (startPara == null || endPara == null)
            throw new IllegalStateException(
                "Parent of the bookmark start or end is not a paragraph, cannot handle this scenario yet.");

        // Limit ourselves to a reasonably simple scenario.
        if (startPara.getParentNode() != endPara.getParentNode())
            throw new IllegalStateException(
                "Start and end paragraphs have different parents, cannot handle this scenario yet.");

        // We want to copy all paragraphs from the start paragraph up to (and including) the end paragraph,
        // therefore the node at which we stop is one after the end paragraph.
        Node endNode = endPara.getNextSibling();

        for (Node curNode = startPara; curNode != endNode; curNode = curNode.getNextSibling())
        {
            // This creates a copy of the current node and imports it (makes it valid) in the context
            // of the destination document. Importing means adjusting styles and list identifiers correctly.
            Node newNode = importer.importNode(curNode, true);

            dstNode.appendChild(newNode);
        }
    }

    @Test
    public void createBookmark() throws Exception
    {
        //ExStart:CreateBookmark
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("My Bookmark");
        builder.writeln("Text inside a bookmark.");

        builder.startBookmark("Nested Bookmark");
        builder.writeln("Text inside a NestedBookmark.");
        builder.endBookmark("Nested Bookmark");

        builder.writeln("Text after Nested Bookmark.");
        builder.endBookmark("My Bookmark");

        PdfSaveOptions options = new PdfSaveOptions();
        options.getOutlineOptions().getBookmarksOutlineLevels().add("My Bookmark", 1);
        options.getOutlineOptions().getBookmarksOutlineLevels().add("Nested Bookmark", 2);

        doc.save(getArtifactsDir() + "WorkingWithBookmarks.CreateBookmark.pdf", options);
        //ExEnd:CreateBookmark
    }

    @Test
    public void showHideBookmarks() throws Exception
    {
        //ExStart:ShowHideBookmarks
        Document doc = new Document(getMyDir() + "Bookmarks.docx");

        showHideBookmarkedContent(doc, "MyBookmark1", false);
        
        doc.save(getArtifactsDir() + "WorkingWithBookmarks.ShowHideBookmarks.docx");
        //ExEnd:ShowHideBookmarks
    }

    //ExStart:ShowHideBookmarkedContent
    public void showHideBookmarkedContent(Document doc, String bookmarkName, boolean showHide) throws Exception
    {
        Bookmark bm = doc.getRange().getBookmarks().get(bookmarkName);

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveToDocumentEnd();

        // {IF "{MERGEFIELD bookmark}" = "true" "" ""}
        Field field = builder.insertField("IF \"", null);
        builder.moveTo(field.getStart().getNextSibling());
        builder.insertField("MERGEFIELD " + bookmarkName + "", null);
        builder.write("\" = \"true\" ");
        builder.write("\"");
        builder.write("\"");
        builder.write(" \"\"");

        Node currentNode = field.getStart();
        boolean flag = true;
        while (currentNode != null && flag)
        {
            if (currentNode.getNodeType() == NodeType.RUN)
                if ("\"".equals(currentNode.toString(SaveFormat.TEXT).trim()))
                    flag = false;

            Node nextNode = currentNode.getNextSibling();

            bm.getBookmarkStart().getParentNode().insertBefore(currentNode, bm.getBookmarkStart());
            currentNode = nextNode;
        }

        Node endNode = bm.getBookmarkEnd();
        flag = true;
        while (currentNode != null && flag)
        {
            if (currentNode.getNodeType() == NodeType.FIELD_END)
                flag = false;

            Node nextNode = currentNode.getNextSibling();

            bm.getBookmarkEnd().getParentNode().insertAfter(currentNode, endNode);
            endNode = currentNode;
            currentNode = nextNode;
        }

        doc.getMailMerge().execute(new String[] { bookmarkName }, new Object[] { showHide });
    }
    //ExEnd:ShowHideBookmarkedContent

    @Test
    public void untangleRowBookmarks() throws Exception
    {
        Document doc = new Document(getMyDir() + "Table column bookmarks.docx");

        // This performs the custom task of putting the row bookmark ends into the same row with the bookmark starts.
        untangle(doc);

        // Now we can easily delete rows by a bookmark without damaging any other row's bookmarks.
        deleteRowByBookmark(doc, "ROW2");

        // This is just to check that the other bookmark was not damaged.
        if (doc.getRange().getBookmarks().get("ROW1").getBookmarkEnd() == null)
            throw new Exception("Wrong, the end of the bookmark was deleted.");

        doc.save(getArtifactsDir() + "WorkingWithBookmarks.UntangleRowBookmarks.docx");
    }

    private void untangle(Document doc) throws Exception
    {
        for (Bookmark bookmark : doc.getRange().getBookmarks())
        {
            // Get the parent row of both the bookmark and bookmark end node.
            Row row1 = (Row) bookmark.getBookmarkStart().getAncestor(Row.class);
            Row row2 = (Row) bookmark.getBookmarkEnd().getAncestor(Row.class);

            // If both rows are found okay, and the bookmark start and end are contained in adjacent rows,
            // move the bookmark end node to the end of the last paragraph in the top row's last cell.
            if (row1 != null && row2 != null && row1.getNextSibling() == row2)
                row1.getLastCell().getLastParagraph().appendChild(bookmark.getBookmarkEnd());
        }
    }

    private void deleteRowByBookmark(Document doc, String bookmarkName)
    {
        Bookmark bookmark = doc.getRange().getBookmarks().get(bookmarkName);

        Row row = (Row) bookmark?.BookmarkStart.GetAncestor(typeof(Row));
        row?.Remove();
    }
}

