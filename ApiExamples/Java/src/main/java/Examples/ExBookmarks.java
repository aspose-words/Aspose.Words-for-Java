package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.text.MessageFormat;
import java.util.Iterator;

public class ExBookmarks extends ApiExampleBase {
    //ExStart
    //ExFor:Bookmark
    //ExFor:Bookmark.Name
    //ExFor:Bookmark.Text
    //ExFor:Bookmark.Remove
    //ExFor:Bookmark.BookmarkStart
    //ExFor:Bookmark.BookmarkEnd
    //ExFor:BookmarkStart
    //ExFor:BookmarkStart.#ctor
    //ExFor:BookmarkEnd
    //ExFor:BookmarkEnd.#ctor
    //ExFor:BookmarkStart.Accept(DocumentVisitor)
    //ExFor:BookmarkEnd.Accept(DocumentVisitor)
    //ExFor:BookmarkStart.Bookmark
    //ExFor:BookmarkStart.GetText
    //ExFor:BookmarkStart.Name
    //ExFor:BookmarkEnd.Name
    //ExFor:BookmarkCollection
    //ExFor:BookmarkCollection.Item(Int32)
    //ExFor:BookmarkCollection.Item(String)
    //ExFor:BookmarkCollection.Count
    //ExFor:BookmarkCollection.GetEnumerator
    //ExFor:Range.Bookmarks
    //ExFor:DocumentVisitor.VisitBookmarkStart
    //ExFor:DocumentVisitor.VisitBookmarkEnd
    //ExSummary:Shows how to add bookmarks and update their contents.
    @Test //ExSkip
    public void createUpdateAndPrintBookmarks() throws Exception {
        // Create a document with 3 bookmarks: "MyBookmark 1", "MyBookmark 2", "MyBookmark 3"
        Document doc = createDocumentWithBookmarks();
        BookmarkCollection bookmarks = doc.getRange().getBookmarks();
        Assert.assertEquals(bookmarks.getCount(), 3); //ExSkip
        Assert.assertEquals(bookmarks.get(0).getName(), "MyBookmark 1"); //ExSkip
        Assert.assertEquals(bookmarks.get(1).getText(), "Text content of MyBookmark 2"); //ExSkip

        // Look at initial values of our bookmarks
        printAllBookmarkInfo(bookmarks);

        // Obtain bookmarks from a bookmark collection by index/name and update their values
        bookmarks.get(0).setName("Updated name of " + bookmarks.get(0).getName());
        bookmarks.get("MyBookmark 2").setText("Updated text content of " + bookmarks.get(1).getName());
        // Remove the latest bookmark
        // The bookmarked text is not deleted
        bookmarks.get(2).remove();

        bookmarks = doc.getRange().getBookmarks();
        // Check that we have 2 bookmarks after the latest bookmark was deleted
        Assert.assertEquals(bookmarks.getCount(), 2);
        Assert.assertEquals(bookmarks.get(0).getName(), "Updated name of MyBookmark 1"); //ExSkip
        Assert.assertEquals(bookmarks.get(1).getText(), "Updated text content of MyBookmark 2"); //ExSkip

        // Look at updated values of our bookmarks
        printAllBookmarkInfo(bookmarks);
    }

    /// <summary>
    /// Create a document with bookmarks using the start and end nodes.
    /// </summary>
    private static Document createDocumentWithBookmarks() throws Exception {
        DocumentBuilder builder = new DocumentBuilder();
        Document doc = builder.getDocument();

        // An empty document has just one empty paragraph by default
        Paragraph p = doc.getFirstSection().getBody().getFirstParagraph();

        // Add several bookmarks to the document
        for (int i = 1; i <= 3; i++) {
            String bookmarkName = "MyBookmark " + i;

            p.appendChild(new Run(doc, "Text before bookmark."));

            p.appendChild(new BookmarkStart(doc, bookmarkName));
            p.appendChild(new Run(doc, "Text content of " + bookmarkName));
            p.appendChild(new BookmarkEnd(doc, bookmarkName));

            p.appendChild(new Run(doc, "Text after bookmark.\r\n"));
        }

        return builder.getDocument();
    }

    /// <summary>
    /// Use an iterator and a visitor to print info of every bookmark from within a document.
    /// </summary>
    private static void printAllBookmarkInfo(BookmarkCollection bookmarks) throws Exception {
        // Create a DocumentVisitor
        BookmarkInfoPrinter bookmarkVisitor = new BookmarkInfoPrinter();

        // Get the enumerator from the document's BookmarkCollection and iterate over the bookmarks
        Iterator<Bookmark> enumerator = bookmarks.iterator();

        while (enumerator.hasNext()) {
            Bookmark currentBookmark = enumerator.next();

            // Accept our DocumentVisitor it to print information about our bookmarks
            if (currentBookmark != null) {
                currentBookmark.getBookmarkStart().accept(bookmarkVisitor);
                currentBookmark.getBookmarkEnd().accept(bookmarkVisitor);

                // Prints a blank line
                System.out.println(currentBookmark.getBookmarkStart().getText());
            }
        }
    }

    /// <summary>
    /// Visitor that prints bookmark information to the console.
    /// </summary>
    public static class BookmarkInfoPrinter extends DocumentVisitor {
        public int visitBookmarkStart(BookmarkStart bookmarkStart) throws Exception {
            System.out.println(MessageFormat.format("BookmarkStart name: \"{0}\", Content: \"{1}\"", bookmarkStart.getName(),
                    bookmarkStart.getBookmark().getText()));
            return VisitorAction.CONTINUE;
        }

        public int visitBookmarkEnd(BookmarkEnd bookmarkEnd) {
            System.out.println(MessageFormat.format("BookmarkEnd name: \"{0}\"", bookmarkEnd.getName()));
            return VisitorAction.CONTINUE;
        }
    }
    //ExEnd

    @Test
    public void tableColumnBookmarks() throws Exception {
        //ExStart
        //ExFor:Bookmark.IsColumn
        //ExFor:Bookmark.FirstColumn
        //ExFor:Bookmark.LastColumn
        //ExSummary:Shows how to get information about table column bookmark.
        Document doc = new Document(getMyDir() + "TableColumnBookmark.doc");
        for (Bookmark bookmark : doc.getRange().getBookmarks()) {
            System.out.println(MessageFormat.format("Bookmark: {0}{1}", bookmark.getName(), bookmark.isColumn() ? " (Column)" : ""));
            if (bookmark.isColumn()) {
                Row row = (Row) bookmark.getBookmarkStart().getAncestor(NodeType.ROW);
                if (row != null && bookmark.getFirstColumn() < row.getCells().getCount()) {
                    // Print text from the first and last cells containing in bookmark
                    System.out.println(row.getCells().get(bookmark.getFirstColumn()).getText().trim());
                    System.out.println(row.getCells().get(bookmark.getLastColumn()).getText().trim());
                }
            }
        }
        //ExEnd

        Bookmark firstTableColumnBookmark = doc.getRange().getBookmarks().get("FirstTableColumnBookmark");
        Bookmark secondTableColumnBookmark = doc.getRange().getBookmarks().get("SecondTableColumnBookmark");

        Assert.assertTrue(firstTableColumnBookmark.isColumn());
        Assert.assertEquals(firstTableColumnBookmark.getFirstColumn(), 1);
        Assert.assertEquals(firstTableColumnBookmark.getLastColumn(), 3);

        Assert.assertTrue(secondTableColumnBookmark.isColumn());
        Assert.assertEquals(secondTableColumnBookmark.getFirstColumn(), 0);
        Assert.assertEquals(secondTableColumnBookmark.getLastColumn(), 3);
    }

    @Test
    public void clearBookmarks() throws Exception {
        //ExStart
        //ExFor:BookmarkCollection.Clear
        //ExSummary:Shows how to remove all bookmarks from a document.
        // Open a document with 3 bookmarks: "MyBookmark1", "My_Bookmark2", "MyBookmark3"
        Document doc = new Document(getMyDir() + "Bookmarks.docx");

        // Remove all bookmarks from the document
        // The bookmarked text is not deleted
        doc.getRange().getBookmarks().clear();
        //ExEnd

        // Verify that the bookmarks were removed
        Assert.assertEquals(doc.getRange().getBookmarks().getCount(), 0);
    }

    @Test
    public void removeBookmarkFromBookmarkCollection() throws Exception {
        //ExStart
        //ExFor:BookmarkCollection.Remove(Bookmark)
        //ExFor:BookmarkCollection.Remove(String)
        //ExFor:BookmarkCollection.RemoveAt
        //ExSummary:Shows how to remove bookmarks from a document using different methods.
        // Open a document with 3 bookmarks: "MyBookmark1", "My_Bookmark2", "MyBookmark3"
        Document doc = new Document(getMyDir() + "Bookmarks.docx");

        // Remove a particular bookmark from the document
        Bookmark bookmark = doc.getRange().getBookmarks().get(0);
        doc.getRange().getBookmarks().remove(bookmark);

        // Remove a bookmark by specified name
        doc.getRange().getBookmarks().remove("My_Bookmark2");

        // Remove a bookmark at the specified index
        doc.getRange().getBookmarks().removeAt(0);
        //ExEnd

        // In docx we have additional hidden bookmark "_GoBack"
        // When we check bookmarks count, the result will be 1 instead of 0
        Assert.assertEquals(doc.getRange().getBookmarks().getCount(), 1);
    }

    @Test
    public void replaceBookmarkUnderscoresWithWhitespaces() throws Exception {
        //ExStart
        //ExFor:Bookmark.Name
        //ExSummary:Shows how to replace elements in bookmark name.
        // Open a document with 3 bookmarks: "MyBookmark1", "My_Bookmark2", "MyBookmark3"
        Document doc = new Document(getMyDir() + "Bookmarks.docx");
        Assert.assertEquals(doc.getRange().getBookmarks().get(2).getName(), "MyBookmark3"); //ExSkip

        // MS Word document does not support bookmark names with whitespaces by default
        // If you have document which contains bookmark names with underscores, you can simply replace them to whitespaces
        for (Bookmark bookmark : doc.getRange().getBookmarks()) bookmark.setName(bookmark.getName().replace("_", " "));
        //ExEnd
    }
}
