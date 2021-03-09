package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.apache.commons.collections4.IterableUtils;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.text.MessageFormat;
import java.util.Collections;
import java.util.Iterator;


@Test
public class ExBookmarks extends ApiExampleBase {
    @Test
    public void insert() throws Exception {
        //ExStart
        //ExFor:Bookmark.Name
        //ExSummary:Shows how to insert a bookmark.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // A valid bookmark has a name, a BookmarkStart, and a BookmarkEnd node.
        // Any whitespace in the names of bookmarks will be converted to underscores if we open the saved document with Microsoft Word. 
        // If we highlight the bookmark's name in Microsoft Word via Insert -> Links -> Bookmark, and press "Go To",
        // the cursor will jump to the text enclosed between the BookmarkStart and BookmarkEnd nodes.
        builder.startBookmark("My Bookmark");
        builder.write("Contents of MyBookmark.");
        builder.endBookmark("My Bookmark");

        // Bookmarks are stored in this collection.
        Assert.assertEquals("My Bookmark", doc.getRange().getBookmarks().get(0).getName());

        doc.save(getArtifactsDir() + "Bookmarks.Insert.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Bookmarks.Insert.docx");

        Assert.assertEquals("My Bookmark", doc.getRange().getBookmarks().get(0).getName());
    }

    //ExStart
    //ExFor:Bookmark
    //ExFor:Bookmark.Name
    //ExFor:Bookmark.Text
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
    //ExFor:BookmarkCollection.GetEnumerator
    //ExFor:Range.Bookmarks
    //ExFor:DocumentVisitor.VisitBookmarkStart 
    //ExFor:DocumentVisitor.VisitBookmarkEnd
    //ExSummary:Shows how to add bookmarks and update their contents.
    @Test //ExSkip
    public void createUpdateAndPrintBookmarks() throws Exception {
        // Create a document with three bookmarks, then use a custom document visitor implementation to print their contents.
        Document doc = createDocumentWithBookmarks(3);
        BookmarkCollection bookmarks = doc.getRange().getBookmarks();
        Assert.assertEquals(3, bookmarks.getCount()); //ExSkip

        printAllBookmarkInfo(bookmarks);

        // Bookmarks can be accessed in the bookmark collection by index or name, and their names can be updated.
        bookmarks.get(0).setName("{bookmarks[0].Name}_NewName");
        bookmarks.get("MyBookmark_2").setText("Updated text contents of {bookmarks[1].Name}");

        // Print all bookmarks again to see updated values.
        printAllBookmarkInfo(bookmarks);
    }

    /// <summary>
    /// Create a document with a given number of bookmarks.
    /// </summary>
    private static Document createDocumentWithBookmarks(int numberOfBookmarks) throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= numberOfBookmarks; i++) {
            String bookmarkName = "MyBookmark_" + i;

            builder.write("Text before bookmark.");
            builder.startBookmark(bookmarkName);
            builder.write(MessageFormat.format("Text inside {0}.", bookmarkName));
            builder.endBookmark(bookmarkName);
            builder.writeln("Text after bookmark.");
        }

        return doc;
    }

    /// <summary>
    /// Use an iterator and a visitor to print info of every bookmark in the collection.
    /// </summary>
    private static void printAllBookmarkInfo(BookmarkCollection bookmarks) throws Exception {
        BookmarkInfoPrinter bookmarkVisitor = new BookmarkInfoPrinter();

        // Get each bookmark in the collection to accept a visitor that will print its contents.
        Iterator<Bookmark> enumerator = bookmarks.iterator();

        while (enumerator.hasNext()) {
            Bookmark currentBookmark = enumerator.next();

            if (currentBookmark != null) {
                currentBookmark.getBookmarkStart().accept(bookmarkVisitor);
                currentBookmark.getBookmarkEnd().accept(bookmarkVisitor);

                System.out.println(currentBookmark.getBookmarkStart().getText());
            }
        }
    }

    /// <summary>
    /// Prints contents of every visited bookmark to the console.
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
        //ExSummary:Shows how to get information about table column bookmarks.
        Document doc = new Document(getMyDir() + "Table column bookmarks.doc");
        for (Bookmark bookmark : doc.getRange().getBookmarks()) {
            // If a bookmark encloses columns of a table, it is a table column bookmark, and its IsColumn flag set to true.
            System.out.println(MessageFormat.format("Bookmark: {0}{1}", bookmark.getName(), bookmark.isColumn() ? " (Column)" : ""));
            if (bookmark.isColumn()) {
                Row row = (Row) bookmark.getBookmarkStart().getAncestor(NodeType.ROW);
                if (row != null && bookmark.getFirstColumn() < row.getCells().getCount()) {
                    // Print the contents of the first and last columns enclosed by the bookmark.
                    System.out.println(row.getCells().get(bookmark.getFirstColumn()).getText().trim());
                    System.out.println(row.getCells().get(bookmark.getLastColumn()).getText().trim());
                }
            }
        }
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

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
    public void remove() throws Exception {
        //ExStart
        //ExFor:BookmarkCollection.Clear
        //ExFor:BookmarkCollection.Count
        //ExFor:BookmarkCollection.Remove(Bookmark)
        //ExFor:BookmarkCollection.Remove(String)
        //ExFor:BookmarkCollection.RemoveAt
        //ExFor:Bookmark.Remove
        //ExSummary:Shows how to remove bookmarks from a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert five bookmarks with text inside their boundaries.
        for (int i = 1; i <= 5; i++) {
            String bookmarkName = "MyBookmark_" + i;

            builder.startBookmark(bookmarkName);
            builder.write(MessageFormat.format("Text inside {0}.", bookmarkName));
            builder.endBookmark(bookmarkName);
            builder.insertBreak(BreakType.PARAGRAPH_BREAK);
        }

        // This collection stores bookmarks.
        BookmarkCollection bookmarks = doc.getRange().getBookmarks();

        Assert.assertEquals(5, bookmarks.getCount());

        // There are several ways of removing bookmarks.
        // 1 -  Calling the bookmark's Remove method:
        bookmarks.get("MyBookmark_1").remove();

        Assert.assertFalse(IterableUtils.matchesAny(bookmarks, b -> b.getName() == "MyBookmark_1"));

        // 2 -  Passing the bookmark to the collection's Remove method:
        Bookmark bookmark = doc.getRange().getBookmarks().get(0);
        doc.getRange().getBookmarks().remove(bookmark);

        Assert.assertFalse(IterableUtils.matchesAny(bookmarks, b -> b.getName() == "MyBookmark_2"));

        // 3 -  Removing a bookmark from the collection by name:
        doc.getRange().getBookmarks().remove("MyBookmark_3");

        Assert.assertFalse(IterableUtils.matchesAny(bookmarks, b -> b.getName() == "MyBookmark_3"));

        // 4 -  Removing a bookmark at an index in the bookmark collection:
        doc.getRange().getBookmarks().removeAt(0);

        Assert.assertFalse(IterableUtils.matchesAny(bookmarks, b -> b.getName() == "MyBookmark_4"));

        // We can clear the entire bookmark collection.
        bookmarks.clear();

        // The text that was inside the bookmarks is still present in the document.
        Assert.assertTrue(IterableUtils.size(bookmarks) == 0);
        Assert.assertEquals("Text inside MyBookmark_1.\r" +
                "Text inside MyBookmark_2.\r" +
                "Text inside MyBookmark_3.\r" +
                "Text inside MyBookmark_4.\r" +
                "Text inside MyBookmark_5.", doc.getText().trim());
        //ExEnd
    }
}
