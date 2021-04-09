// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import org.testng.Assert;
import com.aspose.words.BookmarkCollection;
import java.util.Iterator;
import com.aspose.words.Bookmark;
import com.aspose.ms.System.msConsole;
import com.aspose.words.DocumentVisitor;
import com.aspose.words.VisitorAction;
import com.aspose.words.BookmarkStart;
import com.aspose.words.BookmarkEnd;
import com.aspose.words.NodeType;
import com.aspose.words.Row;
import com.aspose.words.ControlChar;
import com.aspose.words.BreakType;


@Test
public class ExBookmarks extends ApiExampleBase
{
    @Test
    public void insert() throws Exception
    {
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
    public void createUpdateAndPrintBookmarks() throws Exception
    {
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
    private static Document createDocumentWithBookmarks(int numberOfBookmarks) throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= numberOfBookmarks; i++)
        {
            String bookmarkName = "MyBookmark_" + i;

            builder.write("Text before bookmark.");
            builder.startBookmark(bookmarkName);
            builder.write($"Text inside {bookmarkName}.");
            builder.endBookmark(bookmarkName);
            builder.writeln("Text after bookmark.");
        }

        return doc;
    }

    /// <summary>
    /// Use an iterator and a visitor to print info of every bookmark in the collection.
    /// </summary>
    private static void printAllBookmarkInfo(BookmarkCollection bookmarks) throws Exception
    {
        BookmarkInfoPrinter bookmarkVisitor = new BookmarkInfoPrinter();

        // Get each bookmark in the collection to accept a visitor that will print its contents.
        Iterator<Bookmark> enumerator = bookmarks.iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
            {
                Bookmark currentBookmark = enumerator.next();

                if (currentBookmark != null)
                {
                    currentBookmark.getBookmarkStart().accept(bookmarkVisitor);
                    currentBookmark.getBookmarkEnd().accept(bookmarkVisitor);

                    System.out.println(currentBookmark.getBookmarkStart().getText());
                }
            }
        }
        finally { if (enumerator != null) enumerator.close(); }
    }

    /// <summary>
    /// Prints contents of every visited bookmark to the console.
    /// </summary>
    public static class BookmarkInfoPrinter extends DocumentVisitor
    {
        public /*override*/ /*VisitorAction*/int visitBookmarkStart(BookmarkStart bookmarkStart)
        {
            System.out.println("BookmarkStart name: \"{bookmarkStart.Name}\", Contents: \"{bookmarkStart.Bookmark.Text}\"");
            return VisitorAction.CONTINUE;
        }

        public /*override*/ /*VisitorAction*/int visitBookmarkEnd(BookmarkEnd bookmarkEnd)
        {
            System.out.println("BookmarkEnd name: \"{bookmarkEnd.Name}\"");
            return VisitorAction.CONTINUE;
        }
    }
    //ExEnd

    @Test
    public void tableColumnBookmarks() throws Exception
    {
        //ExStart
        //ExFor:Bookmark.IsColumn
        //ExFor:Bookmark.FirstColumn
        //ExFor:Bookmark.LastColumn
        //ExSummary:Shows how to get information about table column bookmarks.
        Document doc = new Document(getMyDir() + "Table column bookmarks.doc");

        for (Bookmark bookmark : doc.getRange().getBookmarks())
        {
            // If a bookmark encloses columns of a table, it is a table column bookmark, and its IsColumn flag set to true.
            msConsole.WriteLine($"Bookmark: {bookmark.Name}{(bookmark.IsColumn ? " (Column)" : "")}");
            if (bookmark.isColumn())
            {
                if (bookmark.getBookmarkStart().getAncestor(NodeType.ROW) instanceof Row row &&
                    bookmark.FirstColumn < row.Cells.Count)
                {
                    // Print the contents of the first and last columns enclosed by the bookmark.
                    msConsole.WriteLine(row.Cells[bookmark.getFirstColumn()].GetText().TrimEnd(ControlChar.CELL_CHAR));
                    msConsole.WriteLine(row.Cells[bookmark.getLastColumn()].GetText().TrimEnd(ControlChar.CELL_CHAR));
                }
            }
        }
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        Bookmark firstTableColumnBookmark = doc.getRange().getBookmarks().get("FirstTableColumnBookmark");
        Bookmark secondTableColumnBookmark = doc.getRange().getBookmarks().get("SecondTableColumnBookmark");

        Assert.assertTrue(firstTableColumnBookmark.isColumn());
        Assert.assertEquals(1, firstTableColumnBookmark.getFirstColumn());
        Assert.assertEquals(3, firstTableColumnBookmark.getLastColumn());

        Assert.assertTrue(secondTableColumnBookmark.isColumn());
        Assert.assertEquals(0, secondTableColumnBookmark.getFirstColumn());
        Assert.assertEquals(3, secondTableColumnBookmark.getLastColumn());
    }

    @Test
    public void remove() throws Exception
    {
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
        for (int i = 1; i <= 5; i++)
        {
            String bookmarkName = "MyBookmark_" + i;

            builder.startBookmark(bookmarkName);
            builder.write($"Text inside {bookmarkName}.");
            builder.endBookmark(bookmarkName);
            builder.insertBreak(BreakType.PARAGRAPH_BREAK);
        }

        // This collection stores bookmarks.
        BookmarkCollection bookmarks = doc.getRange().getBookmarks();

        Assert.assertEquals(5, bookmarks.getCount());

        // There are several ways of removing bookmarks.
        // 1 -  Calling the bookmark's Remove method:
        bookmarks.get("MyBookmark_1").remove();

        Assert.False(bookmarks.Any(b => b.Name == "MyBookmark_1"));

        // 2 -  Passing the bookmark to the collection's Remove method:
        Bookmark bookmark = doc.getRange().getBookmarks().get(0);
        doc.getRange().getBookmarks().remove(bookmark);

        Assert.False(bookmarks.Any(b => b.Name == "MyBookmark_2"));
        
        // 3 -  Removing a bookmark from the collection by name:
        doc.getRange().getBookmarks().remove("MyBookmark_3");

        Assert.False(bookmarks.Any(b => b.Name == "MyBookmark_3"));

        // 4 -  Removing a bookmark at an index in the bookmark collection:
        doc.getRange().getBookmarks().removeAt(0);

        Assert.False(bookmarks.Any(b => b.Name == "MyBookmark_4"));

        // We can clear the entire bookmark collection.
        bookmarks.clear();

        // The text that was inside the bookmarks is still present in the document.
        Assert.That(bookmarks, Is.Empty);
        Assert.assertEquals("Text inside MyBookmark_1.\r" +
                        "Text inside MyBookmark_2.\r" +
                        "Text inside MyBookmark_3.\r" +
                        "Text inside MyBookmark_4.\r" +
                        "Text inside MyBookmark_5.", doc.getText().trim());
        //ExEnd
    }
}
