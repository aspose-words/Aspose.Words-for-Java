package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
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
    @Test
    public void bookmarkNameAndText() throws Exception {
        //ExStart
        //ExFor:Bookmark
        //ExFor:Bookmark.Name
        //ExFor:Bookmark.Text
        //ExFor:Range.Bookmarks
        //ExId:BookmarksGetNameSetText
        //ExSummary:Shows how to get or set bookmark name and text.
        Document doc = new Document(getMyDir() + "Bookmark.doc");

        // Use the indexer of the Bookmarks collection to obtain the desired bookmark.
        Bookmark bookmark = doc.getRange().getBookmarks().get("MyBookmark");

        // Get the name and text of the bookmark.
        String name = bookmark.getName();
        String text = bookmark.getText();

        // Set the name and text of the bookmark.
        bookmark.setName("RenamedBookmark");
        bookmark.setText("This is a new bookmarked text.");
        //ExEnd

        Assert.assertEquals(name, "MyBookmark");
        Assert.assertEquals(text, "This is a bookmarked text.");
    }

    @Test
    public void bookmarkRemove() throws Exception {
        //ExStart
        //ExFor:Bookmark.Remove
        //ExSummary:Shows how to remove a particular bookmark from a document.
        Document doc = new Document(getMyDir() + "Bookmark.doc");

        // Use the indexer of the Bookmarks collection to obtain the desired bookmark.
        Bookmark bookmark = doc.getRange().getBookmarks().get("MyBookmark");

        // Remove the bookmark. The bookmarked text is not deleted.
        bookmark.remove();
        //ExEnd

        // Verify that the bookmarks were removed from the document.
        Assert.assertEquals(doc.getRange().getBookmarks().getCount(), 0);
    }

    @Test
    public void clearBookmarks() throws Exception {
        //ExStart
        //ExFor:BookmarkCollection.Clear
        //ExSummary:Shows how to remove all bookmarks from a document.
        Document doc = new Document(getMyDir() + "Bookmark.doc");
        doc.getRange().getBookmarks().clear();
        //ExEnd

        // Verify that the bookmarks were removed from the document.
        Assert.assertEquals(doc.getRange().getBookmarks().getCount(), 0);
    }

    @Test
    public void accessBookmarks() throws Exception {
        //ExStart
        //ExFor:BookmarkCollection
        //ExFor:BookmarkCollection.Item(Int32)
        //ExFor:BookmarkCollection.Item(String)
        //ExId:BookmarksAccess
        //ExSummary:Shows how to obtain bookmarks from a bookmark collection.
        Document doc = new Document(getMyDir() + "Bookmarks.doc");

        // By index.
        Bookmark bookmark1 = doc.getRange().getBookmarks().get(0);

        // By name.
        Bookmark bookmark2 = doc.getRange().getBookmarks().get("Bookmark2");
        //ExEnd
    }

    @Test
    public void bookmarkCollectionRemove() throws Exception {
        //ExStart
        //ExFor:BookmarkCollection.Remove(Bookmark)
        //ExFor:BookmarkCollection.Remove(String)
        //ExFor:BookmarkCollection.RemoveAt
        //ExSummary:Demonstrates different methods of removing bookmarks from a document.
        Document doc = new Document(getMyDir() + "Bookmarks.doc");

        // Remove a particular bookmark from the document.
        Bookmark bookmark = doc.getRange().getBookmarks().get(0);
        doc.getRange().getBookmarks().remove(bookmark);

        // Remove a bookmark by specified name.
        doc.getRange().getBookmarks().remove("Bookmark2");

        // Remove a bookmark at the specified index.
        doc.getRange().getBookmarks().removeAt(0);
        //ExEnd

        Assert.assertEquals(doc.getRange().getBookmarks().getCount(), 0);
    }

    @Test
    public void bookmarksInsertBookmarkWithDocumentBuilder() throws Exception {
        //ExStart
        //ExId:BookmarksInsertBookmark
        //ExSummary:Shows how to create a new bookmark.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("MyBookmark");
        builder.writeln("Text inside a bookmark.");
        builder.endBookmark("MyBookmark");
        //ExEnd
    }

    @Test
    public void getBookmarkCount() throws Exception {
        //ExStart
        //ExFor:BookmarkCollection.Count
        //ExSummary:Shows how to count the number of bookmarks in a document.
        Document doc = new Document(getMyDir() + "Bookmark.doc");

        int count = doc.getRange().getBookmarks().getCount();
        //ExEnd

        Assert.assertEquals(count, 1);
    }

    @Test
    public void createBookmarkWithNodes() throws Exception {
        //ExStart
        //ExFor:BookmarkStart
        //ExFor:BookmarkStart.#ctor
        //ExFor:BookmarkEnd
        //ExFor:BookmarkEnd.#ctor
        //ExSummary:Shows how to create a bookmark by inserting bookmark start and end nodes.
        Document doc = new Document();

        // An empty document has just one empty paragraph by default.
        Paragraph p = doc.getFirstSection().getBody().getFirstParagraph();

        p.appendChild(new Run(doc, "Text before bookmark. "));

        p.appendChild(new BookmarkStart(doc, "My bookmark"));
        p.appendChild(new Run(doc, "Text inside bookmark. "));
        p.appendChild(new BookmarkEnd(doc, "My bookmark"));

        p.appendChild(new Run(doc, "Text after bookmark."));

        doc.save(getArtifactsDir() + "Bookmarks.CreateBookmarkWithNodes.doc");
        //ExEnd

        Assert.assertEquals(doc.getRange().getBookmarks().get("My bookmark").getText(), "Text inside bookmark. ");
    }

    @Test
    public void replaceBookmarkUnderscoresWithWhitespaces() throws Exception {
        //ExStart
        //ExFor:Bookmark.Name
        //ExSummary:Shows how to replace elements in bookmark name
        Document doc = new Document(getMyDir() + "Bookmarks.Replace.docx");

        Assert.assertEquals(doc.getRange().getBookmarks().get(0).getName(), "My_Bookmark"); //ExSkip

        // MS Word document does not support bookmark names with whitespaces by default.
        // If you have document which contains bookmark names with underscores, you can simply replace them to whitespaces.
        for (Bookmark bookmark : doc.getRange().getBookmarks()) {
            bookmark.setName(bookmark.getName().replace("_", " "));
        }
        //ExEnd

        // Check that our replace was correct
        Assert.assertEquals(doc.getRange().getBookmarks().get(0).getName(), "My Bookmark");
    }

    //ExStart
    //ExFor:Bookmark.BookmarkStart
    //ExFor:Bookmark.BookmarkEnd
    //ExFor:BookmarkStart.Accept(DocumentVisitor)
    //ExFor:BookmarkEnd.Accept(DocumentVisitor)
    //ExFor:BookmarkStart.Bookmark
    //ExFor:BookmarkStart.GetText
    //ExFor:BookmarkStart.Name
    //ExFor:BookmarkEnd.Name
    //ExFor:DocumentVisitor.VisitBookmarkStart
    //ExFor:DocumentVisitor.VisitBookmarkEnd
    //ExFor:BookmarkCollection.GetEnumerator
    //ExSummary:Shows how add bookmarks and update their contents.
    @Test //ExSkip
    public void createUpdateAndPrintBookmarks() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some bookmarks to the document
        for (int i = 1; i < 4; i++) {
            String bookmarkName = "Bookmark " + i;

            builder.startBookmark(bookmarkName);
            builder.write("Text content of " + bookmarkName);
            builder.endBookmark(bookmarkName);
        }

        BookmarkCollection bookmarks = doc.getRange().getBookmarks();

        // Look at initial values of our bookmarks
        printAllBookmarkInfo(bookmarks);

        Assert.assertEquals(bookmarks.get(0).getName(), "Bookmark 1"); //ExSkip
        Assert.assertEquals(bookmarks.get(1).getText(), "Text content of Bookmark 2"); //ExSkip
        Assert.assertEquals(bookmarks.getCount(), 3); //ExSkip

        // Update some values
        bookmarks.get(0).setName("Updated name of " + bookmarks.get(0).getName());
        bookmarks.get(1).setText("Updated text content of " + bookmarks.get(1).getName());
        bookmarks.get(2).remove();

        bookmarks = doc.getRange().getBookmarks();

        // Look at updated values of our bookmarks
        printAllBookmarkInfo(bookmarks);

        Assert.assertEquals(bookmarks.get(0).getName(), "Updated name of Bookmark 1"); //ExSkip
        Assert.assertEquals(bookmarks.get(1).getText(), "Updated text content of Bookmark 2"); //ExSkip
        Assert.assertEquals(bookmarks.getCount(), 2); //ExSkip
    }

    /// <summary>
    /// Use an iterator and a visitor to print info of every bookmark from within a document.
    /// </summary>
    private static void printAllBookmarkInfo(final BookmarkCollection bookmarks) throws Exception {
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
        public int visitBookmarkStart(final BookmarkStart bookmarkStart) throws Exception {
            System.out.println(MessageFormat.format("BookmarkStart name: \"{0}\", Content: \"{1}\"", bookmarkStart.getName(),
                    bookmarkStart.getBookmark().getText()));
            return VisitorAction.CONTINUE;
        }

        public int visitBookmarkEnd(final BookmarkEnd bookmarkEnd) {
            System.out.println(MessageFormat.format("BookmarkEnd name: \"{0}\"", bookmarkEnd.getName()));
            return VisitorAction.CONTINUE;
        }
    }
    //ExEnd
}
