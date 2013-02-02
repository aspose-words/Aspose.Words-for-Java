//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package Examples;

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.Bookmark;
import org.testng.Assert;
import com.aspose.words.DocumentBuilder;

@Test
public class ExBookmarks extends ExBase
{
    public void bookmarkNameAndText() throws Exception
    {
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
    public void bookmarkRemove() throws Exception
    {
        //ExStart
        //ExFor:Bookmark.Remove
        //ExSummary:Shows how to remove a particular bookmark from a document.
        Document doc = new Document(getMyDir() + "Bookmark.doc");

        // Use the indexer of the Bookmarks collection to obtain the desired bookmark.
        Bookmark bookmark = doc.getRange().getBookmarks().get("MyBookmark");

        // Remove the bookmark. The bookmarked text is not deleted.
        bookmark.remove();
        //ExEnd

        Assert.assertEquals(doc.getRange().getBookmarks().getCount(), 0);
    }

    @Test
    public void ClearBookmarks() throws Exception
    {
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
    public void accessBookmarks() throws Exception
    {
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
    public void bookmarkCollectionRemove() throws Exception
    {
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
    public void bookmarksInsertBookmark() throws Exception
    {
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
    public void GetBookmarkCount() throws Exception
    {
        //ExStart
        //ExFor:BookmarkCollection.Count
        //ExSummary:Shows how to count the number of bookmarks in a document.
        Document doc = new Document(getMyDir() + "Bookmark.doc");

        int count = doc.getRange().getBookmarks().getCount();
        //ExEnd

        Assert.assertEquals(count, 1);
    }
}

