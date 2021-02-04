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
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.BookmarksOutlineLevelCollection;
import org.testng.Assert;


@Test
public class ExBookmarksOutlineLevelCollection extends ApiExampleBase
{
    @Test
    public void bookmarkLevels() throws Exception
    {
        //ExStart
        //ExFor:BookmarksOutlineLevelCollection
        //ExFor:BookmarksOutlineLevelCollection.Add(String, Int32)
        //ExFor:BookmarksOutlineLevelCollection.Clear
        //ExFor:BookmarksOutlineLevelCollection.Contains(System.String)
        //ExFor:BookmarksOutlineLevelCollection.Count
        //ExFor:BookmarksOutlineLevelCollection.IndexOfKey(System.String)
        //ExFor:BookmarksOutlineLevelCollection.Item(System.Int32)
        //ExFor:BookmarksOutlineLevelCollection.Item(System.String)
        //ExFor:BookmarksOutlineLevelCollection.Remove(System.String)
        //ExFor:BookmarksOutlineLevelCollection.RemoveAt(System.Int32)
        //ExFor:OutlineOptions.BookmarksOutlineLevels
        //ExSummary:Shows how to set outline levels for bookmarks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bookmark with another bookmark nested inside it.
        builder.startBookmark("Bookmark 1");
        builder.writeln("Text inside Bookmark 1.");

        builder.startBookmark("Bookmark 2");
        builder.writeln("Text inside Bookmark 1 and 2.");
        builder.endBookmark("Bookmark 2");

        builder.writeln("Text inside Bookmark 1.");
        builder.endBookmark("Bookmark 1");

        // Insert another bookmark.
        builder.startBookmark("Bookmark 3");
        builder.writeln("Text inside Bookmark 3.");
        builder.endBookmark("Bookmark 3");

        // When saving to .pdf, bookmarks can be accessed via a drop-down menu and used as anchors by most readers.
        // Bookmarks can also have numeric values for outline levels,
        // enabling lower level outline entries to hide higher-level child entries when collapsed in the reader.
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        BookmarksOutlineLevelCollection outlineLevels = pdfSaveOptions.getOutlineOptions().getBookmarksOutlineLevels();

        outlineLevels.add("Bookmark 1", 1);
        outlineLevels.add("Bookmark 2", 2);
        outlineLevels.add("Bookmark 3", 3);

        Assert.assertEquals(3, outlineLevels.getCount());
        Assert.assertTrue(outlineLevels.contains("Bookmark 1"));
        Assert.assertEquals(1, outlineLevels.get(0));
        Assert.assertEquals(2, outlineLevels.get("Bookmark 2"));
        Assert.assertEquals(2, outlineLevels.indexOfKey("Bookmark 3"));

        // We can remove two elements so that only the outline level designation for "Bookmark 1" is left.
        outlineLevels.removeAt(2);
        outlineLevels.remove("Bookmark 2");

        // There are nine outline levels. Their numbering will be optimized during the save operation.
        // In this case, levels "5" and "9" will become "2" and "3".
        outlineLevels.add("Bookmark 2", 5);
        outlineLevels.add("Bookmark 3", 9);

        doc.save(getArtifactsDir() + "BookmarksOutlineLevelCollection.BookmarkLevels.pdf", pdfSaveOptions);

        // Emptying this collection will preserve the bookmarks and put them all on the same outline level.
        outlineLevels.clear();
        //ExEnd

                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
        bookmarkEditor.BindPdf(getArtifactsDir() + "BookmarksOutlineLevelCollection.BookmarkLevels.pdf");

        Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

        Assert.AreEqual(3, bookmarks.Count);
        Assert.AreEqual("Bookmark 1", bookmarks[0].Title);
        Assert.AreEqual("Bookmark 2", bookmarks[1].Title);
        Assert.AreEqual("Bookmark 3", bookmarks[2].Title);            
            }
}
