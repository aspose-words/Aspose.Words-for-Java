// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
        // Open a blank document, create a DocumentBuilder, and use the builder to add some text wrapped inside bookmarks
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Note that whitespaces in bookmark names will be converted into underscores when saved to Microsoft Word formats
        // such as .doc and .docx, but will be preserved in other formats like .pdf or .xps
        builder.startBookmark("Bookmark 1");
        builder.writeln("Text inside Bookmark 1.");

        builder.startBookmark("Bookmark 2");
        builder.writeln("Text inside Bookmark 1 and 2.");
        builder.endBookmark("Bookmark 2");

        builder.writeln("Text inside Bookmark 1.");
        builder.endBookmark("Bookmark 1");

        builder.startBookmark("Bookmark 3");
        builder.writeln("Text inside Bookmark 3.");
        builder.endBookmark("Bookmark 3");

        // We can specify outline levels for our bookmarks so that they show up in the table of contents and are indented by an amount
        // of space proportional to the indent level in a SaveOptions object
        // Some pdf/xps readers such as Google Chrome also allow the collapsing of all higher level bookmarks by adjacent lower level bookmarks
        // This feature applies to .pdf or .xps file formats, so only their respective SaveOptions subclasses will support it
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

        // We can remove two elements so that only the outline level designation for "Bookmark 1" is left
        outlineLevels.removeAt(2);
        outlineLevels.remove("Bookmark 2");

        // We have 9 bookmark levels to work with, and bookmark levels are also sorted in ascending order,
        // and get numbered in succession along that order
        // Practically this means that our three levels "1, 5, 9", will be seen as "1, 2, 3" in the output
        outlineLevels.add("Bookmark 2", 5);
        outlineLevels.add("Bookmark 3", 9);

        // Save the document as a .pdf and find links to the bookmarks and their outline levels
        doc.save(getArtifactsDir() + "BookmarksOutlineLevelCollection.BookmarkLevels.pdf", pdfSaveOptions);

        // We can empty this dictionary to remove the contents table
        outlineLevels.clear();
        //ExEnd

                // Bind pdf with Aspose.Pdf
        PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
        bookmarkEditor.BindPdf(getArtifactsDir() + "BookmarksOutlineLevelCollection.BookmarkLevels.pdf");

        // Get all bookmarks from the document
        Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

        Assert.AreEqual(3, bookmarks.Count);

        // Assert that all the bookmarks title are with whitespaces
        Assert.AreEqual("Bookmark 1", bookmarks[0].Title);
        Assert.AreEqual("Bookmark 2", bookmarks[1].Title);
        Assert.AreEqual("Bookmark 3", bookmarks[2].Title);            
            }
}
