/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package programmingwithdocuments.workingwithbookmarks.untanglerowbookmarks.java;

import java.io.File;
import java.net.URI;

import com.aspose.words.*;


public class UntangleRowBookmarks
{
    /**
     * The main entry point for the application.
     */
    public static void main(String[] args) throws Exception
    {
        String dataDir = "src/programmingwithdocuments/workingwithbookmarks/untanglerowbookmarks/data/";

        // Load a document.
        Document doc = new Document(dataDir + "TestDefect1352.doc");

        // This perform the custom task of putting the row bookmark ends into the same row with the bookmark starts.
        untangleRowBookmarks(doc);

        // Now we can easily delete rows by a bookmark without damaging any other row's bookmarks.
        deleteRowByBookmark(doc, "ROW2");

        // This is just to check that the other bookmark was not damaged.
        if (doc.getRange().getBookmarks().get("ROW1").getBookmarkEnd() == null)
            throw new Exception("Wrong, the end of the bookmark was deleted.");

        // Save the finished document.
        doc.save(dataDir + "TestDefect1352 Out.doc");
    }

    private static void untangleRowBookmarks(Document doc) throws Exception
    {
        for (Bookmark bookmark : doc.getRange().getBookmarks())
        {
            // Get the parent row of both the bookmark and bookmark end node.
            Row row1 = (Row)bookmark.getBookmarkStart().getAncestor(Row.class);
            Row row2 = (Row)bookmark.getBookmarkEnd().getAncestor(Row.class);

            // If both rows are found okay and the bookmark start and end are contained
            // in adjacent rows, then just move the bookmark end node to the end
            // of the last paragraph in the last cell of the top row.
            if ((row1 != null) && (row2 != null) && (row1.getNextSibling() == row2))
                row1.getLastCell().getLastParagraph().appendChild(bookmark.getBookmarkEnd());
        }
    }

    private static void deleteRowByBookmark(Document doc, String bookmarkName) throws Exception
    {
        // Find the bookmark in the document. Exit if cannot find it.
        Bookmark bookmark = doc.getRange().getBookmarks().get(bookmarkName);
        if (bookmark == null)
            return;

        // Get the parent row of the bookmark. Exit if the bookmark is not in a row.
        Row row = (Row)bookmark.getBookmarkStart().getAncestor(Row.class);
        if (row == null)
            return;

        // Remove the row.
        row.remove();
    }
}