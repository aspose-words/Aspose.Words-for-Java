package com.aspose.words.examples.featurescomparison.bookmarks.deletebookmark;

import com.aspose.words.Bookmark;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

public class AsposeBookmarksDelete
{
    // See more @ http://www.aspose.com/docs/display/wordsjava/Bookmarks+in+Aspose.Words

    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeBookmarksDelete.class);

        Document doc = new Document(dataDir + "Aspose_Bookmark.doc");

        // By name.
        Bookmark bookmark = doc.getRange().getBookmarks().get("AsposeBookmark");
        bookmark.remove();

        doc.save(dataDir + "Aspose_BookmarkDeleted.doc", SaveFormat.DOC);
        System.out.println("Done.");
    }
}
