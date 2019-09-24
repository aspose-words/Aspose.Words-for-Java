package com.aspose.words.examples.programming_documents.bookmarks;

import com.aspose.words.Bookmark;
import com.aspose.words.Document;
import com.aspose.words.examples.Utils;


public class AccessBookmarks {
    /**
     * The main entry point for the application.
     */
    public static void main(String[] args) throws Exception {
        //ExStart:AccessBookmarks
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AccessBookmarks.class);
        Document doc = new Document(dataDir + "Bookmark.doc");
        Bookmark bookmark1 = doc.getRange().getBookmarks().get(0);

        Bookmark bookmark = doc.getRange().getBookmarks().get("MyBookmark");
        doc.save(dataDir + "output.doc");
        System.out.println("\nTable bookmarked successfully.\nFile saved at " + dataDir);

        //ExEnd:AccessBookmarks
    }


}