package com.aspose.words.examples.programming_documents.bookmarks;

import com.aspose.words.Bookmark;
import com.aspose.words.Document;
import com.aspose.words.examples.Utils;


public class ObtainBookmarkByIndexAndName
{
    /**
     * The main entry point for the application.
     */
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ObtainBookmarkByIndexAndName.class);

        Document doc = new Document(dataDir + "Bookmarks.doc");

        // By index.
        Bookmark bookmark1 = doc.getRange().getBookmarks().get(0);
        System.out.println("\nBookmark by index is " + bookmark1.getText());
        // By name.
        Bookmark bookmark2 = doc.getRange().getBookmarks().get("Bookmark2");
        System.out.println("\nBookmark by name is " + bookmark2.getText());
    }
}