/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.bookmarks;

import com.aspose.words.Bookmark;
import com.aspose.words.Document;
import com.aspose.words.examples.Utils;


public class ObtainBookmarkByIndexAndName
{
    /**
     * The main entry point for the application.
     */
    public static void main(String[] args) throws Exception
    {
        // ExStart:ObtainBookmarkByIndexAndName
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ObtainBookmarkByIndexAndName.class);

        Document doc = new Document(dataDir + "Bookmarks.doc");

        // By index.
        Bookmark bookmark1 = doc.getRange().getBookmarks().get(0);
        System.out.println("\nBookmark by index is " + bookmark1.getText());
        // By name.
        Bookmark bookmark2 = doc.getRange().getBookmarks().get("Bookmark2");
        System.out.println("\nBookmark by name is " + bookmark2.getText());
       // ExEnd:ObtainBookmarkByIndexAndName

    }

}