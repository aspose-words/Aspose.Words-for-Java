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
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;


public class AccessBookmarks
{
    /**
     * The main entry point for the application.
     */
    public static void main(String[] args) throws Exception
    {
        // ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AccessBookmarks.class);
        Document doc = new Document(dataDir + "Bookmark.doc");
        Bookmark bookmark1 = doc.getRange().getBookmarks().get(0);

        Bookmark bookmark = doc.getRange().getBookmarks().get("MyBookmark");
        dataDir = dataDir + "output.doc";
        doc.save(dataDir);
        // ExEnd:1
        System.out.println("\nTable bookmarked successfully.\nFile saved at " + dataDir);
    }

}