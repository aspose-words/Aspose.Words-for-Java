/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.bookmarks;

import com.aspose.words.Bookmark;
import com.aspose.words.*;
import com.aspose.words.Row;
import  com.aspose.words.SaveFormat.*;
import com.aspose.words.examples.Utils;


public class CreateBookmark
{
    /**
     * The main entry point for the application.
     */
    public static void main(String[] args) throws Exception
    {
        // ExStart:CreateBookmark
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("MyBookmark");
        builder.writeln("Text inside a bookmark.");
        builder.endBookmark("MyBookmark");
       // ExEnd:CreateBookmark
        System.out.println("\nBookmark created successfully.");
    }

}