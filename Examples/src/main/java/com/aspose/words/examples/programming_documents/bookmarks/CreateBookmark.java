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
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(CreateBookmark.class);
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("MyBookmark");
        builder.writeln("Text inside a bookmark.");
        builder.endBookmark("MyBookmark");

        builder.startBookmark("Nested Bookmark");
        builder.writeln("Text inside a NestedBookmark.");
        builder.endBookmark("Nested Bookmark");

        builder.writeln("Text after Nested Bookmark.");
        builder.endBookmark("My Bookmark");

        PdfSaveOptions options = new PdfSaveOptions();
        options.getOutlineOptions().setDefaultBookmarksOutlineLevel(1);
        options.getOutlineOptions().setDefaultBookmarksOutlineLevel(2);

        doc.save(dataDir + "output.pdf");
        System.out.println("\nBookmark created successfully.");
    }

}