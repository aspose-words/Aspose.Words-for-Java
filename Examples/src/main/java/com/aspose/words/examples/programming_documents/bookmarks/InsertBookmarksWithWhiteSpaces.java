package com.aspose.words.examples.programming_documents.bookmarks;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.examples.Utils;


public class InsertBookmarksWithWhiteSpaces {
    /**
     * The main entry point for the application.
     */
    public static void main(String[] args) throws Exception {

        //ExStart:InsertBookmarksWithWhiteSpaces
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(InsertBookmarksWithWhiteSpaces.class);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("My Bookmark");
        builder.writeln("Text inside a bookmark.");

        builder.startBookmark("Nested Bookmark");
        builder.writeln("Text inside a NestedBookmark.");
        builder.endBookmark("Nested Bookmark");

        builder.writeln("Text after Nested Bookmark.");
        builder.endBookmark("My Bookmark");


        PdfSaveOptions options = new PdfSaveOptions();
        options.getOutlineOptions().getBookmarksOutlineLevels().add("My Bookmark", 1);
        options.getOutlineOptions().getBookmarksOutlineLevels().add("Nested Bookmark", 2);

        dataDir = dataDir + "Insert.Bookmarks_out_.pdf";
        doc.save(dataDir, options);

        System.out.println("\nBookmarks with white spaces inserted successfully.\nFile saved at " + dataDir);
        //ExEnd:InsertBookmarksWithWhiteSpaces
    }
}