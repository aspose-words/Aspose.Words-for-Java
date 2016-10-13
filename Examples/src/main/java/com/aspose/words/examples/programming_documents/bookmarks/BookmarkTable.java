package com.aspose.words.examples.programming_documents.bookmarks;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.examples.Utils;


public class BookmarkTable
{
    /**
     * The main entry point for the application.
     */
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(BookmarkTable.class);
        //Create empty document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // We call this method to start building the table.
        builder.startTable();
        builder.insertCell();

        // Start bookmark here after calling InsertCell
        builder.startBookmark("MyBookmark");

        builder.write("Row 1, Cell 1 Content.");

        // Build the second cell
        builder.insertCell();
        builder.write("Row 1, Cell 2 Content.");
        // Call the following method to end the row and start a new row.
        builder.endRow();

        // Build the first cell of the second row.
        builder.insertCell();
        builder.write("Row 2, Cell 1 Content");

        // Build the second cell.
        builder.insertCell();
        builder.write("Row 2, Cell 2 Content.");
        builder.endRow();

        // Signal that we have finished building the table.
        builder.endTable();

        //End of bookmark
        builder.endBookmark("MyBookmark");

        doc.save(dataDir+ "output.doc");
        System.out.println("\nTable bookmarked successfully.\nFile saved at " + dataDir);
    }

}