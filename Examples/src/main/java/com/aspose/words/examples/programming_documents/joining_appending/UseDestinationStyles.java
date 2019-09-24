package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.examples.Utils;


public class UseDestinationStyles {
    private static String gDataDir;

    public static void main(String[] args) throws Exception {

        //ExStart:UseDestinationStyles
        // The path to the documents directory.
        gDataDir = Utils.getDataDir(UseDestinationStyles.class);

        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc = new Document(gDataDir + "TestFile.Source.doc");

        // Append the source document using the styles of the destination document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES);

        // Save the joined document to disk.
        dstDoc.save(gDataDir + "TestFile.UseDestinationStyles Out.doc");
        //ExEnd:UseDestinationStyles


        System.out.println("Documents appended successfully.");
    }
}