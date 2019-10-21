package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.examples.Utils;


public class SimpleAppendDocument {
    private static String gDataDir;

    public static void main(String[] args) throws Exception {
        //ExStart:SimpleAppendDocument
        // The path to the documents directory.
        gDataDir = Utils.getDataDir(SimpleAppendDocument.class);

        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc = new Document(gDataDir + "TestFile.Source.doc");

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        dstDoc.save(gDataDir + "TestFile.SimpleAppendDocument Out.docx");
        //ExEnd:SimpleAppendDocument

        System.out.println("Documents appended successfully.");
    }

}