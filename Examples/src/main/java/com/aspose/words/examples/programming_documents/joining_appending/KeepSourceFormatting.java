package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.examples.Utils;


public class KeepSourceFormatting {

    public static void main(String[] args) throws Exception {

        //ExStart:KeepSourceFormatting
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(KeepSourceFormatting.class);

        Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");
        Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

        // Keep the formatting from the source document when appending it to the destination document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        // Save the joined document to disk.
        dstDoc.save(dataDir + "output.docx");
        //ExEnd:KeepSourceFormatting

    }
}