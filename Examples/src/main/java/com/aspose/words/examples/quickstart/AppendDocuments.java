package com.aspose.words.examples.quickstart;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.examples.Utils;

public class AppendDocuments {
    public static void main(String[] args) throws Exception {
        //ExStart:
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AppendDocuments.class);

        // Load the destination and source documents from disk.
        Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");
        Document srcDoc = new Document(dataDir + "TestFile.Source.doc");
        // Append the source document to the destination document while keeping the original formatting of the source document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(dataDir + "TestFile Out.docx");
        //ExEnd:

        System.out.println("Documents appended successfully.");
    }
}