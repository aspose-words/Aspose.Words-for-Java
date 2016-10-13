package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.examples.Utils;


public class UnlinkHeadersFooters {

    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(UnlinkHeadersFooters.class);

        Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");
        Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

        // Even a document with no headers or footers can still have the LinkToPrevious setting set to true.
        // Unlink the headers and footers in the source document to stop this from continuing the headers and footers
        // from the destination document.
        srcDoc.getFirstSection().getHeadersFooters().linkToPrevious(false);

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(dataDir + "output.doc");

    }
}