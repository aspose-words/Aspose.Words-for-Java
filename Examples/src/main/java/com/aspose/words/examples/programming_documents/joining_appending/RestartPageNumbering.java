package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.SectionStart;
import com.aspose.words.examples.Utils;


public class RestartPageNumbering {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RestartPageNumbering.class);

        Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");
        Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

        // Set the appended document to appear on the next page.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.NEW_PAGE);
        // Restart the page numbering for the document to be appended.
        srcDoc.getFirstSection().getPageSetup().setRestartPageNumbering(true);

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(dataDir + "output.doc");

    }
}