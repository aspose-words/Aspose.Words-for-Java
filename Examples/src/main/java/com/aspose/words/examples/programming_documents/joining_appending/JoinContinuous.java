package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.SectionStart;
import com.aspose.words.examples.Utils;


public class JoinContinuous {
    private static String gDataDir;

    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        gDataDir = Utils.getDataDir(JoinContinuous.class);

        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc = new Document(gDataDir + "TestFile.Source.doc");

        // Make the document appear straight after the destination documents content.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.CONTINUOUS);

        // Append the source document using the original styles found in the source document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(gDataDir + "output.doc");

    }
}