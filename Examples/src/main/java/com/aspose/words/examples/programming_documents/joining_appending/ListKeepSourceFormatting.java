package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.SectionStart;
import com.aspose.words.examples.Utils;


public class ListKeepSourceFormatting {
    public static void main(String[] args) throws Exception {

        //ExStart:ListKeepSourceFormatting
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ListKeepSourceFormatting.class);

        Document dstDoc = new Document(dataDir + "TestFile.DestinationList.doc");
        Document srcDoc = new Document(dataDir + "TestFile.SourceList.doc");

        // Append the content of the document so it flows continuously.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.CONTINUOUS);

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(dataDir + "output.doc");
        //ExEnd:ListKeepSourceFormatting

    }
}