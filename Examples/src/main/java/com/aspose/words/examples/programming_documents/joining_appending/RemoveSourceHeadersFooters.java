package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.Section;
import com.aspose.words.examples.Utils;


public class RemoveSourceHeadersFooters {

    public static void main(String[] args) throws Exception {

        //ExStart:RemoveSourceHeadersFooters
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RemoveSourceHeadersFooters.class);

        Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");
        Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

        // Remove the headers and footers from each of the sections in the source document.
        for (Section section : srcDoc.getSections()) {
            section.clearHeadersFooters();
        }

        // Even after the headers and footers are cleared from the source document, the "LinkToPrevious" setting
        // for HeadersFooters can still be set. This will cause the headers and footers to continue from the destination
        // document. This should set to false to avoid this behaviour.
        srcDoc.getFirstSection().getHeadersFooters().linkToPrevious(false);

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(dataDir + "output.doc");
        //ExEnd:RemoveSourceHeadersFooters

    }
}