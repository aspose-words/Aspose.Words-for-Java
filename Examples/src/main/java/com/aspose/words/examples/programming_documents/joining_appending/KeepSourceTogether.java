package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class KeepSourceTogether {

    public static void main(String[] args) throws Exception {

        //ExStart:KeepSourceTogether
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(KeepSourceTogether.class);

        Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");
        Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

        // Set the source document to appear straight after the destination document's content.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.CONTINUOUS);

        // Iterate through all sections in the source document.
        for (Paragraph para : (Iterable<Paragraph>) srcDoc.getChildNodes(NodeType.PARAGRAPH, true)) {
            para.getParagraphFormat().setKeepWithNext(true);
        }

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(dataDir + "output.doc");
        //ExEnd:KeepSourceTogether

    }
}