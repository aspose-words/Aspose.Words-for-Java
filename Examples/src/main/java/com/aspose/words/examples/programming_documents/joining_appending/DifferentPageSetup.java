package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class DifferentPageSetup {

    public static void main(String[] args) throws Exception {

        //ExStart:DifferentPageSetup
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DifferentPageSetup.class);
        String fileName = "TestFile.DestinationList.doc";

        Document dstDoc = new Document(dataDir + fileName);
        Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

        // Set the source document to appear straight after the destination document's content.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.CONTINUOUS);

        // Restart the page numbering on the start of the source document.
        srcDoc.getFirstSection().getPageSetup().setRestartPageNumbering(true);
        srcDoc.getFirstSection().getPageSetup().setPageStartingNumber(1);

        // To ensure this does not happen when the source document has different page setup settings make sure the Settings are
        // identical between the last section of the destination document. If there are further continuous sections tha
        // follow on in the source document then this will need to be Repeated for those sections as well.
        srcDoc.getFirstSection().getPageSetup().setPageWidth(dstDoc.getLastSection().getPageSetup().getPageWidth());
        srcDoc.getFirstSection().getPageSetup().setPageHeight(dstDoc.getLastSection().getPageSetup().getPageHeight());
        srcDoc.getFirstSection().getPageSetup().setOrientation(dstDoc.getLastSection().getPageSetup().getOrientation());

        // Iterate through all sections in the source document.
        for (Paragraph para : (Iterable<Paragraph>) srcDoc.getChildNodes(NodeType.PARAGRAPH, true))
        {
            para.getParagraphFormat().setKeepWithNext(true);
        }

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dataDir = dataDir + Utils.GetOutputFilePath(fileName);
        dstDoc.save(dataDir);
        //ExEnd:DifferentPageSetup
    }
}