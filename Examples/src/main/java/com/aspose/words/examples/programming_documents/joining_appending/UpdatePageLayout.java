package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.examples.Utils;


public class UpdatePageLayout {

    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(UpdatePageLayout.class);

        Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");
        Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

        // If the destination document is rendered to PDF, image etc or UpdatePageLayout is called before the source document
        // is appended then any changes made after will not be reflected in the rendered output.
        dstDoc.updatePageLayout();

        // Join the documents.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        // For the changes to be updated to rendered output, UpdatePageLayout must be called again.
        // If not called again the appended document will not appear in the output of the next rendering.
        dstDoc.updatePageLayout();

        // Save the joined document to PDF.
        dstDoc.save(dataDir + "output.pdf");

    }
}