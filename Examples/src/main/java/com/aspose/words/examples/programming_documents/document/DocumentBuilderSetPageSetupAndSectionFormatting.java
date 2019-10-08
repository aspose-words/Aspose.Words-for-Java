package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Orientation;
import com.aspose.words.PaperSize;
import com.aspose.words.examples.Utils;


public class DocumentBuilderSetPageSetupAndSectionFormatting {
    public static void main(String[] args) throws Exception {

        //ExStart:DocumentBuilderSetPageSetupAndSectionFormatting
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderSetPageSetupAndSectionFormatting.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getPageSetup().setOrientation(Orientation.LANDSCAPE);
        builder.getPageSetup().setLeftMargin(50);
        builder.getPageSetup().setPaperSize(PaperSize.PAPER_10_X_14);

        doc.save(dataDir + "output.doc");
        //ExEnd:DocumentBuilderSetPageSetupAndSectionFormatting

    }
}