package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.ParagraphFormat;
import com.aspose.words.examples.Utils;


public class DocumentBuilderSetParagraphFormatting {
    public static void main(String[] args) throws Exception {


        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderSetParagraphFormatting.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        ParagraphFormat paragraphFormat = builder.getParagraphFormat();
        paragraphFormat.setAlignment(ParagraphAlignment.CENTER);
        paragraphFormat.setLeftIndent(50);
        paragraphFormat.setRightIndent(50);
        paragraphFormat.setSpaceAfter(25);
        paragraphFormat.setKeepTogether(true);

        builder.writeln("I'm a very nice formatted paragraph. I'm intended to demonstrate how the left and right indents affect word wrapping.");
        builder.writeln("I'm another nice formatted paragraph. I'm intended to demonstrate how the space after paragraph looks like.");
        doc.save(dataDir + "output.doc");

    }
}
