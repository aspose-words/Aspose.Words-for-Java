package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;

import java.awt.*;


public class DocumentBuilderInsertHyperlink {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderInsertHyperlink.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Please make sure to visit ");

        // Specify font formatting for the hyperlink.
        builder.getFont().setColor(Color.MAGENTA);
        //  builder.getFont().setUnderline();
        //builder.Font.Underline = Underline.Single;
        // Insert the link.
        builder.insertHyperlink("Aspose Website", "http://www.aspose.com", false);

        // Revert to default formatting.
        builder.getFont().clearFormatting();
        builder.write(" for more information.");
        doc.save(dataDir + "output.doc");

    }
}