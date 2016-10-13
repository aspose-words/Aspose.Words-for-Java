
package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.*;
import com.aspose.words.Font;
import com.aspose.words.examples.Utils;

import java.awt.*;


public class DocumentBuilderInsertParagraph {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderInsertParagraph.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Font font = builder.getFont();
        font.setSize(16);
        font.setColor(Color.DARK_GRAY);
        font.setBold(true);
        font.setName("Algerian");
        font.setUnderline(2);

        ParagraphFormat paragraphFormat = builder.getParagraphFormat();
        paragraphFormat.setFirstLineIndent(12);
        paragraphFormat.setAlignment(1);
        paragraphFormat.setKeepTogether(true);


        builder.write("This is a sample Paragraph");
        doc.save(dataDir + "output.doc");

    }
}
