package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Font;
import com.aspose.words.Underline;
import com.aspose.words.examples.Utils;

import java.awt.*;


public class DocumentBuilderSetFontFormatting {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderSetFontFormatting.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Font font = builder.getFont();
        font.setSize(16);
        font.setColor(Color.blue);
        font.setBold(true);
        font.setName("Arial");
        font.setUnderline(Underline.DOTTED);
        builder.write("I'm a very nice formatted string.");
        doc.save(dataDir + "output.doc");

    }
}