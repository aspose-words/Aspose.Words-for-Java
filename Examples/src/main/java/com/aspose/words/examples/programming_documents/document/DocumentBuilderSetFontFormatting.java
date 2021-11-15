package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Font;
import com.aspose.words.Underline;

import java.awt.*;

public class DocumentBuilderSetFontFormatting {

    public static void main(String[] args) throws Exception {
        // ExStart: DocumentBuilderSetFontFormatting
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set font formatting properties
        Font font = builder.getFont();
        font.setBold(true);
        font.setColor(Color.BLUE);
        font.setItalic(true);
        font.setName("Arial");
        font.setSize(24);
        font.setSpacing(5);
        font.setUnderline(Underline.DOUBLE);

        builder.write("I'm a very nice formatted string.");
        // ExEnd: DocumentBuilderSetFontFormatting
    }

}
