package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Font;
import com.aspose.words.Underline;

import java.awt.*;

public class WriteAndFont {
    public static void main(String[] args) throws Exception {
        
        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Font font = builder.getFont();
        font.setSize(16);
        font.setColor(Color.blue);
        font.setBold(true);
        font.setName("Algerian");
        font.setUnderline(Underline.DOUBLE);
        builder.write("aspose......... aspose_words_java");
    }
}