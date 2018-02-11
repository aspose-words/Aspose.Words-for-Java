package com.aspose.words.examples.programming_documents.styles;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Style;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.StyleType;
import com.aspose.words.examples.Utils;

public class InsertStyleSeparator {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getSharedDataDir(InsertStyleSeparator.class) + "Styles/";

        // ExStart:ParagraphInsertStyleSeparator
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Style paraStyle = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "MyParaStyle");
        paraStyle.getFont().setBold(false);
        paraStyle.getFont().setSize(8);
        paraStyle.getFont().setName("Arial");

        // Append text with "Heading 1" style.
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.write("Heading 1");
        builder.insertStyleSeparator();

        // Append text with another style.
        builder.getParagraphFormat().setStyleName(paraStyle.getName());
        builder.write("This is text with some other formatting ");

        dataDir = dataDir + "InsertStyleSeparator_out.doc";

        // Save the document to disk.
        doc.save(dataDir);
        // ExEnd:ParagraphInsertStyleSeparator

        System.out.println("\nApplied different paragraph styles to two different parts of a text line successfully.\nFile saved at " + dataDir);
    }
}
