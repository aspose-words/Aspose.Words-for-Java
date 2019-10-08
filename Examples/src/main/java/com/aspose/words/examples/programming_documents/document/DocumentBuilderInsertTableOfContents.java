package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.BreakType;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.examples.Utils;


public class DocumentBuilderInsertTableOfContents {
    public static void main(String[] args) throws Exception {

        //ExStart:DocumentBuilderInsertTableOfContents
        String dataDir = Utils.getDataDir(DocumentBuilderInsertTableOfContents.class);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.insertBreak(BreakType.PAGE_BREAK);

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.writeln("Heading 1");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);
        builder.writeln("Heading 1.1");
        builder.writeln("Heading 1.2");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.writeln("Heading 2");
        builder.writeln("Heading 3");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);
        builder.writeln("Heading 3.1");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_3);
        builder.writeln("Heading 3.1.1");
        builder.writeln("Heading 3.1.2");
        builder.writeln("Heading 3.1.3");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);
        builder.writeln("Heading 3.2");
        builder.writeln("Heading 3.3");

        doc.updateFields();
        doc.save(dataDir + "output.doc");
        //ExEnd:DocumentBuilderInsertTableOfContents

    }
}