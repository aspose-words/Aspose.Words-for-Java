package com.aspose.words.examples.programming_documents.tableofcontents;

import com.aspose.words.BreakType;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.examples.Utils;

public class InsertATableOfContentsUsingHeadingStyles {

    private static final String dataDir = Utils.getSharedDataDir(InsertATableOfContentsUsingHeadingStyles.class) + "TableOfContents/";

    public static void main(String[] args) throws Exception {

        //ExStart:InsertATableOfContentsUsingHeadingStyles
        Document doc = new Document();

        // Create a document builder to insert content with into document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table of contents at the beginning of the document.
        builder.insertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Start the actual document content on the second page.
        builder.insertBreak(BreakType.PAGE_BREAK);

        // Build a document with complex structure by applying different heading styles thus creating TOC entries.
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

        // Call the method below to update the TOC.
        doc.updateFields();

        doc.save(dataDir + "InsertATableOfContentsUsingHeadingStyles_out.docx");
        //ExEnd:InsertATableOfContentsUsingHeadingStyles
    }
}