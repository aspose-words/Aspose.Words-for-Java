/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class DocumentBuilderInsertTableOfContents {
    public static void main(String[] args) throws Exception {
        //ExStart:1
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

        //ExEnd:1
    }
}