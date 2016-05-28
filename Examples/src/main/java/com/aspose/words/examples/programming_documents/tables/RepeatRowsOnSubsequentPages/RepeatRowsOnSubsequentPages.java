/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables.RepeatRowsOnSubsequentPages;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class RepeatRowsOnSubsequentPages {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RepeatRowsOnSubsequentPages.class);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);


        Table table = builder.startTable();
        builder.getRowFormat().setHeadingFormat(true);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.getCellFormat().setWidth(100);
        builder.insertCell();
        builder.writeln("Heading row 1");
        builder.endRow();
        builder.insertCell();
        builder.writeln("Heading row 2");
        builder.endRow();

        builder.getCellFormat().setWidth(50);
        builder.getParagraphFormat().clearFormatting();

        // Insert some content so the table is long enough to continue onto the next page.
        for (int i = 0; i < 50; i++)
        {
            builder.insertCell();
            builder.getRowFormat().setHeadingFormat(false);
            builder.writeln("Column 1 Text");
            builder.insertCell();
            builder.writeln("Column 2 Text");
            builder.endRow();
        }
        // Save the document to disk.
        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}