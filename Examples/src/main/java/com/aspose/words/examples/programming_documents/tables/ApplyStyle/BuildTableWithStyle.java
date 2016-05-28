/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables.ApplyStyle;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class BuildTableWithStyle {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(BuildTableWithStyle.class);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        // We must insert at least one row first before setting any table formatting.
        builder.insertCell();
        // Set the table style used based of the unique style identifier.
        // Note that not all table styles are available when saving as .doc format.
        table.setStyleIdentifier(StyleIdentifier.MEDIUM_SHADING_1_ACCENT_2);
        // Apply which features should be formatted by the style.
        table.setStyleOptions(TableStyleOptions.FIRST_COLUMN | TableStyleOptions.ROW_BANDS |TableStyleOptions.FIRST_ROW);
        table.autoFit(AutoFitBehavior.AUTO_FIT_TO_CONTENTS);

        // Continue with building the table as normal.
        builder.writeln("Item");
        builder.getCellFormat().setRightPadding(40);
        builder.insertCell();
        builder.writeln("Quantity (kg)");
        builder.endRow();

        builder.insertCell();
        builder.write("Apples");
        builder.insertCell();
        builder.write("20");
        builder.endRow();

        builder.insertCell();
        builder.write("Bananas");
        builder.insertCell();
        builder.write("40");
        builder.endRow();

        builder.insertCell();
        builder.write("Carrots");
        builder.insertCell();
        builder.writeln("50");
        builder.endRow();

        // Save the document to disk.
        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}