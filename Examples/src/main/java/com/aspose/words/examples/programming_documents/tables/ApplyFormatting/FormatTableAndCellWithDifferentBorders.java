/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables.ApplyFormatting;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.awt.*;


public class FormatTableAndCellWithDifferentBorders {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(FormatTableAndCellWithDifferentBorders.class);
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);


        Table table = builder.startTable();
        builder.insertCell();

// Set the borders for the entire table.
        table.setBorders(LineStyle.SINGLE, 2.0, Color.black);
// Set the cell shading for this cell.
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.CYAN);
        builder.writeln("Cell #1");

        builder.insertCell();
// Specify a different cell shading for the second cell.
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.DARK_GRAY);
        builder.writeln("Cell #2");

// End this row.
        builder.endRow();

// Clear the cell formatting from previous operations.
        builder.getCellFormat().clearFormatting();

// Create the second row.
        builder.insertCell();

// Create larger borders for the first cell of this row. This will be different.
// compared to the borders set for the table.
        builder.getCellFormat().getBorders().getLeft().setLineStyle(4);
        builder.getCellFormat().getBorders().getRight().setLineStyle(4);
        builder.getCellFormat().getBorders().getTop().setLineStyle(4);
        builder.getCellFormat().getBorders().getBottom().setLineStyle(4);
        builder.writeln("Cell #3");

        builder.insertCell();
// Clear the cell formatting from the previous cell.
        builder.getCellFormat().clearFormatting();
        builder.writeln("Cell #4");
        // Save the document to disk.
        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}