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
import com.sun.prism.paint.Color;


public class ExpandFormattingOnCellsAndRowFromStyle {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ExpandFormattingOnCellsAndRowFromStyle.class);

        Document doc = new Document(dataDir + "Table.TableStyle.docx");

        // Get the first cell of the first table in the document.
        Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);
        Cell firstCell = table.getFirstRow().getFirstCell();

        // First print the color of the cell shading. This should be empty as the current shading
        // is stored in the table style.
        java.awt.Color cellShadingBefore = firstCell.getCellFormat().getShading().getBackgroundPatternColor();
        System.out.println("Cell shading before style expansion: " + cellShadingBefore.toString());

        // Expand table style formatting to direct formatting.
        doc.expandTableStylesToDirectFormatting();

        // Now print the cell shading after expanding table styles. A blue background pattern color
        // should have been applied from the table style.
        java.awt.Color  cellShadingAfter = firstCell.getCellFormat().getShading().getBackgroundPatternColor();
        System.out.println("Cell shading after style expansion: " + cellShadingAfter.toString());

        // Save the document to disk.
        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}