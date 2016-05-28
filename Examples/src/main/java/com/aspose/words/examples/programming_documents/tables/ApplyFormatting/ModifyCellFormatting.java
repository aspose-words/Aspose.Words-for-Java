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


public class ModifyCellFormatting {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ModifyCellFormatting.class);
        Document doc = new Document(dataDir + "Table.Document.doc");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Retrieve the first row in the table.
        Cell firstCell = table.getFirstRow().getFirstCell();

        // Modify some cell level properties.
        firstCell.getCellFormat().setWidth(30);
        firstCell.getCellFormat().setOrientation(TextOrientation.HORIZONTAL_ROTATED_FAR_EAST);
        firstCell.getCellFormat().getShading().setBackgroundPatternColor(Color.ORANGE);

        // Save the document to disk.
        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}