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
import com.sun.prism.paint.*;

import java.awt.*;
import java.awt.Color;


public class BuildTableWithBordersEnabled {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(BuildTableWithBordersEnabled.class);
        Document doc = new Document(dataDir + "Table.EmptyTable.doc");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        table.clearBorders();
        table.setBorders(LineStyle.SINGLE,1,Color.RED);

        // Save the document to disk.
        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}