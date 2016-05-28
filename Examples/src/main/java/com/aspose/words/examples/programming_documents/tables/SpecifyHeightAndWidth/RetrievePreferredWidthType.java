/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables.SpecifyHeightAndWidth;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.awt.*;


public class RetrievePreferredWidthType {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RetrievePreferredWidthType.class);
        Document doc = new Document(dataDir + "Table.SimpleTable.doc");


        // Retrieve the first table in the document.
        Table table = (Table)doc.getChild(NodeType.TABLE,0,true);
        table.setAllowAutoFit(true);

        Cell firstCell = table.getFirstRow().getFirstCell();
        int type = firstCell.getCellFormat().getPreferredWidth().getType();
        double value = firstCell.getCellFormat().getPreferredWidth().getValue();

        System.out.println("Type : "+ type + "\nValue : " + value);

        //ExEnd:1
    }
}