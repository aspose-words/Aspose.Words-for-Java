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


public class DocumentBuilderSetTableCellFormatting {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderSetTableCellFormatting.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertCell();
        CellFormat cellFormat =builder.getCellFormat();
        cellFormat.setWidth(250);
        cellFormat.setLeftPadding(30);
        cellFormat.setRightPadding(30);
        cellFormat.setBottomPadding(30);
        cellFormat.setTopPadding(30);

        builder.writeln("I'm a wonderful formatted cell.");
        builder.endRow();
        builder.endTable();
        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}