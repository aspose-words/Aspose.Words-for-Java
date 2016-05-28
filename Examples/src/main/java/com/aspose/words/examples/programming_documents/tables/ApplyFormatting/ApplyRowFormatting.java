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


public class ApplyRowFormatting {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApplyRowFormatting.class);
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);


        Table table = builder.startTable();
        builder.insertCell();

// Set the row formatting
        RowFormat rowFormat = builder.getRowFormat();
        rowFormat.setHeight(100);
        rowFormat.setHeightRule(HeightRule.EXACTLY);
// These formatting properties are set on the table and are applied to all rows in the table.
        table.setLeftPadding(30);
        table.setRightPadding(30);
        table.setTopPadding(30);
        table.setBottomPadding(30);

        builder.writeln("I'm a wonderful formatted row.");

        builder.endRow();
        builder.endTable();

        // Save the document to disk.
        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}