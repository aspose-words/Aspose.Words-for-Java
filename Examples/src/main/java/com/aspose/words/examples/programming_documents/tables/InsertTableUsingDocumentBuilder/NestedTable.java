/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables.InsertTableUsingDocumentBuilder;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.awt.*;


public class NestedTable {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(NestedTable.class);

        // For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the outer table.
        Cell cell = builder.insertCell();
        builder.writeln("Outer Table Cell 1");

        builder.insertCell();
        builder.writeln("Outer Table Cell 2");

        // This call is important in order to create a nested table within the first table
        // Without this call the cells inserted below will be appended to the outer table.builder.endTable();

        builder.endTable();
        // Move to the first cell of the outer table.
        builder.moveTo(cell.getFirstParagraph());

        // Build the inner table.
        builder.insertCell();
        builder.writeln("Inner Table Cell 1");

        builder.insertCell();
        builder.writeln("Inner Table Cell 2");
        builder.endTable();

        // Save the document to disk.
        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}