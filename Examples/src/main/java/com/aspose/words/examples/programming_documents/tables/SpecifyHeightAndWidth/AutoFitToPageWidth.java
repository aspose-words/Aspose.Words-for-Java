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


public class AutoFitToPageWidth {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AutoFitToPageWidth.class);
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);


        Table table = builder.startTable();
        builder.insertCell();

        table.setPreferredWidth( PreferredWidth.fromPercent(50));
        builder.writeln("Cell #1");

        builder.insertCell();
        builder.writeln("Cell #2");

        builder.insertCell();
        builder.writeln("Cell #3");

        // Save the document to disk.
        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}