/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables.SpecifyHeightAndWidth;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.PreferredWidth;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;

import java.awt.*;


public class SetPreferredWidthSettings {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SetPreferredWidthSettings.class);
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);


        Table table = builder.startTable();
        builder.insertCell();

        builder.getCellFormat().setPreferredWidth(PreferredWidth.fromPoints(40));
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.CYAN);
        builder.writeln("Cell at 40 points width");

        // Insert a relative (percent) sized cell.
        builder.insertCell();
        builder.getCellFormat().setPreferredWidth(PreferredWidth.fromPercent(20));
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.BLUE);
        builder.writeln("Cell at 20% width");

        // Insert a auto sized cell.
        builder.insertCell();
        builder.getCellFormat().setPreferredWidth(PreferredWidth.AUTO);
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.GREEN);
        builder.writeln("Cell automatically sized. The size of this cell is calculated from the table preferred width.");
        builder.writeln("In this case the cell will fill up the rest of the available space.");
        // Save the document to disk.
        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}