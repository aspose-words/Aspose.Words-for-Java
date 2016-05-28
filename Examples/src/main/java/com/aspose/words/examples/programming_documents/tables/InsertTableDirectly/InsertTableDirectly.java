/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables.InsertTableDirectly;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.awt.*;


public class InsertTableDirectly {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(InsertTableDirectly.class);

        // For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);


        Table table = new Table(doc);
        // Add the table to the document.
        doc.getFirstSection().getBody().appendChild(table);

        Row row = new Row(doc);
        row.getRowFormat().setAllowBreakAcrossPages(true);
        table.appendChild(row);

// We can now apply any auto fit settings.
        table.autoFit(AutoFitBehavior.FIXED_COLUMN_WIDTHS);

// Create a cell and add it to the row
        Cell cell = new Cell(doc);
        cell.getCellFormat().getShading().setBackgroundPatternColor(Color.blue);

// Add a paragraph to the cell as well as a new run with some text.
        cell.appendChild(new Paragraph(doc));
        cell.getFirstParagraph().appendChild(new Run(doc, "Row 1, Cell 1 Text"));

// Add the cell to the row.
        row.appendChild(cell);

// We would then repeat the process for the other cells and rows in the table.
// We can also speed things up by cloning existing cells and rows.
        row.appendChild(cell.deepClone(false));
        row.getLastCell().appendChild(new Paragraph(doc));
        row.getLastCell().getFirstParagraph().appendChild(new Run(doc, "Row 1, Cell 2 Text"));

        // Save the document to disk.
        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}