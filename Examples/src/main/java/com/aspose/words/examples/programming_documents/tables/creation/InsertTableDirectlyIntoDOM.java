package com.aspose.words.examples.programming_documents.tables.creation;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.awt.*;

public class InsertTableDirectlyIntoDOM {

    private static final String dataDir = Utils.getSharedDataDir(InsertTableDirectlyIntoDOM.class) + "Tables/";

    public static void main(String[] args) throws Exception {
        //ExStart:InsertTableDirectlyIntoDOM
        Document doc = new Document();

        // We start by creating the table object. Note how we must pass the document object
        // to the constructor of each node. This is because every node we create must belong
        // to some document.
        Table table = new Table(doc);
        // Add the table to the document.
        doc.getFirstSection().getBody().appendChild(table);

        // Here we could call EnsureMinimum to create the rows and cells for us. This method is used
        // to ensure that the specified node is valid, in this case a valid table should have at least one
        // row and one cell, therefore this method creates them for us.

        // Instead we will handle creating the row and table ourselves. This would be the best way to do this
        // if we were creating a table inside an algorthim for example.
        Row row = new Row(doc);
        row.getRowFormat().setAllowBreakAcrossPages(true);
        table.appendChild(row);

        // We can now apply any auto fit settings.
        table.autoFit(AutoFitBehavior.FIXED_COLUMN_WIDTHS);

        // Create a cell and add it to the row
        Cell cell = new Cell(doc);
        cell.getCellFormat().getShading().setBackgroundPatternColor(Color.BLUE);
        cell.getCellFormat().setWidth(80);

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

        doc.save(dataDir + "Table_InsertTableUsingNodes_Out.doc");
        //ExEnd:InsertTableDirectlyIntoDOM
    }
}