package com.aspose.words.examples.programming_documents.tables.creation;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class InsertCloneOfExistingTable {

    private static final String dataDir = Utils.getSharedDataDir(InsertCloneOfExistingTable.class) + "Tables/";

    public static void main(String[] args) throws Exception {
        //ExStart:InsertCloneOfExistingTable
        // Make a clone of a table in the document and insert it after the original table
        cloneOfATable();

        // Remove all content from the cells of a cloned table
        removeAllContentFromCellsOfAClonedTable();

        // Make a clone of the last row of a table and append it to the table
        cloneLastRowOfATable();
        //ExEnd:InsertCloneOfExistingTable
    }

    //ExStart:cloneOfATable
    public static void cloneOfATable() throws Exception {
        Document doc = new Document(dataDir + "Table.SimpleTable.doc");

        // Retrieve the first table in the document.
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Create a clone of the table.
        Table tableClone = (Table) table.deepClone(true);

        // Insert the cloned table into the document after the original
        table.getParentNode().insertAfter(tableClone, table);

        // Insert an empty paragraph between the two tables or else they will be combined into one
        // upon save. This has to do with document validation.
        table.getParentNode().insertAfter(new Paragraph(doc), table);

        doc.save(dataDir + "Table_CloneTableAndInsert_Out.doc");
    }
    //ExEnd:cloneOfATable

    //ExStart:removeAllContentFromCellsOfAClonedTable
    public static void removeAllContentFromCellsOfAClonedTable() throws Exception {
        Document doc = new Document(dataDir + "Table.SimpleTable.doc");

        // Retrieve the first table in the document.
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Create a clone of the table.
        Table tableClone = (Table) table.deepClone(true);

        for (Cell cell : (Iterable<Cell>) tableClone.getChildNodes(NodeType.CELL, true)) {
            cell.removeAllChildren();
        }

        // Insert the cloned table into the document after the original
        table.getParentNode().insertAfter(tableClone, table);

        // Insert an empty paragraph between the two tables or else they will be combined into one
        // upon save. This has to do with document validation.
        table.getParentNode().insertAfter(new Paragraph(doc), table);

        doc.save(dataDir + "RemoveAllContentFromCellsOfAClonedTable_Out.doc");
    }
    //ExEnd:removeAllContentFromCellsOfAClonedTable

    //ExStart:cloneLastRowOfATable
    public static void cloneLastRowOfATable() throws Exception {
        Document doc = new Document(dataDir + "Table.SimpleTable.doc");

        // Retrieve the first table in the document.
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Clone the last row in the table.
        Row clonedRow = (Row) table.getLastRow().deepClone(true);

        // Remove all content from the cloned row's cells. This makes the row ready for
        // new content to be inserted into.
        for (Cell cell : clonedRow.getCells())
            cell.removeAllChildren();

        // Add the row to the end of the table.
        table.appendChild(clonedRow);

        doc.save(dataDir + "Table.AddCloneRowToTable_Out.doc");
    }
    //ExEnd:cloneLastRowOfATable
}
