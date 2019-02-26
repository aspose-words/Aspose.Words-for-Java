//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import org.testng.Assert;
import org.testng.annotations.Test;

import java.util.ArrayList;

import com.aspose.words.*;

public class ExTableColumn extends ApiExampleBase
{

    //ExStart
    //ExId:ColumnFacade
    //ExSummary:Demonstrates a facade object for working with a column of a table.

    /**
     * Represents a facade object for a column of a table in a Microsoft Word document.
     */
    public static class Column
    {
        private Column(Table table, int columnIndex)
        {
            if (table == null) throw new IllegalArgumentException("table");

            mTable = table;
            mColumnIndex = columnIndex;
        }

        /**
         * Returns a new column facade from the table and supplied zero-based index.
         */
        public static Column fromIndex(Table table, int columnIndex)
        {
            return new Column(table, columnIndex);
        }

        /**
         * Returns the cells which make up the column.
         */
        public Cell[] getCells()
        {
            ArrayList columnCells = getColumnCells();
            return (Cell[]) columnCells.toArray(new Cell[columnCells.size()]);
        }

        /**
         * Returns the index of the given cell in the column.
         */
        public int indexOf(Cell cell)
        {
            return getColumnCells().indexOf(cell);
        }

        /**
         * Inserts a brand new column before this column into the table.
         */
        public Column insertColumnBefore()
        {
            Cell[] columnCells = getCells();

            if (columnCells.length == 0) throw new IllegalArgumentException("Column must not be empty");

            // Create a clone of this column.
            for (Cell cell : columnCells)
                cell.getParentRow().insertBefore(cell.deepClone(false), cell);

            // This is the new column.
            Column column = new Column(columnCells[0].getParentRow().getParentTable(), mColumnIndex);

            // We want to make sure that the cells are all valid to work with (have at least one paragraph).
            for (Cell cell : column.getCells())
                cell.ensureMinimum();

            // Increase the index which this column represents since there is now one extra column infront.
            mColumnIndex++;

            return column;
        }

        /**
         * Removes the column from the table.
         */
        public void remove()
        {
            for (Cell cell : getCells())
                cell.remove();
        }

        /**
         * Returns the text of the column.
         */
        public String toTxt() throws Exception
        {
            StringBuilder builder = new StringBuilder();

            for (Cell cell : getCells())
                builder.append(cell.toString(SaveFormat.TEXT));

            return builder.toString();
        }

        /**
         * Provides an up-to-date collection of cells which make up the column represented by this facade.
         */
        private ArrayList getColumnCells()
        {
            ArrayList columnCells = new ArrayList();

            for (Row row : mTable.getRows())
            {
                Cell cell = row.getCells().get(mColumnIndex);
                if (cell != null) columnCells.add(cell);
            }

            return columnCells;
        }

        private int mColumnIndex;
        private Table mTable;
    }
    //ExEnd

    @Test
    public void RemoveColumnFromTable() throws Exception
    {
        //ExStart
        //ExId:RemoveTableColumn
        //ExSummary:Shows how to remove a column from a table in a document.
        Document doc = new Document(getMyDir() + "Table.Document.doc");
        Table table = (Table) doc.getChild(NodeType.TABLE, 1, true);

        // Get the third column from the table and remove it.
        Column column = Column.fromIndex(table, 2);
        column.remove();
        //ExEnd

        doc.save(getMyDir() + "\\Artifacts\\Table.RemoveColumn.doc");

        Assert.assertEquals(table.getChildNodes(NodeType.CELL, true).getCount(), 16);
        Assert.assertEquals(table.getRows().get(2).getCells().get(2).toString(SaveFormat.TEXT).trim(), "Cell 3 contents");
        Assert.assertEquals(table.getLastRow().getCells().get(2).toString(SaveFormat.TEXT).trim(), "Cell 3 contents");
    }

    @Test
    public void InsertNewColumnIntoTable() throws Exception
    {
        Document doc = new Document(getMyDir() + "Table.Document.doc");
        Table table = (Table) doc.getChild(NodeType.TABLE, 1, true);

        //ExStart
        //ExId:InsertNewColumn
        //ExSummary:Shows how to insert a blank column into a table.
        // Get the second column in the table.
        Column column = Column.fromIndex(table, 1);

        // Create a new column to the left of this column.
        // This is the same as using the "Insert Column Before" command in Microsoft Word.
        Column newColumn = column.insertColumnBefore();

        // Add some text to each of the column cells.
        for (Cell cell : newColumn.getCells())
            cell.getFirstParagraph().appendChild(new Run(doc, "Column Text " + newColumn.indexOf(cell)));
        //ExEnd

        doc.save(getMyDir() + "\\Artifacts\\Table.InsertColumn.doc");

        Assert.assertEquals(table.getChildNodes(NodeType.CELL, true).getCount(), 24);
        Assert.assertEquals(table.getFirstRow().getCells().get(1).toString(SaveFormat.TEXT).trim(), "Column Text 0");
        Assert.assertEquals(table.getLastRow().getCells().get(1).toString(SaveFormat.TEXT).trim(), "Column Text 3");
    }

    @Test
    public void TableColumnToTxt() throws Exception
    {
        Document doc = new Document(getMyDir() + "Table.Document.doc");
        Table table = (Table) doc.getChild(NodeType.TABLE, 1, true);

        //ExStart
        //ExId:TableColumnToTxt
        //ExSummary:Shows how to get the plain text of a table column.
        // Get the first column in the table.
        Column column = Column.fromIndex(table, 0);

        // Print the plain text of the column to the screen.
        System.out.println(column.toTxt());
        //ExEnd

        Assert.assertEquals(column.toTxt(), "\r\nRow 1\r\nRow 2\r\nRow 3\r\n");
    }
}
