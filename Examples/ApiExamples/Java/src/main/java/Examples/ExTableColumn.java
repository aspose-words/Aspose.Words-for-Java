package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.util.ArrayList;

public class ExTableColumn extends ApiExampleBase {

    /**
     * Represents a facade object for a column of a table in a Microsoft Word document.
     */
    public static class Column {
        private Column(final Table table, final int columnIndex) {
            if (table == null) throw new IllegalArgumentException("table");

            mTable = table;
            mColumnIndex = columnIndex;
        }

        /**
         * Returns a new column facade from the table and supplied zero-based index.
         */
        public static Column fromIndex(final Table table, final int columnIndex) {
            return new Column(table, columnIndex);
        }

        /**
         * Returns the cells which make up the column.
         */
        public Cell[] getCells() {
            ArrayList columnCells = getColumnCells();
            return (Cell[]) columnCells.toArray(new Cell[columnCells.size()]);
        }

        /**
         * Returns the index of the given cell in the column.
         */
        public int indexOf(final Cell cell) {
            return getColumnCells().indexOf(cell);
        }

        /**
         * Inserts a new column before this column into the table.
         */
        public Column insertColumnBefore() {
            Cell[] columnCells = getCells();

            if (columnCells.length == 0) {
                throw new IllegalArgumentException("Column must not be empty");
            }

            // Create a clone of this column
            for (Cell cell : columnCells) {
                cell.getParentRow().insertBefore(cell.deepClone(false), cell);
            }

            Column column = new Column(columnCells[0].getParentRow().getParentTable(), mColumnIndex);

            // We want to make sure that the cells are all valid to work with (have at least one paragraph).
            for (Cell cell : column.getCells()) {
                cell.ensureMinimum();
            }

            // Increment the index of this column represents since there is a new column before it.
            mColumnIndex++;

            return column;
        }

        /**
         * Removes the column from the table.
         */
        public void remove() {
            for (Cell cell : getCells()) {
                cell.remove();
            }
        }

        /**
         * Returns the text of the column.
         */
        public String toTxt() throws Exception {
            StringBuilder builder = new StringBuilder();

            for (Cell cell : getCells()) {
                builder.append(cell.toString(SaveFormat.TEXT));
            }

            return builder.toString();
        }

        /**
         * Provides an up-to-date collection of cells which make up the column represented by this facade.
         */
        private ArrayList getColumnCells() {
            ArrayList columnCells = new ArrayList();

            for (Row row : mTable.getRows()) {
                Cell cell = row.getCells().get(mColumnIndex);
                if (cell != null) {
                    columnCells.add(cell);
                }
            }

            return columnCells;
        }

        private int mColumnIndex;
        private final Table mTable;
    }

    @Test
    public void removeColumnFromTable() throws Exception {
        Document doc = new Document(getMyDir() + "Tables.docx");
        Table table = (Table) doc.getChild(NodeType.TABLE, 1, true);

        Column column = Column.fromIndex(table, 2);
        column.remove();

        doc.save(getArtifactsDir() + "TableColumn.RemoveColumn.doc");

        Assert.assertEquals(table.getChildNodes(NodeType.CELL, true).getCount(), 16);
        Assert.assertEquals(table.getRows().get(2).getCells().get(2).toString(SaveFormat.TEXT).trim(), "Cell 7 contents");
        Assert.assertEquals(table.getLastRow().getCells().get(2).toString(SaveFormat.TEXT).trim(), "Cell 11 contents");
    }

    @Test
    public void insert() throws Exception {
        Document doc = new Document(getMyDir() + "Tables.docx");
        Table table = (Table) doc.getChild(NodeType.TABLE, 1, true);

        Column column = Column.fromIndex(table, 1);

        // Create a new column to the left of this column.
        // This is the same as using the "Insert Column Before" command in Microsoft Word.
        Column newColumn = column.insertColumnBefore();

        // Add some text to each cell in the column.
        for (Cell cell : newColumn.getCells())
            cell.getFirstParagraph().appendChild(new Run(doc, "Column Text " + newColumn.indexOf(cell)));

        doc.save(getArtifactsDir() + "TableColumn.Insert.doc");

        Assert.assertEquals(table.getChildNodes(NodeType.CELL, true).getCount(), 24);
        Assert.assertEquals(table.getFirstRow().getCells().get(1).toString(SaveFormat.TEXT).trim(), "Column Text 0");
        Assert.assertEquals(table.getLastRow().getCells().get(1).toString(SaveFormat.TEXT).trim(), "Column Text 3");
    }

    @Test
    public void tableColumnToTxt() throws Exception {
        Document doc = new Document(getMyDir() + "Tables.docx");
        Table table = (Table) doc.getChild(NodeType.TABLE, 1, true);

        Column column = Column.fromIndex(table, 0);
        System.out.println(column.toTxt());

        Assert.assertEquals(column.toTxt(), "\rRow 1\rRow 2\rRow 3\r");
    }
}
