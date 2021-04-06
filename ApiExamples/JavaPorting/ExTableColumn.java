// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Table;
import com.aspose.words.Cell;
import com.aspose.ms.System.Collections.msArrayList;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.words.SaveFormat;
import java.util.ArrayList;
import com.aspose.words.Row;
import com.aspose.words.Document;
import com.aspose.words.NodeType;
import org.testng.Assert;
import com.aspose.words.Run;
import com.aspose.ms.System.msConsole;


@Test
public class ExTableColumn extends ApiExampleBase
{
    /// <summary>
    /// Represents a facade object for a column of a table in a Microsoft Word document.
    /// </summary>
    public static class Column
    {
        private Column(Table table, int columnIndex)
        {
            mTable =  !!Autoporter warning: Not supported language construction  throw new IllegalArgumentException("table");
            mColumnIndex = columnIndex;
        }

        /// <summary>
        /// Returns a new column facade from the table and supplied zero-based index.
        /// </summary>
        public static Column fromIndex(Table table, int columnIndex)
        {
            return new Column(table, columnIndex);
        }

        /// <summary>
        /// Returns the cells which make up the column.
        /// </summary>
        public Cell[] getCells() { return msArrayList.toArray(getColumnCells(), new Cell[0]); }

        /// <summary>
        /// Returns the index of the given cell in the column.
        /// </summary>
        public int indexOf(Cell cell)
        {
            return getColumnCells().indexOf(cell);
        }

        /// <summary>
        /// Inserts a new column before this column into the table.
        /// </summary>
        public Column insertColumnBefore()
        {
            Cell[] columnCells = getCells();

            if (columnCells.length == 0)
                throw new IllegalArgumentException("Column must not be empty");

            // Create a clone of this column
            for (Cell cell : columnCells)
                cell.getParentRow().insertBefore(cell.deepClone(false), cell);
            
            Column newColumn = new Column(columnCells[0].getParentRow().getParentTable(), mColumnIndex);

            // We want to make sure that the cells are all valid to work with (have at least one paragraph).
            for (Cell cell : newColumn.getCells())
                cell.ensureMinimum();

            // Increment the index of this column represents since there is a new column before it.
            mColumnIndex++;

            return newColumn;
        }

        /// <summary>
        /// Removes the column from the table.
        /// </summary>
        public void remove()
        {
            for (Cell cell : getCells())
                cell.remove();
        }

        /// <summary>
        /// Returns the text of the column. 
        /// </summary>
        public String toTxt() throws Exception
        {
            StringBuilder builder = new StringBuilder();

            for (Cell cell : getCells())
                msStringBuilder.append(builder, cell.toString(SaveFormat.TEXT));

            return builder.toString();
        }

        /// <summary>
        /// Provides an up-to-date collection of cells which make up the column represented by this facade.
        /// </summary>
        private ArrayList<Cell> getColumnCells()
        {
            ArrayList<Cell> columnCells = new ArrayList<Cell>();

            for (Row row : mTable.getRows().<Row>OfType() !!Autoporter error: Undefined expression type )
            {
                Cell cell = row.getCells().get(mColumnIndex);
                if (cell != null)
                    columnCells.add(cell);
            }

            return columnCells;
        }

        private int mColumnIndex;
        private /*final*/ Table mTable;
    }
    
    @Test
    public void removeColumnFromTable() throws Exception
    {
        Document doc = new Document(getMyDir() + "Tables.docx");
        Table table = (Table) doc.getChild(NodeType.TABLE, 1, true);

        Column column = Column.fromIndex(table, 2);
        column.remove();
        
        doc.save(getArtifactsDir() + "TableColumn.RemoveColumn.doc");

        Assert.assertEquals(16, table.getChildNodes(NodeType.CELL, true).getCount());
        Assert.assertEquals("Cell 7 contents", table.getRows().get(2).getCells().get(2).toString(SaveFormat.TEXT).trim());
        Assert.assertEquals("Cell 11 contents", table.getLastRow().getCells().get(2).toString(SaveFormat.TEXT).trim());
    }

    @Test
    public void insert() throws Exception
    {
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

        Assert.assertEquals(24, table.getChildNodes(NodeType.CELL, true).getCount());
        Assert.assertEquals("Column Text 0", table.getFirstRow().getCells().get(1).toString(SaveFormat.TEXT).trim());
        Assert.assertEquals("Column Text 3", table.getLastRow().getCells().get(1).toString(SaveFormat.TEXT).trim());
    }

    @Test
    public void tableColumnToTxt() throws Exception
    {
        Document doc = new Document(getMyDir() + "Tables.docx");
        Table table = (Table) doc.getChild(NodeType.TABLE, 1, true);

        Column column = Column.fromIndex(table, 0);
        System.out.println(column.toTxt());

        Assert.assertEquals("\r\nRow 1\r\nRow 2\r\nRow 3\r\n", column.toTxt());
    }
}
