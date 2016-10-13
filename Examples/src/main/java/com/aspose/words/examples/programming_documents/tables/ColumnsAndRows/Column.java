package com.aspose.words.examples.programming_documents.tables.ColumnsAndRows;

import java.util.ArrayList;

import com.aspose.words.Cell;
import com.aspose.words.Row;
import com.aspose.words.SaveFormat;
import com.aspose.words.Table;

/**
 * Represents a facade object for a column of a table in a Microsoft Word
 * document.
 */
public class Column {
	
	private int mColumnIndex;
	private Table mTable;
	
	private Column(Table table, int columnIndex) {
		if (table == null)
			throw new IllegalArgumentException("table");

		mTable = table;
		mColumnIndex = columnIndex;
	}

	/**
	 * Returns a new column facade from the table and supplied zero-based index.
	 */
	public static Column fromIndex(Table table, int columnIndex) {
		return new Column(table, columnIndex);
	}

	/**
	 * Returns the cells which make up the column.
	 */
	public Cell[] getCells() {
		ArrayList<Cell> columnCells = getColumnCells();
		return columnCells.toArray(new Cell[columnCells.size()]);
	}

	/**
	 * Returns the index of the given cell in the column.
	 */
	public int indexOf(Cell cell) {
		return getColumnCells().indexOf(cell);
	}

	/**
	 * Inserts a brand new column before this column into the table.
	 * @throws Exception 
	 */
	public Column insertColumnBefore() throws Exception {
		Cell[] columnCells = getCells();

		if (columnCells.length == 0)
			throw new IllegalArgumentException("Column must not be empty");

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
	public void remove() {
		for (Cell cell : getCells())
			cell.remove();
	}

	/**
	 * Returns the text of the column.
	 */
	public String toTxt() throws Exception {
		StringBuilder builder = new StringBuilder();

		for (Cell cell : getCells())
			builder.append(cell.toString(SaveFormat.TEXT));

		return builder.toString();
	}

	/**
	 * Provides an up-to-date collection of cells which make up the column
	 * represented by this facade.
	 */
	private ArrayList<Cell> getColumnCells() {
		ArrayList<Cell> columnCells = new ArrayList<Cell>();

		for (Row row : mTable.getRows()) {
			Cell cell = row.getCells().get(mColumnIndex);
			if (cell != null)
				columnCells.add(cell);
		}

		return columnCells;
	}
}