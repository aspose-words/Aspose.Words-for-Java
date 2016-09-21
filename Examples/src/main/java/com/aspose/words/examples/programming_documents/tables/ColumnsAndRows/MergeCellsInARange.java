package com.aspose.words.examples.programming_documents.tables.ColumnsAndRows;

import java.awt.Point;
import java.awt.Rectangle;

import com.aspose.words.Cell;
import com.aspose.words.CellMerge;
import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.Row;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;

public class MergeCellsInARange {
	
	private static final String dataDir = Utils.getSharedDataDir(MergeCellsInARange.class) + "Tables/";
	
	public static void main(String[] args) throws Exception {
		Document doc = new Document(dataDir + "Table.SimpleTable.doc");
			
		// Retrieve the first table in the document.
		Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
				
		// We want to merge the range of cells found in between these two cells.
		Cell cellStartRange = table.getRows().get(1).getCells().get(1);
		Cell cellEndRange = table.getRows().get(2).getCells().get(2);

		// Merge all the cells between the two specified cells into one.
		mergeCells(cellStartRange, cellEndRange);
		
		doc.save(dataDir + "Table.MergeCellsInARange Out.doc");
	}

	/**
	 * Merges the range of cells found between the two specified cells both
	 * horizontally and vertically. Can span over multiple rows.
	 */
	public static void mergeCells(Cell startCell, Cell endCell) {
		Table parentTable = startCell.getParentRow().getParentTable();

		// Find the row and cell indices for the start and end cell.
		Point startCellPos = new Point(startCell.getParentRow().indexOf(startCell), parentTable.indexOf(startCell.getParentRow()));
		Point endCellPos = new Point(endCell.getParentRow().indexOf(endCell), parentTable.indexOf(endCell.getParentRow()));
		// Create the range of cells to be merged based off these indices. Inverse each index if the end cell if before the start cell.
		Rectangle mergeRange = new Rectangle(Math.min(startCellPos.x, endCellPos.x), Math.min(startCellPos.y, endCellPos.y), Math.abs(endCellPos.x - startCellPos.x) + 1,
				Math.abs(endCellPos.y - startCellPos.y) + 1);

		for (Row row : parentTable.getRows()) {
			for (Cell cell : row.getCells()) {
				Point currentPos = new Point(row.indexOf(cell), parentTable.indexOf(row));

				// Check if the current cell is inside our merge range then merge it.
				if (mergeRange.contains(currentPos)) {
					if (currentPos.x == mergeRange.x)
						cell.getCellFormat().setHorizontalMerge(CellMerge.FIRST);
					else
						cell.getCellFormat().setHorizontalMerge(CellMerge.PREVIOUS);

					if (currentPos.y == mergeRange.y)
						cell.getCellFormat().setVerticalMerge(CellMerge.FIRST);
					else
						cell.getCellFormat().setVerticalMerge(CellMerge.PREVIOUS);
				}
			}
		}
	}
}
