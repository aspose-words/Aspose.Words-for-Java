package com.aspose.words.examples.programming_documents.tables.ColumnsAndRows;

import com.aspose.words.Cell;
import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.Run;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;

public class WorkingWithColumns {
	
	private static final String dataDir = Utils.getSharedDataDir(WorkingWithColumns.class) + "Tables/";
	
	public static void main(String[] args) throws Exception {
		Document doc = new Document(dataDir + "Table.Document.doc");
		Table table = (Table)doc.getChild(NodeType.TABLE, 1, true);
		
		// Insert a blank column into a table
		insertABlankColumnIntoATable(doc, table);
		
		// Get the plain text of a table column
		getTextOfATableColumn(table);
		
		//Remove a column from a table in a document
		removeAColumnFromATable();
	}

	public static void insertABlankColumnIntoATable(Document doc, Table table) throws Exception {
		// Get the second column in the table.
		Column column = Column.fromIndex(table, 1);

		// Create a new column to the left of this column.
		// This is the same as using the "Insert Column Before" command in Microsoft Word.
		Column newColumn = column.insertColumnBefore();

		// Add some text to each of the column cells.
		for (Cell cell : newColumn.getCells()) {
			cell.getFirstParagraph().appendChild(new Run(doc, "Column Text " + newColumn.indexOf(cell)));
		}
	}
	
	public static void getTextOfATableColumn(Table table) throws Exception {
		// Get the first column in the table.
		Column column = Column.fromIndex(table, 0);

		// Print the plain text of the column to the screen.
		System.out.println(column.toTxt());
	}
	
	public static void removeAColumnFromATable() throws Exception {
		Document doc = new Document(dataDir + "Table.Document.doc");
		Table table = (Table)doc.getChild(NodeType.TABLE, 1, true);

		// Get the third column from the table and remove it.
		Column column = Column.fromIndex(table, 2);
		column.remove();

		doc.save(dataDir + "Table.RemoveColumn Out.doc");
	}
}
