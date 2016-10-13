package com.aspose.words.examples.programming_documents.tables.ColumnsAndRows;

import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.Row;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;

public class JoinAndSplitTables {

	private static final String dataDir = Utils.getSharedDataDir(JoinAndSplitTables.class) + "Tables/";

	public static void main(String[] args) throws Exception {
		// Combine the rows from two tables into one
		combineTwoTablesIntoOne();

		// Split a Table into Two Separate Tables
		splitATableIntoTwoSeparateTables();
	}

	public static void combineTwoTablesIntoOne() throws Exception {
		// Load the document.
		Document doc = new Document(dataDir + "Table.Document.doc");

		// Get the first and second table in the document.
		// The rows from the second table will be appended to the end of the first table.
		Table firstTable = (Table) doc.getChild(NodeType.TABLE, 0, true);
		Table secondTable = (Table) doc.getChild(NodeType.TABLE, 1, true);

		// Append all rows from the current table to the next.
		// Due to the design of tables even tables with different cell count and widths can be joined into one table.
		while (secondTable.hasChildNodes())
			firstTable.getRows().add(secondTable.getFirstRow());

		// Remove the empty table container.
		secondTable.remove();

		doc.save(dataDir + "Table.CombineTables Out.doc");
	}

	public static void splitATableIntoTwoSeparateTables() throws Exception {
		// Load the document.
		Document doc = new Document(dataDir + "Table.SimpleTable.doc");

		// Get the first table in the document.
		Table firstTable = (Table) doc.getChild(NodeType.TABLE, 0, true);

		// We will split the table at the third row (inclusive).
		Row row = firstTable.getRows().get(2);

		// Create a new container for the split table.
		Table table = (Table) firstTable.deepClone(false);

		// Insert the container after the original.
		firstTable.getParentNode().insertAfter(table, firstTable);

		// Add a buffer paragraph to ensure the tables stay apart.
		firstTable.getParentNode().insertAfter(new Paragraph(doc), firstTable);

		Row currentRow;

		do {
			currentRow = firstTable.getLastRow();
			table.prependChild(currentRow);
		} while (currentRow != row);

		doc.save(dataDir + "Table.SplitTable Out.doc");
	}
}
