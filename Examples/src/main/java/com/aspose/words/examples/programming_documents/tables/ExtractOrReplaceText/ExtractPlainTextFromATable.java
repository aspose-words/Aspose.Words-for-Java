package com.aspose.words.examples.programming_documents.tables.ExtractOrReplaceText;

import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;

public class ExtractPlainTextFromATable {

	private static final String dataDir = Utils.getSharedDataDir(ExtractPlainTextFromATable.class) + "Tables/";

	public static void main(String[] args) throws Exception {
		// Print the text range of a table
		printTextRangeOfATable();
		// Print the text range of row and table elements
		printTextRangeOfRowAndTableElements();
	}

	public static void printTextRangeOfATable() throws Exception {
		Document doc = new Document(dataDir + "Table.SimpleTable.doc");

		// Get the first table in the document.
		Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);

		// The range text will include control characters such as "\a" for a cell.
		// You can call ToTxt() on the desired node to find the plain text.

		// Print the plain text range of the table to the screen.
		System.out.println("Contents of the table: ");
		System.out.println(table.getRange().getText());
	}
	
	public static void printTextRangeOfRowAndTableElements() throws Exception {
		Document doc = new Document(dataDir + "Table.SimpleTable.doc");

		// Get the first table in the document.
		Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);
		
		// Print the contents of the first row to the screen.
		System.out.println("\nContents of the row: ");
		System.out.println(table.getFirstRow().getRange().getText());

		// Print the contents of the last cell in the table to the screen.
		System.out.println("\nContents of the cell: ");
		System.out.println(table.getLastRow().getLastCell().getRange().getText());
	}
}