package com.aspose.words.examples.programming_documents.tables.ColumnsAndRows;

import com.aspose.words.Cell;
import com.aspose.words.Document;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Row;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;
import com.aspose.words.examples.programming_documents.tables.ExtractOrReplaceText.ExtractPlainTextFromATable;

public class FindIndexOfTableElements {
	
	private static final String dataDir = Utils.getSharedDataDir(ExtractPlainTextFromATable.class) + "Tables/";
	
	public static void main(String[] args) throws Exception {
		Document doc = new Document(dataDir + "Table.SimpleTable.doc");
		Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);
			
		findIndexOfTableInADocument(doc, table);
	}

	public static void findIndexOfTableInADocument(Document doc, Table table) {
		NodeCollection allTables = doc.getChildNodes(NodeType.TABLE, true);
		int tableIndex = allTables.indexOf(table);
		System.out.println("Table Index: " + tableIndex);
	}
	
	public static void findIndexOfARowInATable(Table table, Row row) {
		int rowIndex = table.indexOf(row);
		System.out.println("Row Index: " + rowIndex);
	}
	
	public static void findIndexOfACellInARow(Row row, Cell cell) {
		int cellIndex = row.indexOf(cell);
		System.out.println("Cell Index: " + cellIndex);
	}
}
