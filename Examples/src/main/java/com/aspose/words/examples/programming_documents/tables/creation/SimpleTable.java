package com.aspose.words.examples.programming_documents.tables.creation;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class SimpleTable {
	
	private static final String dataDir = Utils.getSharedDataDir(SimpleTable.class) + "Tables/";
	
	public static void main(String[] args) throws Exception {
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		// We call this method to start building the table.
		builder.startTable();
		builder.insertCell();
		builder.write("Row 1, Cell 1 Content.");

		// Build the second cell
		builder.insertCell();
		builder.write("Row 1, Cell 2 Content.");
		// Call the following method to end the row and start a new row.
		builder.endRow();

		// Build the first cell of the second row.
		builder.insertCell();
		builder.write("Row 2, Cell 1 Content");

		// Build the second cell.
		builder.insertCell();
		builder.write("Row 2, Cell 2 Content.");
		builder.endRow();

		// Signal that we have finished building the table.
		builder.endTable();

		// Save the document to disk.
		doc.save(dataDir + "DocumentBuilder_CreateSimpleTable_Out.doc");
	}
}