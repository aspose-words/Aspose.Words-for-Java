package com.aspose.words.examples.programming_documents.tables.ApplyFormatting;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.PreferredWidth;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;

public class SpecifyAPreferredWidthOnATable {
	
	private static final String dataDir = Utils.getSharedDataDir(SpecifyAPreferredWidthOnATable.class) + "Tables/";
	
	public static void main(String[] args) throws Exception {
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		// Insert a table with a width that takes up half the page width.
		Table table = builder.startTable();

		// Insert a few cells
		builder.insertCell();
		table.setPreferredWidth(PreferredWidth.fromPercent(50));
		builder.writeln("Cell #1");

		builder.insertCell();
		builder.writeln("Cell #2");

		builder.insertCell();
		builder.writeln("Cell #3");

		doc.save(dataDir + "Table.PreferredWidth Out.doc");
	}
}
