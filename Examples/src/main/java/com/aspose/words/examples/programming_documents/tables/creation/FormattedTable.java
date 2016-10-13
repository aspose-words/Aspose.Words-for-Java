package com.aspose.words.examples.programming_documents.tables.creation;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.awt.*;

public class FormattedTable {

	private static final String dataDir = Utils.getSharedDataDir(FormattedTable.class) + "Tables/";

	public static void main(String[] args) throws Exception {
		// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		Table table = builder.startTable();

		// Make the header row.
		builder.insertCell();

		// Set the left indent for the table. Table wide formatting must be applied after
		// at least one row is present in the table.
		table.setLeftIndent(20.0);

		// Set height and define the height rule for the header row.
		builder.getRowFormat().setHeight(40.0);
		builder.getRowFormat().setHeightRule(HeightRule.AT_LEAST);

		// Some special features for the header row.
		builder.getCellFormat().getShading().setBackgroundPatternColor(Color.GRAY);
		builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
		builder.getFont().setSize(16);
		builder.getFont().setName("Bauhaus 93");
		builder.getFont().setBold(true);

		builder.getCellFormat().setWidth(100.0);
		builder.write("Header Row,\n Cell 1");

		// We don't need to specify the width of this cell because it's inherited from the previous cell.
		builder.insertCell();
		builder.write("Header Row,\n Cell 2");

		builder.insertCell();
		builder.getCellFormat().setWidth(200.0);
		builder.write("Header Row,\n Cell 3");
		builder.endRow();

		// Set features for the other rows and cells.
		builder.getCellFormat().getShading().setBackgroundPatternColor(Color.BLUE);
		builder.getCellFormat().setWidth(100.0);
		builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
		
		// Reset height and define a different height rule for table body
		builder.getRowFormat().setHeight(30.0);
		builder.getRowFormat().setHeight(HeightRule.AUTO);
		builder.insertCell();
		
		// Reset font formatting.
		builder.getFont().setSize(12);
		builder.getFont().setBold(false);

		// Build the other cells.
		builder.write("Row 1, Cell 1 Content");
		builder.insertCell();
		builder.write("Row 1, Cell 2 Content");

		builder.insertCell();
		builder.getCellFormat().setWidth(200.0);
		builder.write("Row 1, Cell 3 Content");
		builder.endRow();

		builder.insertCell();
		builder.getCellFormat().setWidth(100.0);
		builder.write("Row 2, Cell 1 Content");

		builder.insertCell();
		builder.write("Row 2, Cell 2 Content");

		builder.insertCell();
		builder.getCellFormat().setWidth(200.0);
		builder.write("Row 2, Cell 3 Content.");
		builder.endRow();
		builder.endTable();

		// Save the document to disk.
		doc.save(dataDir + "DocumentBuilder_CreateFormattedTable_Out.doc");
	}
}