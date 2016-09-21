package com.aspose.words.examples.programming_documents.tables.ColumnsAndRows;

import com.aspose.words.AutoFitBehavior;
import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;

public class ApplyAutoFitSettingsToATable {
	
	private static final String dataDir = Utils.getSharedDataDir(ApplyAutoFitSettingsToATable.class) + "Tables/";
	
	public static void main(String[] args) throws Exception {
		// Auto fits a table to fit the page width
		autoFittingATableToWindow();
		
		// Auto fits a table in the document to its contents
		autoFittingATableToContents();
		
		// Disabling AutoFitting on a Table and Use Fixed Column Widths
		disablingAutoFittingOnATableAndUseFixedColumnWidths();
	}

	public static void autoFittingATableToWindow() throws Exception {
		// Open the document
		Document doc = new Document(dataDir + "TestFile.doc");
		Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

		// Auto fit the first table to the page width.
		table.autoFit(AutoFitBehavior.AUTO_FIT_TO_WINDOW);

		// Save the document to disk.
		doc.save(dataDir + "TestFile.AutoFitToWindow Out.doc");
	}
	
	public static void autoFittingATableToContents() throws Exception {
		// Open the document
		Document doc = new Document(dataDir + "TestFile.doc");
		Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);

		// Auto fit the table to the cell contents
		table.autoFit(AutoFitBehavior.AUTO_FIT_TO_CONTENTS);

		// Save the document to disk.
		doc.save(dataDir + "TestFile.AutoFitToContents Out.doc");
	}
	
	public static void disablingAutoFittingOnATableAndUseFixedColumnWidths() throws Exception {
		// Open the document
		Document doc = new Document(dataDir + "TestFile.doc");
		Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);

		// Disable autofitting on this table.
		table.autoFit(AutoFitBehavior.FIXED_COLUMN_WIDTHS);

		// Save the document to disk.
		doc.save(dataDir + "TestFile.FixedWidth Out.doc");
	}
}
