package com.aspose.words.examples.programming_documents.tables.ApplyFormatting;

import com.aspose.words.Document;
import com.aspose.words.HeightRule;
import com.aspose.words.LineStyle;
import com.aspose.words.NodeType;
import com.aspose.words.Row;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;

public class ApplyFormattingOnTheRowLevel {

	private static final String dataDir = Utils.getSharedDataDir(ApplyFormattingOnTheRowLevel.class) + "Tables/";

	public static void main(String[] args) throws Exception {
		Document doc = new Document(dataDir + "Table.Document.doc");
		Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

		// Retrieve the first row in the table.
		Row firstRow = table.getFirstRow();

		// Modify some row level properties.
		firstRow.getRowFormat().getBorders().setLineStyle(LineStyle.NONE);
		firstRow.getRowFormat().setHeightRule(HeightRule.AUTO);
		firstRow.getRowFormat().setAllowBreakAcrossPages(true);
	}
}
