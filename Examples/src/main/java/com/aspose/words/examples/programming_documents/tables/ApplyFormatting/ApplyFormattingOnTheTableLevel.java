package com.aspose.words.examples.programming_documents.tables.ApplyFormatting;

import java.awt.Color;

import com.aspose.words.BorderType;
import com.aspose.words.Document;
import com.aspose.words.LineStyle;
import com.aspose.words.NodeType;
import com.aspose.words.Table;
import com.aspose.words.TableAlignment;
import com.aspose.words.TextureIndex;
import com.aspose.words.examples.Utils;

public class ApplyFormattingOnTheTableLevel {
	
	private static final String dataDir = Utils.getSharedDataDir(ApplyFormattingOnTheTableLevel.class) + "Tables/";
	
	public static void main(String[] args) throws Exception {
		// Apply a outline border to a table
		applyOutlineBorderToATable();
		
		// Build a table with all borders enabled (grid)
		buildATableWithAllBordersEnabled();
	}

	public static void applyOutlineBorderToATable() throws Exception {
		Document doc = new Document(dataDir + "Table.EmptyTable.doc");
		Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

		// Align the table to the center of the page.
		table.setAlignment(TableAlignment.CENTER);

		// Clear any existing borders from the table.
		table.clearBorders();

		// Set a green border around the table but not inside.
		table.setBorder(BorderType.LEFT, LineStyle.SINGLE, 1.5, Color.GREEN, true);
		table.setBorder(BorderType.RIGHT, LineStyle.SINGLE, 1.5, Color.GREEN, true);
		table.setBorder(BorderType.TOP, LineStyle.SINGLE, 1.5, Color.GREEN, true);
		table.setBorder(BorderType.BOTTOM, LineStyle.SINGLE, 1.5, Color.GREEN, true);

		// Fill the cells with a light green solid color.
		table.setShading(TextureIndex.TEXTURE_SOLID, Color.GREEN, Color.GREEN);

		doc.save(dataDir + "Table.SetOutlineBorders_Out.doc");
	}
	
	public static void buildATableWithAllBordersEnabled() throws Exception {
		Document doc = new Document(dataDir + "Table.EmptyTable.doc");
		Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);

		// Clear any existing borders from the table.
		table.clearBorders();

		// Set a green border around and inside the table.
		table.setBorders(LineStyle.SINGLE, 1.5, Color.GREEN);

		doc.save(dataDir + "Table.SetAllBorders Out.doc");
	}
}
