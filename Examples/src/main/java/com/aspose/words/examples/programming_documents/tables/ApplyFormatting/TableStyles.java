package com.aspose.words.examples.programming_documents.tables.ApplyFormatting;

import java.awt.Color;

import com.aspose.words.AutoFitBehavior;
import com.aspose.words.Cell;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.NodeType;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.Table;
import com.aspose.words.TableStyleOptions;
import com.aspose.words.examples.Utils;

public class TableStyles {

	private static final String dataDir = Utils.getSharedDataDir(TableStyles.class) + "Tables/";

	public static void main(String[] args) throws Exception {
		applyATableStyle();
		
		expandFormattingFromStylesOnToRowsAndCells();
	}

	public static void applyATableStyle() throws Exception {
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		Table table = builder.startTable();
		// We must insert at least one row first before setting any table formatting.
		builder.insertCell();
		// Set the table style used based of the unique style identifier.
		// Note that not all table styles are available when saving as .doc format.
		table.setStyleIdentifier(StyleIdentifier.MEDIUM_SHADING_1_ACCENT_1);
		// Apply which features should be formatted by the style.
		table.setStyleOptions(TableStyleOptions.FIRST_COLUMN | TableStyleOptions.ROW_BANDS | TableStyleOptions.FIRST_ROW);
		table.autoFit(AutoFitBehavior.AUTO_FIT_TO_CONTENTS);

		// Continue with building the table as normal.
		builder.writeln("Item");
		builder.getCellFormat().setRightPadding(40);
		builder.insertCell();
		builder.writeln("Quantity (kg)");
		builder.endRow();

		builder.insertCell();
		builder.writeln("Apples");
		builder.insertCell();
		builder.writeln("20");
		builder.endRow();

		builder.insertCell();
		builder.writeln("Bananas");
		builder.insertCell();
		builder.writeln("40");
		builder.endRow();

		builder.insertCell();
		builder.writeln("Carrots");
		builder.insertCell();
		builder.writeln("50");
		builder.endRow();

		doc.save(dataDir + "DocumentBuilder.SetTableStyle Out.docx");
	}

	public static void expandFormattingFromStylesOnToRowsAndCells() throws Exception {
		Document doc = new Document(dataDir + "Table.TableStyle.docx");

		// Get the first cell of the first table in the document.
		Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
		Cell firstCell = table.getFirstRow().getFirstCell();

		// First print the color of the cell shading. This should be empty as the current shading
		// is stored in the table style.
		Color cellShadingBefore = firstCell.getCellFormat().getShading().getBackgroundPatternColor();
		System.out.println("Cell shading before style expansion: " + cellShadingBefore);

		// Expand table style formatting to direct formatting.
		doc.expandTableStylesToDirectFormatting();

		// Now print the cell shading after expanding table styles. A blue background pattern color
		// should have been applied from the table style.
		Color cellShadingAfter = firstCell.getCellFormat().getShading().getBackgroundPatternColor();
		System.out.println("Cell shading after style expansion: " + cellShadingAfter);
	}
}
