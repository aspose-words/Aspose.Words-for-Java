package com.aspose.words.examples.programming_documents.tables.ApplyFormatting;

import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.Table;
import com.aspose.words.TextWrapping;
import com.aspose.words.VerticalAlignment;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 10/13/2017.
 */
public class GetTablePosition {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(GetTablePosition.class);

		getTablePosition(dataDir);
		getFloatingTablePosition(dataDir);
		setFloatingTablePosition(dataDir);
		SetRelativeHorizontalOrVerticalPosition(dataDir);
	}

	private static void getTablePosition(String dataDir) throws Exception {
		// ExStart:GetTablePosition
		Document doc = new Document(dataDir + "Table.Document.doc");

		// Retrieve the first table in the document.
		Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

		if (table.getTextWrapping() == TextWrapping.AROUND) {
			System.out.println("RelativeHorizontalAlignment: " + table.getRelativeHorizontalAlignment());
			System.out.println("RelativeVerticalAlignment: " + table.getRelativeVerticalAlignment());
		} else {
			System.out.println("Alignment: " + table.getAlignment());
		}

		// ExEnd:GetTablePosition
		System.out.println("\nTable position fetched successfully.");
	}

	private static void getFloatingTablePosition(String dataDir) throws Exception {
		// ExStart:GetFloatingTablePosition
		Document doc = new Document(dataDir + "FloatingTablePosition.docx");
		for (Table table : doc.getFirstSection().getBody().getTables()) {
			// If table is floating type then print its positioning properties.
			if (table.getTextWrapping() == TextWrapping.AROUND) {
				System.out.println(table.getHorizontalAnchor());
				System.out.println(table.getVerticalAnchor());
				System.out.println(table.getAbsoluteHorizontalDistance());
				System.out.println(table.getAbsoluteVerticalDistance());
				System.out.println(table.getAllowOverlap());
				System.out.println("..............................");
			}
		}

		// ExEnd:GetFloatingTablePosition
		System.out.println("\nTable position fetched successfully.");
	}

	private static void setFloatingTablePosition(String dataDir) throws Exception {
		// ExStart:SetFloatingTablePosition
		Document doc = new Document(dataDir + "FloatingTablePosition.docx");
		Table table = doc.getFirstSection().getBody().getTables().get(0);

		// sets absolute table horizontal position at 10pt.
		table.setAbsoluteHorizontalDistance(10);

		// sets vertical table position to center of entity specified by
		// Table.VerticalAnchor.
		table.setRelativeVerticalAlignment(VerticalAlignment.CENTER);

		// Save the document to disk.
		doc.save(dataDir + "Table.SetFloatingTablePosition.docx");

		// ExEnd:SetFloatingTablePosition
		System.out.println("\nTable position fetched successfully.");
	}

	private static void SetRelativeHorizontalOrVerticalPosition(String dataDir) throws Exception {
		// ExStart:SetRelativeHorizontalOrVerticalPosition
		Document doc = new Document(dataDir + "FloatingTablePosition.docx");
		Table table = doc.getFirstSection().getBody().getTables().get(0);

		table.setHorizontalAnchor(RelativeHorizontalPosition.COLUMN);
		table.setVerticalAnchor(RelativeVerticalPosition.PAGE);

		// Save the document to disk.
		doc.save(dataDir + "Table.SetFloatingTablePosition.docx");
		// ExEnd:SetRelativeHorizontalOrVerticalPosition
		System.out.println("\nSet the Table position successfully.");
	}
}
