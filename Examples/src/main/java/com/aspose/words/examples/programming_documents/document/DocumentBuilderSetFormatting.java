package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.Paragraph;
import com.aspose.words.ParagraphFormat;
import com.aspose.words.examples.Utils;

public class DocumentBuilderSetFormatting {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getDataDir(DocumentBuilderSetFormatting.class);

		setAsianTypographyLinebreakGroupProp(dataDir);
		ChangeAsianParagraphSpacingandIndents(dataDir);
		SetSnapToGrid(dataDir);
	}

	public static void setAsianTypographyLinebreakGroupProp(String dataDir) throws Exception {
		// ExStart:SetAsianTypographyLinebreakGroupProp
		Document doc = new Document(dataDir + "Input.docx");

		ParagraphFormat format = doc.getFirstSection().getBody().getParagraphs().get(0).getParagraphFormat();
		format.setFarEastLineBreakControl(false);
		format.setWordWrap(true);
		format.setHangingPunctuation(false);

		dataDir = dataDir + "SetAsianTypographyLinebreakGroupProp_out.doc";
		doc.save(dataDir);
		// ExEnd:SetAsianTypographyLinebreakGroupProp
		System.out.println(
				"\nParagraphFormat properties for Asian Typography line break group are set successfully.\nFile saved at "
						+ dataDir);
	}

	public static void ChangeAsianParagraphSpacingandIndents(String dataDir) throws Exception {
		// ExStart:ChangeAsianParagraphSpacingandIndents
		Document doc = new Document(dataDir + "Input.docx");

		ParagraphFormat format = doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat();
		format.setCharacterUnitLeftIndent(10); // ParagraphFormat.LeftIndent will be updated
		format.setCharacterUnitRightIndent(10); // ParagraphFormat.RightIndent will be updated
		format.setCharacterUnitFirstLineIndent(20); // ParagraphFormat.FirstLineIndent will be updated
		format.setLineUnitBefore(5); // ParagraphFormat.SpaceBefore will be updated
		format.setLineUnitAfter(10); // ParagraphFormat.SpaceAfter will be updated

		dataDir = dataDir + "ChangeAsianParagraphSpacingandIndents_out.doc";
		doc.save(dataDir);
		// ExEnd:ChangeAsianParagraphSpacingandIndents
		System.out.println("\nSpacing and Indents applied successfully to paragraph.\nFile saved at " + dataDir);
	}

	public static void SetSnapToGrid(String dataDir) throws Exception {
		// ExStart:SetSnapToGrid
		Document doc = new Document(dataDir);

		Paragraph par = doc.getFirstSection().getBody().getFirstParagraph();
		par.getParagraphFormat().setSnapToGrid(true);
		par.getRuns().get(0).getFont().setSnapToGrid(true);

		dataDir = dataDir + "SetSnapToGrid_out.doc";
		doc.save(dataDir);
		// ExEnd:SetSnapToGrid
		System.out.println("\nSetSnapToGrid successfully to paragraph.\nFile saved at " + dataDir);
	}
}
