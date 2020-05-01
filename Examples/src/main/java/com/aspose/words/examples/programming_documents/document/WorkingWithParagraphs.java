package com.aspose.words.examples.programming_documents.document;

import java.awt.Color;

import com.aspose.words.BorderCollection;
import com.aspose.words.BorderType;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Font;
import com.aspose.words.LineStyle;
import com.aspose.words.Paragraph;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.ParagraphFormat;
import com.aspose.words.Shading;
import com.aspose.words.Style;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.StyleType;
import com.aspose.words.TextureIndex;
import com.aspose.words.examples.Utils;

public class WorkingWithParagraphs {

	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(WorkingWithParagraphs.class);

		DocumentBuilderInsertParagraph(dataDir);
		DocumentBuilderSetParagraphFormatting(dataDir);
		DocumentBuilderSetSpaceBetweenAsianAndLatinText(dataDir);
		setAsianTypographyLinebreakGroupProp(dataDir);
		DocumentBuilderApplyParagraphStyle(dataDir);
		ParagraphInsertStyleSeparator(dataDir);
		DocumentBuilderApplyBordersAndShadingToParagraph(dataDir);
		ChangeAsianParagraphSpacingandIndents(dataDir);
		SetSnapToGrid(dataDir);
	}

	private static void DocumentBuilderInsertParagraph(String dataDir) throws Exception {

		// ExStart:DocumentBuilderInsertParagraph
		// Open the document.
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);
		Font font = builder.getFont();
		font.setSize(16);
		font.setColor(Color.DARK_GRAY);
		font.setBold(true);
		font.setName("Algerian");
		font.setUnderline(2);

		ParagraphFormat paragraphFormat = builder.getParagraphFormat();
		paragraphFormat.setFirstLineIndent(12);
		paragraphFormat.setAlignment(1);
		paragraphFormat.setKeepTogether(true);

		builder.write("This is a sample Paragraph");
		doc.save(dataDir + "InsertParagraph_out.doc");
		// ExEnd:DocumentBuilderInsertParagraph
	}

	private static void DocumentBuilderSetParagraphFormatting(String dataDir) throws Exception {
		// ExStart:DocumentBuilderSetParagraphFormatting
		// Open the document.
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		ParagraphFormat paragraphFormat = builder.getParagraphFormat();
		paragraphFormat.setAlignment(ParagraphAlignment.CENTER);
		paragraphFormat.setLeftIndent(50);
		paragraphFormat.setRightIndent(50);
		paragraphFormat.setSpaceAfter(25);
		paragraphFormat.setKeepTogether(true);

		builder.writeln(
				"I'm a very nice formatted paragraph. I'm intended to demonstrate how the left and right indents affect word wrapping.");
		builder.writeln(
				"I'm another nice formatted paragraph. I'm intended to demonstrate how the space after paragraph looks like.");
		doc.save(dataDir + "SetParagraphFormatting_out.doc");
		// ExEnd:DocumentBuilderSetParagraphFormatting
	}

	private static void DocumentBuilderSetSpaceBetweenAsianAndLatinText(String dataDir) throws Exception {
		// ExStart:DocumentBuilderSetSpaceBetweenAsianAndLatinText
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		// Set paragraph formatting properties
		ParagraphFormat paragraphFormat = builder.getParagraphFormat();
		paragraphFormat.setAddSpaceBetweenFarEastAndAlpha(true);

		paragraphFormat.setAddSpaceBetweenFarEastAndDigit(true);

		builder.writeln("Automatically adjust space between Asian and Latin text");
		builder.writeln("Automatically adjust space between Asian text and numbers");

		dataDir = dataDir + "DocumentBuilderSetSpacebetweenAsianandLatintext_out.doc";
		doc.save(dataDir);
		// ExEnd:DocumentBuilderSetSpaceBetweenAsianAndLatinText
	}

	private static void setAsianTypographyLinebreakGroupProp(String dataDir) throws Exception {
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

	private static void DocumentBuilderApplyParagraphStyle(String dataDir) throws Exception {

		// ExStart:DocumentBuilderApplyParagraphStyle
		// Open the document.
		Document doc = new Document();

		// Set paragraph style
		DocumentBuilder builder = new DocumentBuilder(doc);
		builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.TITLE);
		builder.write("Hello");

		doc.save(dataDir + "ApplyParagraphStyle_out.doc");
		// ExEnd:DocumentBuilderApplyParagraphStyle
	}

	private static void ParagraphInsertStyleSeparator(String dataDir) throws Exception {
		// ExStart:ParagraphInsertStyleSeparator
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		Style paraStyle = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "MyParaStyle");
		paraStyle.getFont().setBold(false);
		paraStyle.getFont().setSize(8);
		paraStyle.getFont().setName("Arial");

		// Append text with "Heading 1" style.
		builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
		builder.write("Heading 1");
		builder.insertStyleSeparator();

		// Append text with another style.
		builder.getParagraphFormat().setStyleName(paraStyle.getName());
		builder.write("This is text with some other formatting ");

		dataDir = dataDir + "InsertStyleSeparator_out.doc";
		doc.save(dataDir);
		// ExEnd:ParagraphInsertStyleSeparator

		System.out.println(
				"\nApplied different paragraph styles to two different parts of a text line successfully.\nFile saved at "
						+ dataDir);
	}


	private static void DocumentBuilderApplyBordersAndShadingToParagraph(String dataDir) throws Exception {

		// ExStart:DocumentBuilderApplyBordersAndShadingToParagraph
		// Open the document.
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		// Set paragraph borders
		BorderCollection borders = builder.getParagraphFormat().getBorders();
		borders.setDistanceFromText(20);
		borders.getByBorderType(BorderType.LEFT).setLineStyle(LineStyle.DOUBLE);
		borders.getByBorderType(BorderType.RIGHT).setLineStyle(LineStyle.DOUBLE);
		borders.getByBorderType(BorderType.TOP).setLineStyle(LineStyle.DOUBLE);
		borders.getByBorderType(BorderType.BOTTOM).setLineStyle(LineStyle.DOUBLE);
		// Set paragraph shading
		Shading shading = builder.getParagraphFormat().getShading();
		shading.setTexture(TextureIndex.TEXTURE_DIAGONAL_CROSS);
		shading.setBackgroundPatternColor(Color.YELLOW);
		shading.setForegroundPatternColor(Color.GREEN);

		builder.write("I'm a formatted paragraph with double border and nice shading.");
		doc.save(dataDir + "ApplyBordersAndShading_out.doc");
		// ExEnd:DocumentBuilderApplyBordersAndShadingToParagraph
	}

	private static void ChangeAsianParagraphSpacingandIndents(String dataDir) throws Exception {
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

	private static void SetSnapToGrid(String dataDir) throws Exception {
		// ExStart:SetSnapToGrid
		Document doc = new Document(dataDir + "Input.docx");

		Paragraph par = doc.getFirstSection().getBody().getFirstParagraph();
		par.getParagraphFormat().setSnapToGrid(true);
		par.getRuns().get(0).getFont().setSnapToGrid(true);

		dataDir = dataDir + "SetSnapToGrid_out.doc";
		doc.save(dataDir);
		// ExEnd:SetSnapToGrid
		System.out.println("\nSetSnapToGrid successfully to paragraph.\nFile saved at " + dataDir);
	}

}
