package com.aspose.words.examples.programming_documents.document;

import java.awt.Color;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.EmphasisMark;
import com.aspose.words.Font;
import com.aspose.words.Underline;
import com.aspose.words.examples.Utils;

public class WorkWithDocumentBuilder {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(WorkWithDocumentBuilder.class);

		CreateSimpleDocument(dataDir);
		SetFontFormatting(dataDir);
		SetFontEmphasisMark(dataDir);
	}

	private static void CreateSimpleDocument(String dataDir) throws Exception {
		// ExStart:CreateSimpleDocument
		// Open the document.
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);
		builder.write("Some text is added.");
		doc.save(dataDir + "CreateSimpleDocument_out.doc");
		// ExEnd:CreateSimpleDocument
		System.out.println("CreateSimpleDocument_out.doc at " + dataDir);
	}

	private static void SetFontFormatting(String dataDir) throws Exception {
		// ExStart: SetFontFormatting
		// Open the document.
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		// Specify font formatting before adding text.
		Font font = builder.getFont();
		font.setSize(16);
		font.setColor(Color.blue);
		font.setBold(true);
		font.setName("Arial");
		font.setUnderline(Underline.DASH);
		builder.write("Sample text.");

		doc.save(dataDir + "SetFontFormatting_out.doc");
		// ExEnd: SetFontFormatting
		System.out.println("SetFontFormatting_out.doc at " + dataDir);
	}
	
	private static void SetFontEmphasisMark(String dataDir) throws Exception {
		// ExStart: SetFontEmphasisMark
		Document document = new Document();
		DocumentBuilder builder = new DocumentBuilder(document);

		builder.getFont().setEmphasisMark(EmphasisMark.UNDER_SOLID_CIRCLE);

		builder.write("Emphasis text");
		builder.writeln();
		builder.getFont().clearFormatting();
		builder.write("Simple text");

		document.save(dataDir + "FontEmphasisMark_out.doc");
		// ExEnd: SetFontEmphasisMark
	}
}
