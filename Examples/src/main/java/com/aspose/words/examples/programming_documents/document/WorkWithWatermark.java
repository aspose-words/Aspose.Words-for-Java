package com.aspose.words.examples.programming_documents.document;

import java.awt.Color;
import java.awt.Image;
import java.io.File;
import java.io.IOException;

import javax.imageio.ImageIO;
import javax.imageio.stream.ImageInputStream;

import com.aspose.words.Document;
import com.aspose.words.ImageWatermarkOptions;
import com.aspose.words.TextWatermarkOptions;
import com.aspose.words.WatermarkLayout;
import com.aspose.words.WatermarkType;
import com.aspose.words.examples.Utils;

public class WorkWithWatermark {

	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(WorkWithWatermark.class);

		AddTextWatermarkWithSpecificOptions(dataDir);
		AddImageWatermarkWithSpecificOptions(dataDir);
		RemoveWatermarkFromDocument(dataDir);
	}

	public static void AddTextWatermarkWithSpecificOptions(String dataDir) throws Exception {
		// ExStart: AddTextWatermarkWithSpecificOptions
		Document doc = new Document(dataDir + "Document.doc");

		TextWatermarkOptions options = new TextWatermarkOptions();
		options.setFontFamily("Arial");
		options.setFontSize(36);
		options.setColor(Color.BLACK);
		options.setLayout(WatermarkLayout.HORIZONTAL);
		options.isSemitrasparent(false);

		doc.getWatermark().setText("Test", options);

		doc.save(dataDir + "AddTextWatermark_out.docx");
		// ExEnd: AddTextWatermarkWithSpecificOptions
		System.out.println("\nDocument saved successfully.\nFile saved at " + dataDir);
	}

	public static void AddImageWatermarkWithSpecificOptions(String dataDir) throws Exception {
		// ExStart: AddImageWatermarkWithSpecificOptions
		Document doc = new Document(dataDir + "Document.doc");

		ImageWatermarkOptions options = new ImageWatermarkOptions();
		options.setScale(5);
		options.isWashout(false);

		ImageInputStream stream = ImageIO.createImageInputStream(new File(dataDir + "Watermark.png"));
		doc.getWatermark().setImage(ImageIO.read(stream), options);

		doc.save(dataDir + "AddImageWatermark_out.docx");
		// ExEnd: AddImageWatermarkWithSpecificOptions
		System.out.println("\nDocument saved successfully.\nFile saved at " + dataDir);
	}

	public static void RemoveWatermarkFromDocument(String dataDir) throws Exception {
		// ExStart: RemoveWatermarkFromDocument
		Document doc = new Document(dataDir + "AddTextWatermark_out.docx");

		if (doc.getWatermark().getType() == WatermarkType.TEXT) {
			doc.getWatermark().remove();
		}

		doc.save(dataDir + "RemoveWatermark_out.docx");
		// ExEnd: RemoveWatermarkFromDocument
		System.out.println("\nDocument saved successfully.\nFile saved at " + dataDir);
	}

}
