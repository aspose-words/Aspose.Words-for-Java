package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.PageSet;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

public class ConvertToImage {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		String dataDir = Utils.getDataDir(ConvertToHTML.class);

		SavePdfToJpeg(dataDir);
		ConvertDocumentToImage(dataDir);
	}

	public static void ConvertDocxToJpeg(String dataDir) throws Exception
	{
		// ExStart:ConvertDocxToJpeg
		// Load the document from disk.
		Document doc = new Document(dataDir + "TestDoc.pdf");

		// Save the document in JPEG format.
		doc.save(dataDir + "SaveDocx2Jpeg.jpeg");
		// ExEnd:ConvertDocxToJpeg
	}

	public static void SavePdfToJpeg(String dataDir) throws Exception
	{
		// ExStart:SavePdfToJpeg
		// Load the document from disk.
		Document doc = new Document(dataDir + "TestDoc.pdf");

		// Save the document in JPEG format.
		doc.save(dataDir + "SavePdf2Jpeg.jpeg");
		// ExEnd:SavePdfToJpeg
	}
	
	public static void ConvertDocumentToImage(String dataDir) throws Exception
	{
		// ExStart:ConvertDocumentToImage
		// Load the document from disk.
		Document doc = new Document(dataDir + "TestDoc.docx");

		ImageSaveOptions options = new ImageSaveOptions(SaveFormat.JPEG);

		// Set the "PageSet" to "0" to convert only the first page of a document.
		options.setPageSet(new PageSet(0));

		// Change the image's brightness and contrast.
		// Both are on a 0-1 scale and are at 0.5 by default.
		options.setImageBrightness(0.3f);
		options.setImageContrast(0.7f);

		// Change the horizontal resolution.
		// The default value for these properties is 96.0, for a resolution of 96dpi.
		options.setHorizontalResolution(72f);

		// Save the document in JPEG format.
		doc.save(dataDir + "SaveDocx2Jpeg.jpeg", options);
		// ExEnd:ConvertDocumentToImage
	}
}
