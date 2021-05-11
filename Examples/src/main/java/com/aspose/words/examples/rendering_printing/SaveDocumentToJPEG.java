package com.aspose.words.examples.rendering_printing;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import com.aspose.words.Document;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.PageSet;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

public class SaveDocumentToJPEG {

	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(SaveDocumentToJPEG.class);

		// ExStart:SaveDocumentToJPEG
		// Open the document
		Document doc = new Document(dataDir + "Rendering.doc");
		// Save as a JPEG image file with default options
		doc.save(dataDir + "Rendering.JpegDefaultOptions.jpg");

		// Save document to stream as a JPEG with default options
		OutputStream docStream = new FileOutputStream(dataDir + "Rendering.JpegOutStream.jpg");
		doc.save(docStream, SaveFormat.JPEG);

		// Save document to a JPEG image with specified options.
		// Render the third page only and set the JPEG quality to 80%
		// In this case we need to pass the desired SaveFormat to the ImageSaveOptions
		// constructor
		// to signal what type of image to save as.
		ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.JPEG);
		imageOptions.setPageSet(new PageSet(2, 1));
		imageOptions.setJpegQuality(80);
		doc.save(dataDir + "Rendering.JpegCustomOptions.jpg", imageOptions);
		// ExEnd:SaveDocumentToJPEG
	}

}
