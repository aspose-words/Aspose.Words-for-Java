
package com.aspose.words.examples.rendering_printing;

import com.aspose.words.Document;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.TiffCompression;
import com.aspose.words.examples.Utils;

public class SaveAsMultipageTiff {
	
	private static final String dataDir = Utils.getSharedDataDir(SaveAsMultipageTiff.class) + "RenderingAndPrinting/";
	
	public static void main(String[] args) throws Exception {
		
		// Open the document.
		Document doc = new Document(dataDir + "TestFile.doc");
		// Save the document as multipage TIFF.
		doc.save(dataDir + "TestFile Out.tiff");
		
		//Create an ImageSaveOptions object to pass to the Save method
		ImageSaveOptions options = new ImageSaveOptions(SaveFormat.TIFF);
		options.setPageIndex(0);
		options.setPageCount(2);
		options.setTiffCompression(TiffCompression.CCITT_4);
		options.setResolution(160);
		doc.save(dataDir + "TestFileWithOptions Out.tiff", options);
		
		System.out.println("Document saved as multi page TIFF successfully.");
	}
}