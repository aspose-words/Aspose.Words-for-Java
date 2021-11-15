package com.aspose.words.examples.rendering_printing;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class SaveAsMultipageTiff {

    private static final String dataDir = Utils.getSharedDataDir(SaveAsMultipageTiff.class) + "RenderingAndPrinting/";

    public static void main(String[] args) throws Exception {

        // Open the document.
        Document doc = new Document(dataDir + "TestFile.MultipageTIFF.docx");
        //ExStart:SaveAsTIFF
        // Save the document as multipage TIFF.
        doc.save(dataDir + "TestFile.MultipageTIFF_out.tiff");
        //ExEnd:SaveAsTIFF

        // ExStart:SaveAsTIFFUsingImageSaveOptions
        // Create an ImageSaveOptions object to pass to the Save method
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.TIFF);
        options.setPageSet(new PageSet(0, 2));
        options.setTiffCompression(TiffCompression.CCITT_4);
        options.setResolution(160);
        doc.save(dataDir + "TestFileWithOptions_Out.tiff", options);
    	// ExEnd:SaveAsTIFFUsingImageSaveOptions
        System.out.println("Document saved as multi page TIFF successfully.");
    }
}