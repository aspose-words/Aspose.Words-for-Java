/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package renderingandprinting.renderingtoimage.saveasmultipagetiff.java;

import com.aspose.words.*;

import java.io.File;
import java.net.URI;

public class SaveAsMultipageTiff
{
    public static void main(String[] args) throws Exception
    {
            // The path to the documents directory.
        String dataDir = "src/renderingandprinting/renderingtoimage/saveasmultipagetiff/data/";

        // Open the document.
        Document doc = new Document(dataDir + "TestFile.doc");

        //ExStart
        //ExId:SaveAsMultipageTiff_save
        //ExSummary:Convert document to TIFF.
        // Save the document as multipage TIFF.
        doc.save(dataDir + "TestFile Out.tiff");
        //ExEnd

        //ExStart
        //ExId:SaveAsMultipageTiff_SaveWithOptions
        //ExSummary:Convert to TIFF using customized options
        //Create an ImageSaveOptions object to pass to the Save method
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.TIFF);
        options.setPageIndex(0);
        options.setPageCount(2);
        options.setTiffCompression(TiffCompression.CCITT_4);
        options.setResolution(160);

        doc.save(dataDir + "TestFileWithOptions Out.tiff", options);
        //ExEnd
    }
}