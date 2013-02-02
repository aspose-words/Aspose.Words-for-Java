//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package SaveAsMultipageTiff;

import com.aspose.words.Document;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.TiffCompression;

import java.io.File;
import java.net.URI;

class Program
{
    public static void main(String[] args) throws Exception
    {
        // Sample infrastructure.
        URI exeDir = Program.class.getResource("").toURI();
        String dataDir = new File(exeDir.resolve("../../Data")) + File.separator;

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