// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.RtfSaveOptions;
import com.aspose.words.SaveFormat;


@Test
public class ExRtfSaveOptions extends ApiExampleBase
{
    @Test
    public void exportImages() throws Exception
    {
        //ExStart
        //ExFor:RtfSaveOptions
        //ExFor:RtfSaveOptions.ExportCompactSize
        //ExFor:RtfSaveOptions.ExportImagesForOldReaders
        //ExFor:RtfSaveOptions.SaveFormat
        //ExSummary:Shows how to save a document to .rtf with custom options.
        // Open a document with images
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // Configure a RtfSaveOptions instance to make our output document more suitable for older devices
        RtfSaveOptions options = new RtfSaveOptions();
        {
            options.setSaveFormat(SaveFormat.RTF);
            options.setExportCompactSize(true);
            options.setExportImagesForOldReaders(true);
        }

        doc.save(getArtifactsDir() + "RtfSaveOptions.ExportImages.rtf", options);
        //ExEnd
    }
}
