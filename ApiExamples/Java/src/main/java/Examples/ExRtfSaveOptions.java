package Examples;

// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

@Test
public class ExRtfSaveOptions extends ApiExampleBase {
    @Test
    public void exportImages() throws Exception {
        //ExStart
        //ExFor:RtfSaveOptions
        //ExFor:RtfSaveOptions.ExportCompactSize
        //ExFor:RtfSaveOptions.ExportImagesForOldReaders
        //ExFor:RtfSaveOptions.SaveFormat
        //ExSummary:Shows how to save a document to .rtf with custom options.
        // Open a document with images
        Document doc = new Document(getMyDir() + "Rendering.docx");

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

    @Test(enabled = false)
    public void saveImagesAsWmf() throws Exception {
        //ExStart
        //ExFor:RtfSaveOptions.SaveImagesAsWmf
        //ExSummary:Shows how to save all images as Wmf when saving to the Rtf document.
        // Open a document that contains images in the jpeg format
        Document doc = new Document(getMyDir() + "Images.docx");

        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);
        Shape shapeWithJpg = (Shape) shapes.get(0);
        Assert.assertEquals(shapeWithJpg.getImageData().getImageType(), ImageType.JPEG);

        RtfSaveOptions rtfSaveOptions = new RtfSaveOptions();
        rtfSaveOptions.setSaveImagesAsWmf(true);
        doc.save(getArtifactsDir() + "RtfSaveOptions.SaveImagesAsWmf.rtf", rtfSaveOptions);

        doc = new Document(getArtifactsDir() + "RtfSaveOptions.SaveImagesAsWmf.rtf");

        shapes = doc.getChildNodes(NodeType.SHAPE, true);
        Shape shapeWithWmf = (Shape) shapes.get(0);
        Assert.assertEquals(shapeWithWmf.getImageData().getImageType(), ImageType.WMF);
        //ExEnd
    }
}
