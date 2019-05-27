package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.annotations.Test;

import java.awt.*;

public class ExImageSaveOptions extends ApiExampleBase {
    @Test
    public void useGdiEmfRenderer() throws Exception {
        //ExStart
        //ExFor:ImageSaveOptions.UseGdiEmfRenderer
        //ExSummary:Shows how to save metafiles directly without using GDI+ to EMF.
        Document doc = new Document(getMyDir() + "SaveOptions.MyraidPro.docx");

        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.EMF);
        saveOptions.setUseGdiEmfRenderer(false);

        doc.save(getArtifactsDir() + "SaveOptions.UseGdiEmfRenderer.docx", saveOptions);
        //ExEnd
    }

    @Test
    public void saveIntoGif() throws Exception {
        //ExStart
        //ExFor:ImageSaveOptions.PageIndex
        //ExSummary:Shows how to save specific document page as image file.
        Document doc = new Document(getMyDir() + "SaveOptions.MyraidPro.docx");

        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.GIF);
        //Define which page will save
        saveOptions.setPageIndex(0);

        doc.save(getArtifactsDir() + "SaveOptions.PageIndex.gif", saveOptions);
        //ExEnd
    }

    @Test
    public void qualityOptions() throws Exception {
        //ExStart
        //ExFor:GraphicsQualityOptions
        //ExFor:GraphicsQualityOptions.SmoothingMode
        //ExFor:GraphicsQualityOptions.TextRenderingHint
        //ExSummary:Shows how to set render quality options. 
        Document doc = new Document(getMyDir() + "SaveOptions.MyraidPro.docx");

        GraphicsQualityOptions qualityOptions = new GraphicsQualityOptions();
        qualityOptions.getRenderingHints().put(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        qualityOptions.getRenderingHints().put(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);

        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.JPEG);
        saveOptions.setGraphicsQualityOptions(qualityOptions);

        doc.save(getArtifactsDir() + "SaveOptions.QualityOptions.jpeg", saveOptions);
        //ExEnd
    }

    @Test(groups = "SkipMono")
    public void converImageColorsToBlackAndWhite() throws Exception {
        //ExStart
        //ExFor:ImageSaveOptions.ImageColorMode
        //ExFor:ImageSaveOptions.PixelFormat
        //ExSummary:Show how to convert document images to black and white with 1 bit per pixel
        Document doc = new Document(getMyDir() + "ImageSaveOptions.BlackAndWhite.docx");

        ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
        imageSaveOptions.setImageColorMode(ImageColorMode.BLACK_AND_WHITE);
        imageSaveOptions.setPixelFormat(ImagePixelFormat.FORMAT_1_BPP_INDEXED);

        doc.save(getArtifactsDir() + "ImageSaveOptions.BlackAndWhite.png", imageSaveOptions);
        //ExEnd
    }

    @Test
    public void thresholdForFloydSteinbergDithering() throws Exception {
        //ExStart
        //ExFor:ImageSaveOptions.ThresholdForFloydSteinbergDithering
        //ExSummary: Shows how to control the threshold for TIFF binarization in the Floyd-Steinberg method
        Document doc = new Document(getMyDir() + "ImagesSaveOptions.ThresholdForFloydSteinbergDithering.docx");

        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.TIFF);
        options.setTiffCompression(TiffCompression.CCITT_3);
        options.setImageColorMode(ImageColorMode.GRAYSCALE);
        options.setTiffBinarizationMethod(ImageBinarizationMethod.FLOYD_STEINBERG_DITHERING);
        // The default value of this property is 128. The higher value, the darker image.
        options.setThresholdForFloydSteinbergDithering((byte) 254);

        doc.save(getArtifactsDir() + "ImagesSaveOptions.ThresholdForFloydSteinbergDithering.tiff", options);
        //ExEnd
    }
}
