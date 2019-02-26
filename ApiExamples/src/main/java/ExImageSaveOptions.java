//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.GraphicsQualityOptions;
import com.aspose.words.ImageColorMode;
import com.aspose.words.ImagePixelFormat;

import java.awt.*;

public class ExImageSaveOptions extends ApiExampleBase
{
    @Test
    public void useGdiEmfRenderer() throws Exception
    {
        //ExStart
        //ExFor:ImageSaveOptions.UseGdiEmfRenderer
        //ExSummary:Shows how to save metafiles directly without using GDI+ to EMF.
        Document doc = new Document(getMyDir() + "SaveOptions.MyraidPro.docx");

        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.EMF);
        saveOptions.setUseGdiEmfRenderer(false);
        //ExEnd
    }

    @Test
    public void saveIntoGif() throws Exception
        {
        //ExStart
        //ExFor:ImageSaveOptions.UseGdiEmfRenderer
        //ExSummary:Shows how to save specific document page as image file.
        Document doc = new Document(getMyDir() + "SaveOptions.MyraidPro.docx");

        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.GIF);
        //Define which page will save
        saveOptions.setPageIndex(0);

        doc.save(getMyDir() + "\\Artifacts\\SaveOptions.MyraidPro Out.gif", saveOptions);
        //ExEnd
        }

    @Test
    public void qualityOptions() throws Exception
    {
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

        doc.save(getMyDir() + "\\Artifacts\\SaveOptions.QualityOptions Out.jpeg", saveOptions);
        //ExEnd
    }

    @Test
    public void converImageColorsToBlackAndWhite() throws Exception
    {
        //ExStart
        //ExFor:ImageSaveOptions.ImageColorMode
        //ExFor:ImageSaveOptions.PixelFormat
        //ExSummary:Show how to convert document images to black and white with 1 bit per pixel
        Document doc = new Document(getMyDir() + "ImageSaveOptions.BlackAndWhite.docx");

        ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
        imageSaveOptions.setImageColorMode(ImageColorMode.BLACK_AND_WHITE);
        imageSaveOptions.setPixelFormat(ImagePixelFormat.FORMAT_1_BPP_INDEXED);
        
        doc.save(getMyDir() + "\\Artifacts\\ImageSaveOptions.BlackAndWhite Out.png", imageSaveOptions);
        //ExEnd
    }
}
