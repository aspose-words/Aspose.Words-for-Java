//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.PclSaveOptions;

public class ExPclSaveOptions extends ApiExampleBase
{
    @Test
    public void rasterizeElements() throws Exception
    {
        //ExStart
        //ExFor:PclSaveOptions
        //ExFor:PclSaveOptions.RasterizeTransformedElements
        //ExSummary:Shows how rasterized or not transformed elements before saving.
        Document doc = new Document(getMyDir() + "Document.EpubConversion.doc");

        PclSaveOptions saveOptions = new PclSaveOptions();
        saveOptions.setRasterizeTransformedElements(true);

        doc.save(getMyDir() + "\\Artifacts\\Document.EpubConversion.pcl", saveOptions);
        //ExEnd
    }

    @Test
    public void setPrinterFont() throws Exception
    {
        //ExStart
        //ExFor:PclSaveOptions.AddPrinterFont(string, string)
        //ExFor:PclSaveOptions.FallbackFontName
        //ExSummary:Shows how to add information about font that is uploaded to the printer and set the font that will be used if no expected font is found in printer and built-in fonts collections.
        Document doc = new Document(getMyDir() + "Document.EpubConversion.doc");

        PclSaveOptions saveOptions = new PclSaveOptions();
        saveOptions.addPrinterFont("Courier", "Courier");
        saveOptions.setFallbackFontName("Times New Roman");

        doc.save(getMyDir() + "\\Artifacts\\Document.EpubConversion.pcl", saveOptions);
        //ExEnd
    }
}
