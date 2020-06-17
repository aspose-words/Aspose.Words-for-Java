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
import com.aspose.words.PclSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.Section;


@Test
class ExPclSaveOptions !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void rasterizeElements() throws Exception
    {
        //ExStart
        //ExFor:PclSaveOptions
        //ExFor:PclSaveOptions.SaveFormat
        //ExFor:PclSaveOptions.RasterizeTransformedElements
        //ExSummary:Shows how to set whether or not to rasterize complex elements before saving.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PclSaveOptions saveOptions = new PclSaveOptions();
        {
            saveOptions.setSaveFormat(SaveFormat.PCL);
            saveOptions.setRasterizeTransformedElements(true);
        }

        doc.save(getArtifactsDir() + "PclSaveOptions.RasterizeElements.pcl", saveOptions);
        //ExEnd
    }

    @Test
    public void setPrinterFont() throws Exception
    {
        //ExStart
        //ExFor:PclSaveOptions.AddPrinterFont(string, string)
        //ExFor:PclSaveOptions.FallbackFontName
        //ExSummary:Shows how to add information about font that is uploaded to the printer and set the font that will be used if no expected font is found in printer and built-in fonts collections.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PclSaveOptions saveOptions = new PclSaveOptions();
        saveOptions.addPrinterFont("Courier", "Courier");
        saveOptions.setFallbackFontName("Times New Roman");

        doc.save(getArtifactsDir() + "PclSaveOptions.SetPrinterFont.pcl", saveOptions);
        //ExEnd
    }

    @Test (enabled = false, description = "This test is manual check that PaperTray information are preserved in pcl document.")
    public void getPreservedPaperTrayInformation() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Paper tray information is now preserved when saving document to PCL format
        // Following information is transferred from document's model to PCL file
        for (Section section : doc.getSections().<Section>OfType() !!Autoporter error: Undefined expression type )
        {
            section.getPageSetup().setFirstPageTray(15);
            section.getPageSetup().setOtherPagesTray(12);
        }

        doc.save(getArtifactsDir() + "PclSaveOptions.GetPreservedPaperTrayInformation.pcl");
    }
}
