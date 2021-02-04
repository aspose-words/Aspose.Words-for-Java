// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.words.DocumentBuilder;
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
        //ExSummary:Shows how to rasterize complex elements while saving a document to PCL.
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
    public void fallbackFontName() throws Exception
    {
        //ExStart
        //ExFor:PclSaveOptions.FallbackFontName
        //ExSummary:Shows how to declare a font that a printer will apply to printed text as a substitute should its original font be unavailable.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Non-existent font");
        builder.write("Hello world!");

        PclSaveOptions saveOptions = new PclSaveOptions();
        saveOptions.setFallbackFontName("Times New Roman");
        
        // This document will instruct the printer to apply "Times New Roman" to the text with the missing font.
        // Should "Times New Roman" also be unavailable, the printer will default to the "Arial" font.
        doc.save(getArtifactsDir() + "PclSaveOptions.SetPrinterFont.pcl", saveOptions);
        //ExEnd
    }

    @Test
    public void addPrinterFont() throws Exception
    {
        //ExStart
        //ExFor:PclSaveOptions.AddPrinterFont(string, string)
        //ExSummary:Shows how to get a printer to substitute all instances of a specific font with a different font. 
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Courier");
        builder.write("Hello world!");

        PclSaveOptions saveOptions = new PclSaveOptions();
        saveOptions.addPrinterFont("Courier New", "Courier");

        // When printing this document, the printer will use the "Courier New" font
        // to access places where our document used the "Courier" font.
        doc.save(getArtifactsDir() + "PclSaveOptions.AddPrinterFont.pcl", saveOptions);
        //ExEnd
    }

    @Test (description = "This test is a manual check that PaperTray information is preserved in the output pcl document.")
    public void getPreservedPaperTrayInformation() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Paper tray information is now preserved when saving document to PCL format.
        // Following information is transferred from document's model to PCL file.
        for (Section section : doc.getSections().<Section>OfType() !!Autoporter error: Undefined expression type )
        {
            section.getPageSetup().setFirstPageTray(15);
            section.getPageSetup().setOtherPagesTray(12);
        }

        doc.save(getArtifactsDir() + "PclSaveOptions.GetPreservedPaperTrayInformation.pcl");
    }
}
