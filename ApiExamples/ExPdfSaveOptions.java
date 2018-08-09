//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package Examples;

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.SaveFormat;
import org.testng.Assert;
import com.aspose.words.DmlRenderingMode;
import com.aspose.words.PdfImageCompression;
import com.aspose.words.PdfCompliance;
import com.aspose.words.ColorMode;
import com.aspose.words.SaveOptions;
import com.aspose.words.MetafileRenderingOptions;
import com.aspose.words.MetafileRenderingMode;
import com.aspose.words.IWarningCallback;
import com.aspose.words.WarningInfo;
import com.aspose.words.WarningType;
import com.aspose.words.WarningInfoCollection;


@Test
public class ExPdfSaveOptions extends ApiExampleBase
{
    @Test
    public void createMissingOutlineLevels() throws Exception
    {
        //ExStart
        //ExFor:OutlineOptions.CreateMissingOutlineLevels
        //ExSummary:Shows how to create missing outline levels saving the document in PDF
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Creating TOC entries
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);

        builder.writeln("Heading 1");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_4);

        builder.writeln("Heading 1.1.1.1");
        builder.writeln("Heading 1.1.1.2");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_9);

        builder.writeln("Heading 1.1.1.1.1.1.1.1.1");
        builder.writeln("Heading 1.1.1.1.1.1.1.1.2");

        // Create "PdfSaveOptions" with some mandatory parameters
        // "HeadingsOutlineLevels" specifies how many levels of headings to include in the document outline
        // "CreateMissingOutlineLevels" determining whether or not to create missing heading levels
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        pdfSaveOptions.getOutlineOptions().setHeadingsOutlineLevels(9);
        pdfSaveOptions.getOutlineOptions().setCreateMissingOutlineLevels(true);
        pdfSaveOptions.setSaveFormat(SaveFormat.PDF);

        doc.save(getMyDir() + "\\Artifacts\\CreateMissingOutlineLevels.pdf", pdfSaveOptions);
        //ExEnd
    }

    //Note: Test doesn't contain validation result.
    //For validation result, you can add some shapes to the document and assert, that the DML shapes are render correctly
    @Test
    public void drawingMl() throws Exception
    {
        //ExStart
        //ExFor:DmlRenderingMode
        //ExFor:SaveOptions.DmlRenderingMode
        //ExSummary:Shows how to define rendering for DML shapes
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        pdfSaveOptions.setDmlRenderingMode(DmlRenderingMode.DRAWING_ML);

        doc.save(getMyDir() + "\\Artifacts\\DrawingMl.pdf", pdfSaveOptions);
        //ExEnd
    }

    @Test
    public void withoutUpdateFields() throws Exception
    {
        //ExStart
        //ExFor:SaveOptions.UpdateFields
        //ExSummary:Shows how to update fields before saving into a PDF document.
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        pdfSaveOptions.setUpdateFields(false);

        doc.save(getMyDir() + "\\Artifacts\\UpdateFields_False.pdf", pdfSaveOptions);
        //ExEnd
    }

    @Test
    public void withUpdateFields() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        pdfSaveOptions.setUpdateFields(true);

        doc.save(getMyDir() + "\\Artifacts\\UpdateFields_False.pdf", pdfSaveOptions);
    }

    //ToDo: Add gold asserts for PDF files
    // For assert this test you need to open "SaveOptions.PdfImageCompression PDF_A_1_B Out.pdf" and "SaveOptions.PdfImageCompression PDF_A_1_A Out.pdf" 
    // and check that header image in this documents are equal header image in the "SaveOptions.PdfImageComppression Out.pdf" 
    @Test
    public void imageCompression() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.Compliance
        //ExFor:PdfSaveOptions.ImageCompression
        //ExFor:PdfSaveOptions.JpegQuality
        //ExFor:PdfImageCompression
        //ExFor:PdfCompliance
        //ExSummary:Demonstrates how to save images to PDF using JPEG encoding to decrease file size.
        Document doc = new Document(getMyDir() + "SaveOptions.PdfImageCompression.rtf");

        PdfSaveOptions options = new PdfSaveOptions();

        options.setImageCompression(PdfImageCompression.JPEG);
        options.setPreserveFormFields(true);

        doc.save(getMyDir() + "\\Artifacts\\SaveOptions.PdfImageCompression Out.pdf", options);

        PdfSaveOptions optionsA1B = new PdfSaveOptions();
        optionsA1B.setCompliance(PdfCompliance.PDF_A_1_B);
        optionsA1B.setImageCompression(PdfImageCompression.JPEG);
        optionsA1B.setJpegQuality(100); // Use JPEG compression at 50% quality to reduce file size.

        doc.save(getMyDir() + "\\Artifacts\\SaveOptions.PdfImageComppression PDF_A_1_B Out.pdf", optionsA1B);
        //ExEnd

        PdfSaveOptions optionsA1A = new PdfSaveOptions();
        optionsA1A.setCompliance(PdfCompliance.PDF_A_1_A);
        optionsA1A.setExportDocumentStructure(true);
        optionsA1A.setImageCompression(PdfImageCompression.JPEG);

        doc.save(getMyDir() + "\\Artifacts\\SaveOptions.PdfImageComppression PDF_A_1_A Out.pdf", optionsA1A);
    }

    @Test
    public void colorRendering() throws Exception
    {
        //ExStart
        //ExFor:SaveOptions.ColorMode
        //ExSummary:Shows how change image color with save options property
        // Open document with color image
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // Set grayscale mode for document
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        pdfSaveOptions.setColorMode(ColorMode.GRAYSCALE);

        // Assert that color image in document was grey
        doc.save(getMyDir() + "\\Artifacts\\ColorMode.PdfGrayscaleMode.pdf", pdfSaveOptions);
        //ExEnd
    }

    @Test
    public void windowsBarPdfTitle() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.DisplayDocTitle
        //ExSummary:Shows how to display title of the document as title bar.
        Document doc = new Document(getMyDir() + "Rendering.doc");
        doc.getBuiltInDocumentProperties().setTitle("Windows bar pdf title");

        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        pdfSaveOptions.setDisplayDocTitle(true);

        doc.save(getMyDir() + "\\Artifacts\\PdfTitle.pdf", pdfSaveOptions);
        //ExEnd
    }

    @Test
    public void memoryOptimization() throws Exception
    {
        //ExStart
        //ExFor:SaveOptions.MemoryOptimization
        //ExSummary:Shows an option to optimize memory consumption when you work with large documents.
        Document doc = new Document(getMyDir() + "SaveOptions.MemoryOptimization.doc");

        // When set to true it will improve document memory footprint but will add extra time to processing. 
        // This optimization is only applied during save operation.
        SaveOptions saveOptions = SaveOptions.createSaveOptions(SaveFormat.PDF);
        saveOptions.setMemoryOptimization(true);

        doc.save(getMyDir() + "\\Artifacts\\SaveOptions.MemoryOptimization Out.pdf", saveOptions);
        //ExEnd
    }

    @Test
    public void handleBinaryRasterWarnings() throws Exception
    {
        //ExStart
        //ExFor:MetafileRenderingMode.VectorWithFallback
        //ExFor:IWarningCallback
        //ExFor:PdfSaveOptions.MetafileRenderingOptions
        //ExSummary:Shows added fallback to bitmap rendering and changing type of warnings about unsupported metafile records
        Document doc = new Document(getMyDir() + "PdfSaveOptions.HandleRasterWarnings.doc");

        MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions();
        metafileRenderingOptions.setEmulateRasterOperations(false);

        //If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words renders this metafile to a bitmap. 
        metafileRenderingOptions.setRenderingMode(MetafileRenderingMode.VECTOR_WITH_FALLBACK);

        HandleDocumentWarnings callback = new HandleDocumentWarnings();
        doc.setWarningCallback(callback);

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setMetafileRenderingOptions(metafileRenderingOptions);

        doc.save(getMyDir() + "PdfSaveOptions.HandleRasterWarnings Out.pdf", saveOptions);

        Assert.assertEquals(1, callback.mWarnings.getCount());
        Assert.assertTrue(callback.mWarnings.get(0).getDescription().contains("R2_XORPEN"));
    }

    public static class HandleDocumentWarnings implements IWarningCallback
    {
        /// <summary>
        /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
        /// potential issue during document procssing. The callback can be set to listen for warnings generated during document
        /// load and/or document save.
        /// </summary>
        public void warning(WarningInfo info)
        {
            //For now type of warnings about unsupported metafile records changed from DataLoss/UnexpectedContent to MinorFormattingLoss.
            if (info.getWarningType() == WarningType.MINOR_FORMATTING_LOSS)
            {
                System.out.println("Unsupported operation: " + info.getDescription());
                this.mWarnings.warning(info);
            }
        }

        public WarningInfoCollection mWarnings = new WarningInfoCollection();
    }//ExEnd
}
