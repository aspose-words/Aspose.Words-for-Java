package DocsExamples.File_formats_and_conversions.Save_options;

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.MetafileRenderingOptions;
import com.aspose.words.MetafileRenderingMode;
import com.aspose.words.WarningInfo;
import com.aspose.words.IWarningCallback;
import com.aspose.words.WarningType;
import com.aspose.words.WarningInfoCollection;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.PdfDigitalSignatureDetails;
import com.aspose.words.CertificateHolder;
import java.util.Date;
import com.aspose.words.PdfFontEmbeddingMode;
import com.aspose.words.HeaderFooterBookmarksExportMode;
import com.aspose.words.PdfCompliance;
import com.aspose.words.PdfCustomPropertiesExport;
import com.aspose.words.PdfImageCompression;
import com.aspose.words.PdfImageColorSpaceExportMode;
import com.aspose.words.Dml3DEffectsRenderingMode;

@Test
public class WorkingWithPdfSaveOptions extends DocsExamplesBase
{
    @Test
    public void displayDocTitleInWindowTitlebar() throws Exception
    {
        //ExStart:DisplayDocTitleInWindowTitlebar
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setDisplayDocTitle(true); }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.DisplayDocTitleInWindowTitlebar.pdf", saveOptions);
        //ExEnd:DisplayDocTitleInWindowTitlebar
    }

    @Test
    //ExStart:PdfRenderWarnings
    public void pdfRenderWarnings() throws Exception
    {
        Document doc = new Document(getMyDir() + "WMF with image.docx");

        MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions();
        {
            metafileRenderingOptions.setEmulateRasterOperations(false); metafileRenderingOptions.setRenderingMode(MetafileRenderingMode.VECTOR_WITH_FALLBACK);
        }

        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setMetafileRenderingOptions(metafileRenderingOptions); }

        // If Aspose.Words cannot correctly render some of the metafile records
        // to vector graphics then Aspose.Words renders this metafile to a bitmap.
        HandleDocumentWarnings callback = new HandleDocumentWarnings();
        doc.setWarningCallback(callback);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.PdfRenderWarnings.pdf", saveOptions);

        // While the file saves successfully, rendering warnings that occurred during saving are collected here.
        for (WarningInfo warningInfo : callback.mWarnings)
        {
            System.out.println(warningInfo.getDescription());
        }
    }

    //ExStart:RenderMetafileToBitmap
    public static class HandleDocumentWarnings implements IWarningCallback
    {
        /// <summary>
        /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
        /// potential issue during document processing. The callback can be set to listen for warnings generated during
        /// document load and/or document save.
        /// </summary>
        public void warning(WarningInfo info)
        {
            // For now type of warnings about unsupported metafile records changed
            // from DataLoss/UnexpectedContent to MinorFormattingLoss.
            if (info.getWarningType() == WarningType.MINOR_FORMATTING_LOSS)
            {
                System.out.println("Unsupported operation: " + info.getDescription());
                mWarnings.warning(info);
            }
        }

        public WarningInfoCollection mWarnings = new WarningInfoCollection();
    }
    //ExEnd:RenderMetafileToBitmap
    //ExEnd:PdfRenderWarnings

    @Test
    public void digitallySignedPdfUsingCertificateHolder() throws Exception
    {
        //ExStart:DigitallySignedPdfUsingCertificateHolder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.writeln("Test Signed PDF.");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        {
            saveOptions.setDigitalSignatureDetails(new PdfDigitalSignatureDetails(
                CertificateHolder.create(getMyDir() + "morzal.pfx", "aw"), "reason", "location",
                new Date()));
        }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.DigitallySignedPdfUsingCertificateHolder.pdf", saveOptions);
        //ExEnd:DigitallySignedPdfUsingCertificateHolder
    }

    @Test
    public void embeddedAllFonts() throws Exception
    {
        //ExStart:EmbeddAllFonts
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // The output PDF will be embedded with all fonts found in the document.
        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setEmbedFullFonts(true); }
        
        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.EmbeddedFontsInPdf.pdf", saveOptions);
        //ExEnd:EmbeddAllFonts
    }

    @Test
    public void embeddedSubsetFonts() throws Exception
    {
        //ExStart:EmbeddSubsetFonts
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // The output PDF will contain subsets of the fonts in the document.
        // Only the glyphs used in the document are included in the PDF fonts.
        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setEmbedFullFonts(false); }
        
        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.EmbeddSubsetFonts.pdf", saveOptions);
        //ExEnd:EmbeddSubsetFonts
    }

    @Test
    public void disableEmbedWindowsFonts() throws Exception
    {
        //ExStart:DisableEmbedWindowsFonts
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // The output PDF will be saved without embedding standard windows fonts.
        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setFontEmbeddingMode(PdfFontEmbeddingMode.EMBED_NONE); }
        
        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.DisableEmbedWindowsFonts.pdf", saveOptions);
        //ExEnd:DisableEmbedWindowsFonts
    }

    @Test
    public void skipEmbeddedArialAndTimesRomanFonts() throws Exception
    {
        //ExStart:SkipEmbeddedArialAndTimesRomanFonts
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setFontEmbeddingMode(PdfFontEmbeddingMode.EMBED_ALL); }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.SkipEmbeddedArialAndTimesRomanFonts.pdf", saveOptions);
        //ExEnd:SkipEmbeddedArialAndTimesRomanFonts
    }

    @Test
    public void avoidEmbeddingCoreFonts() throws Exception
    {
        //ExStart:AvoidEmbeddingCoreFonts
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setUseCoreFonts(true); }
        
        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.AvoidEmbeddingCoreFonts.pdf", saveOptions);
        //ExEnd:AvoidEmbeddingCoreFonts
    }
    
    @Test
    public void escapeUri() throws Exception
    {
        //ExStart:EscapeUri
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.insertHyperlink("Testlink", 
            "https://www.google.com/search?q=%2Fthe%20test", false);
        builder.writeln();
        builder.insertHyperlink("https://www.google.com/search?q=%2Fthe%20test", 
            "https://www.google.com/search?q=%2Fthe%20test", false);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.EscapeUri.pdf");
        //ExEnd:EscapeUri
    }

    @Test
    public void exportHeaderFooterBookmarks() throws Exception
    {
        //ExStart:ExportHeaderFooterBookmarks
        Document doc = new Document(getMyDir() + "Bookmarks in headers and footers.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.getOutlineOptions().setDefaultBookmarksOutlineLevel(1);
        saveOptions.setHeaderFooterBookmarksExportMode(HeaderFooterBookmarksExportMode.FIRST);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.ExportHeaderFooterBookmarks.pdf", saveOptions);
        //ExEnd:ExportHeaderFooterBookmarks
    }

    @Test
    public void scaleWmfFontsToMetafileSize() throws Exception
    {
        //ExStart:ScaleWmfFontsToMetafileSize
        Document doc = new Document(getMyDir() + "WMF with text.docx");

        MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions();
        {
            metafileRenderingOptions.setScaleWmfFontsToMetafileSize(false);
        }

        // If Aspose.Words cannot correctly render some of the metafile records to vector graphics
        // then Aspose.Words renders this metafile to a bitmap.
        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setMetafileRenderingOptions(metafileRenderingOptions); }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.ScaleWmfFontsToMetafileSize.pdf", saveOptions);
        //ExEnd:ScaleWmfFontsToMetafileSize
    }

    @Test
    public void additionalTextPositioning() throws Exception
    {
        //ExStart:AdditionalTextPositioning
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setAdditionalTextPositioning(true); }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.AdditionalTextPositioning.pdf", saveOptions);
        //ExEnd:AdditionalTextPositioning
    }

    @Test
    public void conversionToPdf17() throws Exception
    {
        //ExStart:ConversionToPDF17
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setCompliance(PdfCompliance.PDF_17); }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.ConversionToPdf17.pdf", saveOptions);
        //ExEnd:ConversionToPDF17
    }

    @Test
    public void downsamplingImages() throws Exception
    {
        //ExStart:DownsamplingImages
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // We can set a minimum threshold for downsampling.
        // This value will prevent the second image in the input document from being downsampled.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        {
            saveOptions.getDownsampleOptions().setResolution(36);
            saveOptions.getDownsampleOptions().setResolutionThreshold(128);
        }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.DownsamplingImages.pdf", saveOptions);
        //ExEnd:DownsamplingImages
    }

    @Test
    public void setOutlineOptions() throws Exception
    {
        //ExStart:SetOutlineOptions
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.getOutlineOptions().setHeadingsOutlineLevels(3);
        saveOptions.getOutlineOptions().setExpandedOutlineLevels(1);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.SetOutlineOptions.pdf", saveOptions);
        //ExEnd:SetOutlineOptions
    }

    @Test
    public void customPropertiesExport() throws Exception
    {
        //ExStart:CustomPropertiesExport
        Document doc = new Document();
        doc.getCustomDocumentProperties().add("Company", "Aspose");

        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setCustomPropertiesExport(PdfCustomPropertiesExport.STANDARD); }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.CustomPropertiesExport.pdf", saveOptions);
        //ExEnd:CustomPropertiesExport
    }

    @Test
    public void exportDocumentStructure() throws Exception
    {
        //ExStart:ExportDocumentStructure
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        // The file size will be increased and the structure will be visible in the "Content" navigation pane
        // of Adobe Acrobat Pro, while editing the .pdf.
        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setExportDocumentStructure(true); }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.ExportDocumentStructure.pdf", saveOptions);
        //ExEnd:ExportDocumentStructure
    }

    @Test
    public void pdfImageComppression() throws Exception
    {
        //ExStart:PdfImageComppression
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        {
            saveOptions.setImageCompression(PdfImageCompression.JPEG); saveOptions.setPreserveFormFields(true);
        }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.PdfImageCompression.pdf", saveOptions);

        PdfSaveOptions saveOptionsA1B = new PdfSaveOptions();
        {
            saveOptionsA1B.setCompliance(PdfCompliance.PDF_A_1_B);
            saveOptionsA1B.setImageCompression(PdfImageCompression.JPEG);
            saveOptionsA1B.setJpegQuality(100); // Use JPEG compression at 50% quality to reduce file size.
            saveOptionsA1B.setImageColorSpaceExportMode(PdfImageColorSpaceExportMode.SIMPLE_CMYK);
        }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.PdfImageCompression.Pdf_A1b.pdf", saveOptionsA1B);
        //ExEnd:PdfImageComppression
    }

    @Test
    public void updateLastPrintedProperty() throws Exception
    {
        //ExStart:UpdateIfLastPrinted
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setUpdateLastPrintedProperty(true); }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.UpdateIfLastPrinted.pdf", saveOptions);
        //ExEnd:UpdateIfLastPrinted
    }

    @Test
    public void dml3DEffectsRendering() throws Exception
    {
        //ExStart:Dml3DEffectsRendering
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setDml3DEffectsRenderingMode(Dml3DEffectsRenderingMode.ADVANCED); }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.Dml3DEffectsRendering.pdf", saveOptions);
        //ExEnd:Dml3DEffectsRendering
    }

    @Test
    public void interpolateImages() throws Exception
    {
        //ExStart:SetImageInterpolation
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setInterpolateImages(true); }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.InterpolateImages.pdf", saveOptions);
        //ExEnd:SetImageInterpolation
    }
}
