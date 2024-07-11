package DocsExamples.File_Formats_and_Conversions.Save_Options;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.MetafileRenderingOptions;
import com.aspose.words.MetafileRenderingMode;
import com.aspose.words.WarningInfo;
import com.aspose.ms.System.msConsole;
import com.aspose.words.IWarningCallback;
import com.aspose.words.WarningType;
import com.aspose.words.WarningInfoCollection;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.PdfDigitalSignatureDetails;
import com.aspose.words.CertificateHolder;
import java.util.Date;
import com.aspose.ms.System.DateTime;
import com.aspose.words.PdfFontEmbeddingMode;
import com.aspose.words.HeaderFooterBookmarksExportMode;
import com.aspose.words.PdfCompliance;
import com.aspose.words.PdfCustomPropertiesExport;
import com.aspose.words.PdfImageCompression;
import com.aspose.words.Dml3DEffectsRenderingMode;
import com.aspose.words.FieldHyperlink;


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
    //GistId:f9c5250f94e595ea3590b3be679475ba
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
    //ExEnd:PdfRenderWarnings

    @Test
    public void digitallySignedPdfUsingCertificateHolder() throws Exception
    {
        //ExStart:DigitallySignedPdfUsingCertificateHolder
        //GistId:bdc15a6de6b25d9d4e66f2ce918fc01b
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
        //ExStart:EmbeddedAllFonts
        //GistId:6debb84fc15c7e5b8e35384d9c116215
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // The output PDF will be embedded with all fonts found in the document.
        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setEmbedFullFonts(true); }
        
        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.EmbeddedAllFonts.pdf", saveOptions);
        //ExEnd:EmbeddedAllFonts
    }

    @Test
    public void embeddedSubsetFonts() throws Exception
    {
        //ExStart:EmbeddedSubsetFonts
        //GistId:6debb84fc15c7e5b8e35384d9c116215
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // The output PDF will contain subsets of the fonts in the document.
        // Only the glyphs used in the document are included in the PDF fonts.
        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setEmbedFullFonts(false); }
        
        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.EmbeddedSubsetFonts.pdf", saveOptions);
        //ExEnd:EmbeddedSubsetFonts
    }

    @Test
    public void disableEmbedWindowsFonts() throws Exception
    {
        //ExStart:DisableEmbedWindowsFonts
        //GistId:6debb84fc15c7e5b8e35384d9c116215
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
        //GistId:6debb84fc15c7e5b8e35384d9c116215
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
        //GistId:6debb84fc15c7e5b8e35384d9c116215
        Document doc = new Document(getMyDir() + "Bookmarks in headers and footers.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.getOutlineOptions().setDefaultBookmarksOutlineLevel(1);
        saveOptions.setHeaderFooterBookmarksExportMode(HeaderFooterBookmarksExportMode.FIRST);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.ExportHeaderFooterBookmarks.pdf", saveOptions);
        //ExEnd:ExportHeaderFooterBookmarks
    }

    @Test
    public void emulateRenderingToSizeOnPage() throws Exception
    {
        //ExStart:EmulateRenderingToSizeOnPage
        Document doc = new Document(getMyDir() + "WMF with text.docx");

        MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions();
        {
            metafileRenderingOptions.setEmulateRenderingToSizeOnPage(false);
        }

        // If Aspose.Words cannot correctly render some of the metafile records to vector graphics
        // then Aspose.Words renders this metafile to a bitmap.
        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setMetafileRenderingOptions(metafileRenderingOptions); }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.EmulateRenderingToSizeOnPage.pdf", saveOptions);
        //ExEnd:EmulateRenderingToSizeOnPage
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
        //ExStart:ConversionToPdf17
        //GistId:a53bdaad548845275c1b9556ee21ae65
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setCompliance(PdfCompliance.PDF_17); }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.ConversionToPdf17.pdf", saveOptions);
        //ExEnd:ConversionToPdf17
    }

    @Test
    public void downsamplingImages() throws Exception
    {
        //ExStart:DownsamplingImages
        //GistId:6debb84fc15c7e5b8e35384d9c116215
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // We can set a minimum threshold for downsampling.
        // This value will prevent the second image in the input document from being downsampled.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        {
            saveOptions.setDownsampleOptions({ saveOptions.getDownsampleOptions().setResolution(36); saveOptions.getDownsampleOptions().setResolutionThreshold(128); });
        }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.DownsamplingImages.pdf", saveOptions);
        //ExEnd:DownsamplingImages
    }

    @Test
    public void outlineOptions() throws Exception
    {
        //ExStart:OutlineOptions
        //GistId:6debb84fc15c7e5b8e35384d9c116215
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.getOutlineOptions().setHeadingsOutlineLevels(3);
        saveOptions.getOutlineOptions().setExpandedOutlineLevels(1);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.OutlineOptions.pdf", saveOptions);
        //ExEnd:OutlineOptions
    }

    @Test
    public void customPropertiesExport() throws Exception
    {
        //ExStart:CustomPropertiesExport
        //GistId:6debb84fc15c7e5b8e35384d9c116215
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
        //GistId:6debb84fc15c7e5b8e35384d9c116215
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        // The file size will be increased and the structure will be visible in the "Content" navigation pane
        // of Adobe Acrobat Pro, while editing the .pdf.
        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setExportDocumentStructure(true); }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.ExportDocumentStructure.pdf", saveOptions);
        //ExEnd:ExportDocumentStructure
    }

    @Test
    public void imageCompression() throws Exception
    {
        //ExStart:ImageCompression
        //GistId:6debb84fc15c7e5b8e35384d9c116215
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        {
            saveOptions.setImageCompression(PdfImageCompression.JPEG); saveOptions.setPreserveFormFields(true);
        }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.ImageCompression.pdf", saveOptions);

        PdfSaveOptions saveOptionsA2U = new PdfSaveOptions();
        {
            saveOptionsA2U.setCompliance(PdfCompliance.PDF_A_2_U);
            saveOptionsA2U.setImageCompression(PdfImageCompression.JPEG);
            saveOptionsA2U.setJpegQuality(100); // Use JPEG compression at 50% quality to reduce file size.
        }

        

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.ImageCompression_A2u.pdf", saveOptionsA2U);
        //ExEnd:ImageCompression
    }

    @Test
    public void updateLastPrinted() throws Exception
    {
        //ExStart:UpdateLastPrinted
        //GistId:83e5c469d0e72b5114fb8a05a1d01977
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setUpdateLastPrintedProperty(true); }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.UpdateLastPrinted.pdf", saveOptions);
        //ExEnd:UpdateLastPrinted
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

    @Test
    public void optimizeOutput() throws Exception
    {
        //ExStart:OptimizeOutput
        //GistId:a53bdaad548845275c1b9556ee21ae65
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setOptimizeOutput(true); }

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.OptimizeOutput.pdf", saveOptions);
        //ExEnd:OptimizeOutput
    }

    @Test
    public void updateScreenTip() throws Exception
    {
        //ExStart:UpdateScreenTip
        //GistId:8b0ab362f95040ada1255a0473acefe2
        Document doc = new Document(getMyDir() + "Table of contents.docx");

        var tocHyperLinks = doc.getRange().getFields()
            .Where(f => f.Type == FieldType.FieldHyperlink)
            .<FieldHyperlink>Cast()
            .Where(f => f.SubAddress.StartsWith("#_Toc"));

        for (FieldHyperlink link : (Iterable<FieldHyperlink>) tocHyperLinks)
            link.setScreenTip(link.getDisplayResult());

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        {
            saveOptions.setCompliance(PdfCompliance.PDF_UA_1);
            saveOptions.setDisplayDocTitle(true);
            saveOptions.setExportDocumentStructure(true);
        }
        saveOptions.getOutlineOptions().setHeadingsOutlineLevels(3);
        saveOptions.getOutlineOptions().setCreateMissingOutlineLevels(true);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.UpdateScreenTip.pdf", saveOptions);
        //ExEnd:UpdateScreenTip
    }
}
