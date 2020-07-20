package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.pdf.TableAbsorber;
import com.aspose.pdf.TextFragmentAbsorber;
import com.aspose.pdf.facades.Bookmarks;
import com.aspose.pdf.facades.PdfBookmarkEditor;
import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Date;

public class ExPdfSaveOptions extends ApiExampleBase {
    @Test
    public void createMissingOutlineLevels() throws Exception {
        //ExStart
        //ExFor:OutlineOptions.CreateMissingOutlineLevels
        //ExFor:ParagraphFormat.IsHeading
        //ExFor:PdfSaveOptions.OutlineOptions
        //ExFor:PdfSaveOptions.SaveFormat
        //ExSummary:Shows how to create PDF document outline entries for headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Creating TOC entries
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        Assert.assertTrue(builder.getParagraphFormat().isHeading());

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

        doc.save(getArtifactsDir() + "PdfSaveOptions.CreateMissingOutlineLevels.pdf", pdfSaveOptions);
        //ExEnd

        // Bind PDF with Aspose.PDF
        PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
        bookmarkEditor.bindPdf(getArtifactsDir() + "PdfSaveOptions.CreateMissingOutlineLevels.pdf");

        // Get all bookmarks from the document
        Bookmarks bookmarks = bookmarkEditor.extractBookmarks();
        Assert.assertEquals(11, bookmarks.size());

        bookmarkEditor.close();
    }

    @Test
    public void tableHeadingOutlines() throws Exception {
        //ExStart
        //ExFor:OutlineOptions.CreateOutlinesForHeadingsInTables
        //ExSummary:Shows how to create PDF document outline entries for headings inside tables.
        // Create a blank document and insert a table with a heading-style text inside it
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startTable();
        builder.insertCell();
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.write("Heading 1");
        builder.endRow();
        builder.insertCell();
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.NORMAL);
        builder.write("Cell 1");
        builder.endTable();

        // Create a PdfSaveOptions object that, when saving to .pdf with it, creates entries in the document outline for all headings levels 1-9,
        // and make sure headings inside tables are registered by the outline also
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        pdfSaveOptions.getOutlineOptions().setHeadingsOutlineLevels(9);
        pdfSaveOptions.getOutlineOptions().setCreateOutlinesForHeadingsInTables(true);

        doc.save(getArtifactsDir() + "PdfSaveOptions.TableHeadingOutlines.pdf", pdfSaveOptions);
        //ExEnd

        com.aspose.pdf.Document pdfDoc = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.TableHeadingOutlines.pdf");

        Assert.assertEquals(1, pdfDoc.getOutlines().size());
        Assert.assertEquals("Heading 1", pdfDoc.getOutlines().get_Item(1).getTitle());

        TableAbsorber tableAbsorber = new TableAbsorber();
        tableAbsorber.visit(pdfDoc.getPages().get_Item(1));

        Assert.assertEquals("Heading 1", tableAbsorber.getTableList().get_Item(0).getRowList().get_Item(0).getCellList().get_Item(0).getTextFragments().get_Item(1).getText());
        Assert.assertEquals("Cell 1", tableAbsorber.getTableList().get_Item(0).getRowList().get_Item(1).getCellList().get_Item(0).getTextFragments().get_Item(1).getText());

        pdfDoc.close();
    }

    @Test(groups = "SkipMono", dataProvider = "updateFieldsDataProvider")
    public void updateFields(boolean doUpdateFields) throws Exception {
        //ExStart
        //ExFor:PdfSaveOptions.Clone
        //ExFor:SaveOptions.UpdateFields
        //ExSummary:Shows how to update fields before saving into a PDF document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two pages of text, including two fields that will need to be updated to display an accurate value
        builder.write("Page ");
        builder.insertField("PAGE", "");
        builder.write(" of ");
        builder.insertField("NUMPAGES", "");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Hello World!");

        PdfSaveOptions options = new PdfSaveOptions();
        options.setUpdateFields(doUpdateFields);

        // PdfSaveOptions objects can be cloned
        Assert.assertNotSame(options, options.deepClone());

        doc.save(getArtifactsDir() + "PdfSaveOptions.UpdateFields.pdf", options);
        //ExEnd

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.UpdateFields.pdf");

        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
        pdfDocument.getPages().accept(textFragmentAbsorber);

        if (doUpdateFields)
            Assert.assertEquals("Page 1 of 2", textFragmentAbsorber.getTextFragments().get_Item(1).getText());
        else
            Assert.assertEquals("Page  of ", textFragmentAbsorber.getTextFragments().get_Item(1).getText());

        pdfDocument.close();
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "updateFieldsDataProvider")
    public static Object[][] updateFieldsDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void imageCompression() throws Exception {
        //ExStart
        //ExFor:PdfSaveOptions.Compliance
        //ExFor:PdfSaveOptions.ImageCompression
        //ExFor:PdfSaveOptions.ImageColorSpaceExportMode
        //ExFor:PdfSaveOptions.JpegQuality
        //ExFor:PdfImageCompression
        //ExFor:PdfCompliance
        //ExFor:PdfImageColorSpaceExportMode
        //ExSummary:Shows how to save images to PDF using JPEG encoding to decrease file size.
        Document doc = new Document(getMyDir() + "Images.docx");

        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        pdfSaveOptions.setImageCompression(PdfImageCompression.JPEG);
        pdfSaveOptions.getDownsampleOptions().setDownsampleImages(false);

        doc.save(getArtifactsDir() + "PdfSaveOptions.ImageCompression.pdf", pdfSaveOptions);

        PdfSaveOptions pdfSaveOptionsA1B = new PdfSaveOptions();
        pdfSaveOptionsA1B.setCompliance(PdfCompliance.PDF_A_1_B);
        pdfSaveOptionsA1B.setImageCompression(PdfImageCompression.JPEG);
        pdfSaveOptionsA1B.getDownsampleOptions().setDownsampleImages(false);
        // Use JPEG compression at 50% quality to reduce file size
        pdfSaveOptionsA1B.setJpegQuality(100);
        pdfSaveOptionsA1B.setImageColorSpaceExportMode(PdfImageColorSpaceExportMode.SIMPLE_CMYK);

        doc.save(getArtifactsDir() + "PdfSaveOptions.ImageCompression.PDF_A_1_B.pdf", pdfSaveOptionsA1B);

        PdfSaveOptions pdfSaveOptionsA1A = new PdfSaveOptions();
        pdfSaveOptionsA1A.setCompliance(PdfCompliance.PDF_A_1_A);
        pdfSaveOptionsA1A.setExportDocumentStructure(true);
        pdfSaveOptionsA1A.setImageCompression(PdfImageCompression.JPEG);
        pdfSaveOptionsA1A.getDownsampleOptions().setDownsampleImages(false);

        doc.save(getArtifactsDir() + "PdfSaveOptions.ImageCompression.PDF_A_1_A.pdf", pdfSaveOptionsA1A);
        //ExEnd
    }

    @Test
    public void colorRendering() throws Exception {
        //ExStart
        //ExFor:PdfSaveOptions
        //ExFor:ColorMode
        //ExFor:FixedPageSaveOptions.ColorMode
        //ExSummary:Shows how change image color with save options property.
        Document doc = new Document(getMyDir() + "Images.docx");

        // Configure PdfSaveOptions to save every image in the input document in greyscale during conversion
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        {
            pdfSaveOptions.setColorMode(ColorMode.GRAYSCALE);
        }

        doc.save(getArtifactsDir() + "PdfSaveOptions.ColorRendering.pdf", pdfSaveOptions);
        //ExEnd
    }

    @Test
    public void windowsBarPdfTitle() throws Exception {
        //ExStart
        //ExFor:PdfSaveOptions.DisplayDocTitle
        //ExSummary:Shows how to display title of the document as title bar.
        Document doc = new Document(getMyDir() + "Rendering.docx");
        doc.getBuiltInDocumentProperties().setTitle("Windows bar pdf title");

        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        pdfSaveOptions.setDisplayDocTitle(true);

        doc.save(getArtifactsDir() + "PdfSaveOptions.WindowsBarPdfTitle.pdf", pdfSaveOptions);
        //ExEnd
    }

    @Test
    public void memoryOptimization() throws Exception {
        //ExStart
        //ExFor:SaveOptions.CreateSaveOptions(SaveFormat)
        //ExFor:SaveOptions.MemoryOptimization
        //ExSummary:Shows an option to optimize memory consumption when you work with large documents.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // When set to true it will improve document memory footprint but will add extra time to processing
        SaveOptions saveOptions = SaveOptions.createSaveOptions(SaveFormat.PDF);
        saveOptions.setMemoryOptimization(true);

        doc.save(getArtifactsDir() + "PdfSaveOptions.MemoryOptimization.pdf", saveOptions);
        //ExEnd
    }

    @Test(dataProvider = "escapeUriDataProvider")
    public void escapeUri(final String uri, final String result, final boolean isEscaped) throws Exception {
        //ExStart
        //ExFor:PdfSaveOptions.EscapeUri
        //ExFor:PdfSaveOptions.OpenHyperlinksInNewWindow
        //ExSummary:Shows how to escape hyperlinks in the document.
        DocumentBuilder builder = new DocumentBuilder();
        builder.insertHyperlink("Testlink", uri, false);

        // Set this property to false if you are sure that hyperlinks in document's model are already escaped
        PdfSaveOptions options = new PdfSaveOptions();
        options.setEscapeUri(isEscaped);
        options.setOpenHyperlinksInNewWindow(true);

        builder.getDocument().save(getArtifactsDir() + "PdfSaveOptions.EscapedUri.pdf", options);
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "escapeUriDataProvider")
    public static Object[][] escapeUriDataProvider() {
        return new Object[][]
                {
                        {"https://www.google.com/search?q= aspose", "app.launchURL(\"https://www.google.com/search?q=%20aspose\", true);", true},
                        {"https://www.google.com/search?q=%20aspose", "app.launchURL(\"https://www.google.com/search?q=%20aspose\", true);", true},
                        {"https://www.google.com/search?q= aspose", "app.launchURL(\"https://www.google.com/search?q= aspose\", true);", false},
                        {"https://www.google.com/search?q=%20aspose", "app.launchURL(\"https://www.google.com/search?q=%20aspose\", true);", false}
                };
    }

    //ExStart
    //ExFor:MetafileRenderingMode
    //ExFor:MetafileRenderingOptions
    //ExFor:MetafileRenderingOptions.EmulateRasterOperations
    //ExFor:MetafileRenderingOptions.RenderingMode
    //ExFor:IWarningCallback
    //ExFor:FixedPageSaveOptions.MetafileRenderingOptions
    //ExSummary:Shows added fallback to bitmap rendering and changing type of warnings about unsupported metafile records.
    @Test(groups = "SkipMono") //ExSkip
    public void handleBinaryRasterWarnings() throws Exception {
        Document doc = new Document(getMyDir() + "WMF with image.docx");

        MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions();
        metafileRenderingOptions.setEmulateRasterOperations(false);
        metafileRenderingOptions.setRenderingMode(MetafileRenderingMode.VECTOR_WITH_FALLBACK);

        // If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words renders this metafile to a bitmap
        HandleDocumentWarnings callback = new HandleDocumentWarnings();
        doc.setWarningCallback(callback);

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setMetafileRenderingOptions(metafileRenderingOptions);

        doc.save(getArtifactsDir() + "PdfSaveOptions.HandleBinaryRasterWarnings.pdf", saveOptions);

        Assert.assertEquals(callback.mWarnings.getCount(), 1);
        Assert.assertTrue(callback.mWarnings.get(0).getDescription().contains("R2_XORPEN"));
    }

    public static class HandleDocumentWarnings implements IWarningCallback {
        /**
         * Our callback only needs to implement the "Warning" method. This method is called whenever there is a
         * potential issue during document processing. The callback can be set to listen for warnings generated during document
         * load and/or document save.
         */
        public void warning(final WarningInfo info) {
            //For now type of warnings about unsupported metafile records changed from DataLoss/UnexpectedContent to MinorFormattingLoss
            if (info.getWarningType() == WarningType.MINOR_FORMATTING_LOSS) {
                System.out.println("Unsupported operation: " + info.getDescription());
                this.mWarnings.warning(info);
            }
        }

        public WarningInfoCollection mWarnings = new WarningInfoCollection();
    }
    //ExEnd

    @Test(dataProvider = "headerFooterBookmarksExportModeDataProvider")
    public void headerFooterBookmarksExportMode(final int headerFooterBookmarksExportMode) throws Exception {
        //ExStart
        //ExFor:HeaderFooterBookmarksExportMode
        //ExFor:OutlineOptions
        //ExFor:OutlineOptions.DefaultBookmarksOutlineLevel
        //ExFor:PdfSaveOptions.HeaderFooterBookmarksExportMode
        //ExFor:PdfSaveOptions.PageMode
        //ExFor:PdfPageMode
        //ExSummary:Shows how bookmarks in headers/footers are exported to pdf.
        Document doc = new Document(getMyDir() + "Bookmarks in headers and footers.docx");

        // You can specify how bookmarks in headers/footers are exported
        // There is a several options for this:
        // "None" - Bookmarks in headers/footers are not exported
        // "First" - Only bookmark in first header/footer of the section is exported
        // "All" - Bookmarks in all headers/footers are exported
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        saveOptions.setHeaderFooterBookmarksExportMode(headerFooterBookmarksExportMode);
        saveOptions.getOutlineOptions().setDefaultBookmarksOutlineLevel(1);
        saveOptions.setPageMode(PdfPageMode.USE_OUTLINES);

        doc.save(getArtifactsDir() + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf", saveOptions);
        //ExEnd
    }

    @DataProvider(name = "headerFooterBookmarksExportModeDataProvider")
    public static Object[][] headerFooterBookmarksExportModeDataProvider() {
        return new Object[][]
                {
                        {com.aspose.words.HeaderFooterBookmarksExportMode.NONE},
                        {com.aspose.words.HeaderFooterBookmarksExportMode.FIRST},
                        {com.aspose.words.HeaderFooterBookmarksExportMode.ALL},
                };
    }

    @Test
    public void unsupportedImageFormatWarning() throws Exception {
        Document doc = new Document(getMyDir() + "Corrupted image.docx");

        SaveWarningCallback saveWarningCallback = new SaveWarningCallback();
        doc.setWarningCallback(saveWarningCallback);

        doc.save(getArtifactsDir() + "PdfSaveOption.UnsupportedImageFormatWarning.pdf", SaveFormat.PDF);

        Assert.assertEquals(saveWarningCallback.mSaveWarnings.get(0).getDescription(),
                "Image can not be processed. Possibly unsupported image format.");
    }

    public static class SaveWarningCallback implements IWarningCallback {
        public void warning(final WarningInfo info) {
            if (info.getWarningType() == WarningType.MINOR_FORMATTING_LOSS) {
                System.out.println(MessageFormat.format("{0}: {1}.", info.getWarningType(), info.getDescription()));
                mSaveWarnings.warning(info);
            }
        }

        WarningInfoCollection mSaveWarnings = new WarningInfoCollection();
    }

    @Test(dataProvider = "fontsScaledToMetafileSizeDataProvider")
    public void fontsScaledToMetafileSize(boolean doScaleWmfFonts) throws Exception {
        //ExStart
        //ExFor:MetafileRenderingOptions.ScaleWmfFontsToMetafileSize
        //ExSummary:Shows how to WMF fonts scaling according to metafile size on the page.
        Document doc = new Document(getMyDir() + "WMF with text.docx");

        // There is a several options for this:
        // 'True' - Aspose.Words emulates font scaling according to metafile size on the page
        // 'False' - Aspose.Words displays the fonts as metafile is rendered to its default size
        // Use 'False' option is used only when metafile is rendered as vector graphics
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.getMetafileRenderingOptions().setScaleWmfFontsToMetafileSize(doScaleWmfFonts);

        doc.save(getArtifactsDir() + "PdfSaveOptions.FontsScaledToMetafileSize.pdf", saveOptions);
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "fontsScaledToMetafileSizeDataProvider")
    public static Object[][] fontsScaledToMetafileSizeDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "additionalTextPositioningDataProvider")
    public void additionalTextPositioning(boolean applyAdditionalTextPositioning) throws Exception {
        //ExStart
        //ExFor:PdfSaveOptions.AdditionalTextPositioning
        //ExSummary:Show how to write additional text positioning operators.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // This may help to overcome issues with inaccurate text positioning with some printers, even if the PDF looks fine,
        // but the file size will increase due to higher text positioning precision used
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        {
            saveOptions.setAdditionalTextPositioning(applyAdditionalTextPositioning);
            saveOptions.setTextCompression(PdfTextCompression.NONE);
        }

        doc.save(getArtifactsDir() + "PdfSaveOptions.AdditionalTextPositioning.pdf", saveOptions);
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "additionalTextPositioningDataProvider")
    public static Object[][] additionalTextPositioningDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "saveAsPdfBookFoldDataProvider")
    public void saveAsPdfBookFold(boolean doRenderTextAsBookfold) throws Exception {
        //ExStart
        //ExFor:PdfSaveOptions.UseBookFoldPrintingSettings
        //ExSummary:Shows how to save a document to the PDF format in the form of a book fold.
        // Open a document with multiple paragraphs
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        // Configure both page setup and PdfSaveOptions to create a book fold
        for (Section s : (Iterable<Section>) doc.getSections()) {
            s.getPageSetup().setMultiplePages(MultiplePagesType.BOOK_FOLD_PRINTING);
        }

        PdfSaveOptions options = new PdfSaveOptions();
        options.setUseBookFoldPrintingSettings(doRenderTextAsBookfold);

        // In order to make a booklet, we will need to print this document, stack the pages
        // in the order they come out of the printer and then fold down the middle
        doc.save(getArtifactsDir() + "PdfSaveOptions.SaveAsPdfBookFold.pdf", options);
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "saveAsPdfBookFoldDataProvider")
    public static Object[][] saveAsPdfBookFoldDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void zoomBehaviour() throws Exception {
        //ExStart
        //ExFor:PdfSaveOptions.ZoomBehavior
        //ExFor:PdfSaveOptions.ZoomFactor
        //ExFor:PdfZoomBehavior
        //ExSummary:Shows how to set the default zooming of an output PDF to 1/4 of default size.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions options = new PdfSaveOptions();
        {
            options.setZoomBehavior(PdfZoomBehavior.ZOOM_FACTOR);
            options.setZoomFactor(25);
        }

        // Upon opening the .pdf with a viewer such as Adobe Acrobat Pro, the zoom level will be at 25% by default,
        // with thumbnails for each page to the left
        doc.save(getArtifactsDir() + "PdfSaveOptions.ZoomBehaviour.pdf", options);
        //ExEnd
    }

    @Test(dataProvider = "pageModeDataProvider")
    public void pageMode(/*PdfPageMode*/int pageMode) throws Exception {
        //ExStart
        //ExFor:PdfSaveOptions.PageMode
        //ExFor:PdfPageMode
        //ExSummary:Shows how to set instructions for some PDF readers to follow when opening an output document.
        Document doc = new Document(getMyDir() + "Document.docx");

        PdfSaveOptions options = new PdfSaveOptions();
        options.setPageMode(pageMode);

        doc.save(getArtifactsDir() + "PdfSaveOptions.PageMode.pdf", options);
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "pageModeDataProvider")
    public static Object[][] pageModeDataProvider() throws Exception {
        return new Object[][]
                {
                        {PdfPageMode.FULL_SCREEN},
                        {PdfPageMode.USE_THUMBS},
                        {PdfPageMode.USE_OC},
                        {PdfPageMode.USE_OUTLINES},
                        {PdfPageMode.USE_NONE},
                };
    }

    @Test(dataProvider = "noteHyperlinksDataProvider")
    public void noteHyperlinks(boolean doCreateHyperlinks) throws Exception {
        //ExStart
        //ExFor:PdfSaveOptions.CreateNoteHyperlinks
        //ExSummary:Shows how to make footnotes and endnotes work like hyperlinks.
        // Open a document with footnotes/endnotes
        Document doc = new Document(getMyDir() + "Footnotes and endnotes.docx");

        // Creating a PdfSaveOptions instance with this flag set will convert footnote/endnote number symbols in the text
        // into hyperlinks pointing to the footnotes, and the actual footnotes/endnotes at the end of pages into links to their
        // referenced body text
        PdfSaveOptions options = new PdfSaveOptions();
        options.setCreateNoteHyperlinks(doCreateHyperlinks);

        doc.save(getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf", options);
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "noteHyperlinksDataProvider")
    public static Object[][] noteHyperlinksDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "customPropertiesExportDataProvider")
    public void customPropertiesExport(/*PdfCustomPropertiesExport*/int pdfCustomPropertiesExportMode) throws Exception {
        //ExStart
        //ExFor:PdfCustomPropertiesExport
        //ExFor:PdfSaveOptions.CustomPropertiesExport
        //ExSummary:Shows how to export custom properties while saving to .pdf.
        Document doc = new Document();

        // Add a custom document property that doesn't use the name of some built in properties
        doc.getCustomDocumentProperties().add("Company", "My value");

        // Configure the PdfSaveOptions like this will display the properties
        // in the "Document Properties" menu of Adobe Acrobat Pro
        PdfSaveOptions options = new PdfSaveOptions();
        options.setCustomPropertiesExport(pdfCustomPropertiesExportMode);

        doc.save(getArtifactsDir() + "PdfSaveOptions.CustomPropertiesExport.pdf", options);
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "customPropertiesExportDataProvider")
    public static Object[][] customPropertiesExportDataProvider() throws Exception {
        return new Object[][]
                {
                        {PdfCustomPropertiesExport.NONE},
                        {PdfCustomPropertiesExport.STANDARD},
                        {PdfCustomPropertiesExport.METADATA},
                };
    }

    @Test(dataProvider = "drawingMLEffectsDataProvider")
    public void drawingMLEffects(/*DmlEffectsRenderingMode*/int effectsRenderingMode) throws Exception {
        //ExStart
        //ExFor:DmlRenderingMode
        //ExFor:DmlEffectsRenderingMode
        //ExFor:PdfSaveOptions.DmlEffectsRenderingMode
        //ExFor:SaveOptions.DmlEffectsRenderingMode
        //ExFor:SaveOptions.DmlRenderingMode
        //ExSummary:Shows how to configure DrawingML rendering quality with PdfSaveOptions.
        Document doc = new Document(getMyDir() + "DrawingML shape effects.docx");

        PdfSaveOptions options = new PdfSaveOptions();
        options.setDmlEffectsRenderingMode(effectsRenderingMode);

        Assert.assertEquals(DmlRenderingMode.DRAWING_ML, options.getDmlRenderingMode());

        doc.save(getArtifactsDir() + "PdfSaveOptions.DrawingMLEffects.pdf", options);
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "drawingMLEffectsDataProvider")
    public static Object[][] drawingMLEffectsDataProvider() throws Exception {
        return new Object[][]
                {
                        {DmlEffectsRenderingMode.NONE},
                        {DmlEffectsRenderingMode.SIMPLIFIED},
                        {DmlEffectsRenderingMode.FINE},
                };
    }

    @Test(dataProvider = "drawingMLFallbackDataProvider")
    public void drawingMLFallback(/*DmlRenderingMode*/int dmlRenderingMode) throws Exception {
        //ExStart
        //ExFor:DmlRenderingMode
        //ExFor:SaveOptions.DmlRenderingMode
        //ExSummary:Shows how to render fallback shapes when saving to Pdf.
        Document doc = new Document(getMyDir() + "DrawingML shape fallbacks.docx");

        PdfSaveOptions options = new PdfSaveOptions();
        options.setDmlRenderingMode(dmlRenderingMode);

        doc.save(getArtifactsDir() + "PdfSaveOptions.DrawingMLFallback.pdf", options);
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "drawingMLFallbackDataProvider")
    public static Object[][] drawingMLFallbackDataProvider() throws Exception {
        return new Object[][]
                {
                        {DmlRenderingMode.FALLBACK},
                        {DmlRenderingMode.DRAWING_ML},
                };
    }

    @Test(dataProvider = "exportDocumentStructureDataProvider")
    public void exportDocumentStructure(boolean doExportStructure) throws Exception {
        //ExStart
        //ExFor:PdfSaveOptions.ExportDocumentStructure
        //ExSummary:Shows how to convert a .docx to .pdf while preserving the document structure.
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        // Create a PdfSaveOptions object and configure it to preserve the logical structure that's in the input document
        // The file size will be increased and the structure will be visible in the "Content" navigation pane
        // of Adobe Acrobat Pro
        PdfSaveOptions options = new PdfSaveOptions();
        options.setExportDocumentStructure(doExportStructure);

        doc.save(getArtifactsDir() + "PdfSaveOptions.ExportDocumentStructure.pdf", options);
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "exportDocumentStructureDataProvider")
    public static Object[][] exportDocumentStructureDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "preblendImagesDataProvider")
    public void preblendImages(boolean doPreblendImages) throws Exception {
        //ExStart
        //ExFor:PdfSaveOptions.PreblendImages
        //ExSummary:Shows how to preblend images with transparent backgrounds.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertImage(getImageDir() + "Transparent background logo.png");

        // Setting this flag in a SaveOptions object may change the quality and size of the output .pdf
        // because of the way some images are rendered
        PdfSaveOptions options = new PdfSaveOptions();
        options.setPreblendImages(doPreblendImages);

        doc.save(getArtifactsDir() + "PdfSaveOptions.PreblendImages.pdf", options);
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "preblendImagesDataProvider")
    public static Object[][] preblendImagesDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void interpolateImages() throws Exception {
        //ExStart
        //ExFor:PdfSaveOptions.InterpolateImages
        //ExSummary:Shows how to improve the quality of an image in the rendered documents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        BufferedImage img = ImageIO.read(new File(getImageDir() + "Transparent background logo.png"));
        builder.insertImage(img);

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setInterpolateImages(true);

        doc.save(getArtifactsDir() + "PdfSaveOptions.InterpolateImages.pdf", saveOptions);
        //ExEnd
    }


    @Test
    public void pdfDigitalSignature() throws Exception {
        //ExStart
        //ExFor:PdfDigitalSignatureDetails
        //ExFor:PdfDigitalSignatureDetails.#ctor
        //ExFor:PdfDigitalSignatureDetails.#ctor(CertificateHolder, String, String, DateTime)
        //ExFor:PdfDigitalSignatureDetails.HashAlgorithm
        //ExFor:PdfDigitalSignatureDetails.Location
        //ExFor:PdfDigitalSignatureDetails.Reason
        //ExFor:PdfDigitalSignatureDetails.SignatureDate
        //ExFor:PdfDigitalSignatureHashAlgorithm
        //ExFor:PdfSaveOptions.DigitalSignatureDetails
        //ExSummary:Shows how to sign a generated PDF using Aspose.Words.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Signed PDF contents.");

        // Load the certificate from disk
        // The other constructor overloads can be used to load certificates from different locations
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        // Pass the certificate and details to the save options class to sign with
        PdfSaveOptions options = new PdfSaveOptions();
        Date signingTime = new Date();
        options.setDigitalSignatureDetails(new PdfDigitalSignatureDetails(certificateHolder, "Test Signing", "Aspose Office", signingTime));

        // We can use this attribute to set a different hash algorithm
        options.getDigitalSignatureDetails().setHashAlgorithm(PdfDigitalSignatureHashAlgorithm.SHA_256);

        Assert.assertEquals(options.getDigitalSignatureDetails().getReason(), "Test Signing");
        Assert.assertEquals(options.getDigitalSignatureDetails().getLocation(), "Aspose Office");
        Assert.assertEquals(options.getDigitalSignatureDetails().getSignatureDate(), signingTime);

        doc.save(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignature.pdf", options);
        //ExEnd
    }

    @Test(enabled = false)
    public void pdfDigitalSignatureTimestamp() throws Exception {
        //ExStart
        //ExFor:PdfDigitalSignatureDetails.TimestampSettings
        //ExFor:PdfDigitalSignatureTimestampSettings
        //ExFor:PdfDigitalSignatureTimestampSettings.#ctor
        //ExFor:PdfDigitalSignatureTimestampSettings.#ctor(String,String,String)
        //ExFor:PdfDigitalSignatureTimestampSettings.#ctor(String,String,String,TimeSpan)
        //ExFor:PdfDigitalSignatureTimestampSettings.Password
        //ExFor:PdfDigitalSignatureTimestampSettings.ServerUrl
        //ExFor:PdfDigitalSignatureTimestampSettings.Timeout
        //ExFor:PdfDigitalSignatureTimestampSettings.UserName
        //ExSummary:Shows how to sign a generated PDF and timestamp it using Aspose.Words.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Signed PDF contents.");

        // Create a digital signature for the document that we will save
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");
        PdfSaveOptions options = new PdfSaveOptions();
        options.setDigitalSignatureDetails(new PdfDigitalSignatureDetails(certificateHolder, "Test Signing", "Aspose Office", new Date()));

        // We can set a verified timestamp for our signature as well, with a valid timestamp authority
        options.getDigitalSignatureDetails().setTimestampSettings(new PdfDigitalSignatureTimestampSettings("https://freetsa.org/tsr", "JohnDoe", "MyPassword"));

        // The default lifespan of the timestamp is 100 seconds
        Assert.assertEquals(options.getDigitalSignatureDetails().getTimestampSettings().getTimeout(), 100000);

        // We can set our own timeout period via the constructor
        options.getDigitalSignatureDetails().setTimestampSettings(new PdfDigitalSignatureTimestampSettings("https://freetsa.org/tsr", "JohnDoe", "MyPassword", (long) 30.0));

        Assert.assertEquals(options.getDigitalSignatureDetails().getTimestampSettings().getTimeout(), 30);
        Assert.assertEquals(options.getDigitalSignatureDetails().getTimestampSettings().getServerUrl(), "https://freetsa.org/tsr");
        Assert.assertEquals(options.getDigitalSignatureDetails().getTimestampSettings().getUserName(), "JohnDoe");
        Assert.assertEquals(options.getDigitalSignatureDetails().getTimestampSettings().getPassword(), "MyPassword");

        doc.save(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf", options);
        //ExEnd
    }

    @Test(dataProvider = "renderMetafileDataProvider")
    public void renderMetafile(/*EmfPlusDualRenderingMode*/int renderingMode) throws Exception {
        //ExStart
        //ExFor:EmfPlusDualRenderingMode
        //ExFor:MetafileRenderingOptions.EmfPlusDualRenderingMode
        //ExFor:MetafileRenderingOptions.UseEmfEmbeddedToWmf
        //ExSummary:Shows how to adjust EMF (Enhanced Windows Metafile) rendering options when saving to PDF.
        Document doc = new Document(getMyDir() + "EMF.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.getMetafileRenderingOptions().setEmfPlusDualRenderingMode(renderingMode);
        saveOptions.getMetafileRenderingOptions().setUseEmfEmbeddedToWmf(true);

        doc.save(getArtifactsDir() + "PdfSaveOptions.RenderMetafile.pdf", saveOptions);
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "renderMetafileDataProvider")
    public static Object[][] renderMetafileDataProvider() throws Exception {
        return new Object[][]
                {
                        {EmfPlusDualRenderingMode.EMF},
                        {EmfPlusDualRenderingMode.EMF_PLUS},
                        {EmfPlusDualRenderingMode.EMF_PLUS_WITH_FALLBACK},
                };
    }

    @Test(groups = "SkipMono")
    public void dml3DEffectsRenderingModeTest() throws Exception {
        Document doc = new Document(getMyDir() + "DrawingML shape 3D effects.docx");

        RenderCallback warningCallback = new RenderCallback();
        doc.setWarningCallback(warningCallback);

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setDml3DEffectsRenderingMode(Dml3DEffectsRenderingMode.ADVANCED);

        doc.save(getArtifactsDir() + "PdfSaveOptions.Dml3DEffectsRenderingModeTest.pdf", saveOptions);

        Assert.assertEquals(warningCallback.Count(), 43);
    }

    public static class RenderCallback implements IWarningCallback {
        public void warning(WarningInfo info) {
            System.out.println(MessageFormat.format("{0}: {1}.", info.getWarningType(), info.getDescription()));
            mWarnings.add(info);
        }

        public int Count() {
            return mWarnings.size();
        }

        private static ArrayList<WarningInfo> mWarnings = new ArrayList<>();
    }
}
