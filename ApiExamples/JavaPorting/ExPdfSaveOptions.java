// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.ms;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.StyleIdentifier;
import org.testng.Assert;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.BreakType;
import com.aspose.words.PdfImageCompression;
import com.aspose.words.PdfCompliance;
import com.aspose.words.PdfImageColorSpaceExportMode;
import com.aspose.ms.System.IO.Stream;
import com.aspose.words.ColorMode;
import com.aspose.words.SaveOptions;
import com.aspose.words.MetafileRenderingOptions;
import com.aspose.words.MetafileRenderingMode;
import com.aspose.words.IWarningCallback;
import com.aspose.words.WarningInfo;
import com.aspose.words.WarningType;
import com.aspose.ms.System.msConsole;
import com.aspose.words.WarningInfoCollection;
import com.aspose.words.HeaderFooterBookmarksExportMode;
import com.aspose.words.PdfPageMode;
import com.aspose.ms.System.Globalization.msCultureInfo;
import com.aspose.words.PdfTextCompression;
import com.aspose.words.Section;
import com.aspose.words.MultiplePagesType;
import com.aspose.ms.System.StringComparison;
import com.aspose.words.PdfZoomBehavior;
import com.aspose.words.PdfCustomPropertiesExport;
import com.aspose.words.DmlEffectsRenderingMode;
import com.aspose.words.DmlRenderingMode;
import java.awt.image.BufferedImage;
import com.aspose.BitmapPal;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.words.Dml3DEffectsRenderingMode;
import com.aspose.words.WarningSource;
import java.util.ArrayList;
import com.aspose.words.CertificateHolder;
import com.aspose.ms.System.DateTime;
import com.aspose.words.PdfDigitalSignatureDetails;
import com.aspose.words.PdfDigitalSignatureHashAlgorithm;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.PdfDigitalSignatureTimestampSettings;
import com.aspose.ms.System.TimeSpan;
import com.aspose.words.EmfPlusDualRenderingMode;
import org.testng.annotations.DataProvider;


@Test
class ExPdfSaveOptions !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void createMissingOutlineLevels() throws Exception
    {
        //ExStart
        //ExFor:OutlineOptions.CreateMissingOutlineLevels
        //ExFor:ParagraphFormat.IsHeading
        //ExFor:PdfSaveOptions.OutlineOptions
        //ExFor:PdfSaveOptions.SaveFormat
        //ExSummary:Shows how to create PDF document outline entries for headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create TOC entries
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
        bookmarkEditor.BindPdf(getArtifactsDir() + "PdfSaveOptions.CreateMissingOutlineLevels.pdf");

        // Get all bookmarks from the document
        Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

        Assert.AreEqual(11, bookmarks.Count);
            }

    @Test
    public void tableHeadingOutlines() throws Exception
    {
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

                Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.TableHeadingOutlines.pdf");

        Assert.AreEqual(1, pdfDoc.Outlines.Count);
        Assert.AreEqual("Heading 1", pdfDoc.Outlines[1].Title);

        TableAbsorber tableAbsorber = new TableAbsorber();
        tableAbsorber.Visit(pdfDoc.Pages[1]);

        Assert.AreEqual("Heading 1", tableAbsorber.TableList[0].RowList[0].CellList[0].TextFragments[1].Text);
        Assert.AreEqual("Cell 1", tableAbsorber.TableList[0].RowList[1].CellList[0].TextFragments[1].Text);
            }

    @Test (groups = "SkipMono", dataProvider = "updateFieldsDataProvider")
    public void updateFields(boolean doUpdateFields) throws Exception
    {
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

                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.UpdateFields.pdf");

        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
        pdfDocument.Pages.Accept(textFragmentAbsorber);

        if (doUpdateFields)
            Assert.AreEqual("Page 1 of 2", textFragmentAbsorber.TextFragments[1].Text);
        else
            Assert.AreEqual("Page  of ", textFragmentAbsorber.TextFragments[1].Text);
            }

	//JAVA-added data provider for test method
	@DataProvider(name = "updateFieldsDataProvider")
	public static Object[][] updateFieldsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void imageCompression() throws Exception
    {
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

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.ImageCompression.pdf");
        Stream pdfDocImageStream = pdfDocument.Pages[1].Resources.Images[1].ToStream();

        try /*JAVA: was using*/
        {
            TestUtil.verifyImage(2467, 1500, pdfDocImageStream);
        }
        finally { if (pdfDocImageStream != null) pdfDocImageStream.close(); }
        
        pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.ImageCompression.PDF_A_1_B.pdf");
        pdfDocImageStream = pdfDocument.Pages[1].Resources.Images[1].ToStream();

        try /*JAVA: was using*/
        {
            Assert.<IllegalArgumentException>Throws(() => TestUtil.verifyImage(2467, 1500, pdfDocImageStream));
        }
        finally { if (pdfDocImageStream != null) pdfDocImageStream.close(); }

        pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.ImageCompression.PDF_A_1_A.pdf");
        pdfDocImageStream = pdfDocument.Pages[1].Resources.Images[1].ToStream();
        
        try /*JAVA: was using*/
        {
            TestUtil.verifyImage(2467, 1500, pdfDocImageStream);
        }
        finally { if (pdfDocImageStream != null) pdfDocImageStream.close(); }
    }

    @Test
    public void colorRendering() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions
        //ExFor:ColorMode
        //ExFor:FixedPageSaveOptions.ColorMode
        //ExSummary:Shows how change image color with save options property.
        Document doc = new Document(getMyDir() + "Images.docx");

        // Configure PdfSaveOptions to save every image in the input document in greyscale during conversion
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions(); { pdfSaveOptions.setColorMode(ColorMode.GRAYSCALE); }
        
        doc.save(getArtifactsDir() + "PdfSaveOptions.ColorRendering.pdf", pdfSaveOptions);
        //ExEnd

                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.ColorRendering.pdf");
        XImage pdfDocImage = pdfDocument.Pages[1].Resources.Images[1];

        Assert.AreEqual(1506, pdfDocImage.Width);
        Assert.AreEqual(918, pdfDocImage.Height);
        Assert.AreEqual(ColorType.Grayscale, pdfDocImage.GetColorType());
            }

    @Test
    public void windowsBarPdfTitle() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.DisplayDocTitle
        //ExSummary:Shows how to display title of the document as title bar.
        Document doc = new Document(getMyDir() + "Rendering.docx");
        doc.getBuiltInDocumentProperties().setTitle("Windows bar pdf title");
        
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions(); { pdfSaveOptions.setDisplayDocTitle(true); }

        doc.save(getArtifactsDir() + "PdfSaveOptions.WindowsBarPdfTitle.pdf", pdfSaveOptions);
        //ExEnd

                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.WindowsBarPdfTitle.pdf");

        Assert.IsTrue(pdfDocument.DisplayDocTitle);
        Assert.AreEqual("Windows bar pdf title", pdfDocument.Info.Title);
            }

    @Test
    public void memoryOptimization() throws Exception
    {
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

    @Test (dataProvider = "escapeUriDataProvider")
    public void escapeUri(String uri, String result, boolean isEscaped) throws Exception
    {
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

                Aspose.Pdf.Document pdfDocument =
            new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.EscapedUri.pdf");

        // Get first page
        Page page = pdfDocument.Pages[1];
        // Get the first link annotation
        LinkAnnotation linkAnnot = (LinkAnnotation)page.Annotations[1];

        JavascriptAction action = (JavascriptAction)linkAnnot.Action;
        String uriText = action.Script;

        Assert.assertEquals(result, uriText);
            }

	//JAVA-added data provider for test method
	@DataProvider(name = "escapeUriDataProvider")
	public static Object[][] escapeUriDataProvider() throws Exception
	{
		return new Object[][]
		{
			{"https://www.google.com/search?q= aspose",  "app.launchURL(\"https://www.google.com/search?q=%20aspose\", true);",  true},
			{"https://www.google.com/search?q=%20aspose",  "app.launchURL(\"https://www.google.com/search?q=%20aspose\", true);",  true},
			{"https://www.google.com/search?q= aspose",  "app.launchURL(\"https://www.google.com/search?q= aspose\", true);",  false},
			{"https://www.google.com/search?q=%20aspose",  "app.launchURL(\"https://www.google.com/search?q=%20aspose\", true);",  false},
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
    @Test (groups = "SkipMono") //ExSkip
    public void handleBinaryRasterWarnings() throws Exception
    {
        Document doc = new Document(getMyDir() + "WMF with image.docx");

        MetafileRenderingOptions metafileRenderingOptions =
            new MetafileRenderingOptions();
            {
                metafileRenderingOptions.setEmulateRasterOperations(false);
                metafileRenderingOptions.setRenderingMode(MetafileRenderingMode.VECTOR_WITH_FALLBACK);
            }

        // If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words
        // renders this metafile to a bitmap
        HandleDocumentWarnings callback = new HandleDocumentWarnings();
        doc.setWarningCallback(callback);

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setMetafileRenderingOptions(metafileRenderingOptions);

        doc.save(getArtifactsDir() + "PdfSaveOptions.HandleBinaryRasterWarnings.pdf", saveOptions);

        Assert.assertEquals(1, callback.Warnings.getCount());
        Assert.assertEquals("'R2_XORPEN' binary raster operation is partly supported.",
            callback.Warnings.get(0).getDescription());
    }

    public static class HandleDocumentWarnings implements IWarningCallback
    {
        /// <summary>
        /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
        /// potential issue during document processing. The callback can be set to listen for warnings generated during document
        /// load and/or document save.
        /// </summary>
        public void warning(WarningInfo info)
        {
            // For now type of warnings about unsupported metafile records changed from
            // DataLoss/UnexpectedContent to MinorFormattingLoss
            if (info.getWarningType() == WarningType.MINOR_FORMATTING_LOSS)
            {
                System.out.println("Unsupported operation: " + info.getDescription());
                Warnings.warning(info);
            }
        }

        public WarningInfoCollection Warnings = new WarningInfoCollection();
    }
    //ExEnd

    @Test (dataProvider = "headerFooterBookmarksExportModeDataProvider")
    public void headerFooterBookmarksExportMode(/*HeaderFooterBookmarksExportMode*/int headerFooterBookmarksExportMode) throws Exception
    {
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
        {
            saveOptions.setHeaderFooterBookmarksExportMode(headerFooterBookmarksExportMode);
            saveOptions.setOutlineOptions({ saveOptions.getOutlineOptions().setDefaultBookmarksOutlineLevel(1); });
            saveOptions.setPageMode(PdfPageMode.USE_OUTLINES);
        }
        doc.save(getArtifactsDir() + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf", saveOptions);
        //ExEnd

                Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf");
        String inputDocLocaleName = new msCultureInfo(doc.getStyles().getDefaultFont().getLocaleId()).getName();

        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
        pdfDoc.Pages.Accept(textFragmentAbsorber);
        switch (headerFooterBookmarksExportMode)
        {
            case com.aspose.words.HeaderFooterBookmarksExportMode.NONE:
                TestUtil.fileContainsString($"<</Type /Catalog/Pages 3 0 R/Lang({inputDocLocaleName})>>\r\n", 
                    getArtifactsDir() + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf");

                Assert.AreEqual(0, pdfDoc.Outlines.Count);
                break;
            case com.aspose.words.HeaderFooterBookmarksExportMode.FIRST:
            case com.aspose.words.HeaderFooterBookmarksExportMode.ALL:
                TestUtil.fileContainsString($"<</Type /Catalog/Pages 3 0 R/Outlines 13 0 R/PageMode /UseOutlines/Lang({inputDocLocaleName})>>", 
                    getArtifactsDir() + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf");

                OutlineCollection outlineItemCollection = pdfDoc.Outlines;

                Assert.AreEqual(4, outlineItemCollection.Count);
                Assert.AreEqual("Bookmark_1", outlineItemCollection[1].Title);
                Assert.AreEqual("1 XYZ 233 806 0", outlineItemCollection[1].Destination.ToString());

                Assert.AreEqual("Bookmark_2", outlineItemCollection[2].Title);
                Assert.AreEqual("1 XYZ 84 47 0", outlineItemCollection[2].Destination.ToString());

                Assert.AreEqual("Bookmark_3", outlineItemCollection[3].Title);
                Assert.AreEqual("2 XYZ 85 806 0", outlineItemCollection[3].Destination.ToString());

                Assert.AreEqual("Bookmark_4", outlineItemCollection[4].Title);
                Assert.AreEqual("2 XYZ 85 48 0", outlineItemCollection[4].Destination.ToString());
                break;
        }
            }

	//JAVA-added data provider for test method
	@DataProvider(name = "headerFooterBookmarksExportModeDataProvider")
	public static Object[][] headerFooterBookmarksExportModeDataProvider() throws Exception
	{
		return new Object[][]
		{
			{com.aspose.words.HeaderFooterBookmarksExportMode.NONE},
			{com.aspose.words.HeaderFooterBookmarksExportMode.FIRST},
			{com.aspose.words.HeaderFooterBookmarksExportMode.ALL},
		};
	}

    @Test
    public void unsupportedImageFormatWarning() throws Exception
    {
        Document doc = new Document(getMyDir() + "Corrupted image.docx");

        SaveWarningCallback saveWarningCallback = new SaveWarningCallback();
        doc.setWarningCallback(saveWarningCallback);

        doc.save(getArtifactsDir() + "PdfSaveOption.UnsupportedImageFormatWarning.pdf", SaveFormat.PDF);

        Assert.That(saveWarningCallback.SaveWarnings.get(0).getDescription(),
            Is.EqualTo("Image can not be processed. Possibly unsupported image format."));
    }

    public static class SaveWarningCallback implements IWarningCallback
    {
        public void warning(WarningInfo info)
        {
            if (info.getWarningType() == WarningType.MINOR_FORMATTING_LOSS)
            {
                System.out.println("{info.WarningType}: {info.Description}.");
                SaveWarnings.warning(info);
            }
        }

        WarningInfoCollection SaveWarnings = new WarningInfoCollection();
	}
	
	@Test (dataProvider = "fontsScaledToMetafileSizeDataProvider")
    public void fontsScaledToMetafileSize(boolean doScaleWmfFonts) throws Exception
    {
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

                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.FontsScaledToMetafileSize.pdf");
        TextFragmentAbsorber textAbsorber = new TextFragmentAbsorber();

        pdfDocument.Pages[1].Accept(textAbsorber);
        Rectangle textFragmentRectangle = textAbsorber.TextFragments[3].Rectangle;

        Assert.AreEqual(doScaleWmfFonts ? 1.589d : 5.045d, textFragmentRectangle.Width, 0.001d);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "fontsScaledToMetafileSizeDataProvider")
	public static Object[][] fontsScaledToMetafileSizeDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "additionalTextPositioningDataProvider")
    public void additionalTextPositioning(boolean applyAdditionalTextPositioning) throws Exception
    {
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

                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.AdditionalTextPositioning.pdf");
        TextFragmentAbsorber textAbsorber = new TextFragmentAbsorber();

        pdfDocument.Pages[1].Accept(textAbsorber);

        SetGlyphsPositionShowText tjOperator = (SetGlyphsPositionShowText)textAbsorber.TextFragments[1].Page.Contents[96];

        Assert.AreEqual(
            applyAdditionalTextPositioning
                ? "[0 (s) 0 (e) 1 (g) 0 (m) 0 (e) 0 (n) 0 (t) 0 (s) 0 ( ) 1 (o) 0 (f) 0 ( ) 1 (t) 0 (e) 0 (x) 0 (t)] TJ"
                : "[(se) 1 (gments ) 1 (of ) 1 (text)] TJ", tjOperator.ToString());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "additionalTextPositioningDataProvider")
	public static Object[][] additionalTextPositioningDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "saveAsPdfBookFoldDataProvider")
    public void saveAsPdfBookFold(boolean doRenderTextAsBookfold) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.UseBookFoldPrintingSettings
        //ExSummary:Shows how to save a document to the PDF format in the form of a book fold.
        // Open a document with multiple paragraphs
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        // Configure both page setup and PdfSaveOptions to create a book fold
        for (Section s : (Iterable<Section>) doc.getSections())
        {
            s.getPageSetup().setMultiplePages(MultiplePagesType.BOOK_FOLD_PRINTING);
        }

        PdfSaveOptions options = new PdfSaveOptions();
        options.setUseBookFoldPrintingSettings(doRenderTextAsBookfold);

        // In order to make a booklet, we will need to print this document, stack the pages
        // in the order they come out of the printer and then fold down the middle
        doc.save(getArtifactsDir() + "PdfSaveOptions.SaveAsPdfBookFold.pdf", options);
        //ExEnd

                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.SaveAsPdfBookFold.pdf");
        TextAbsorber textAbsorber = new TextAbsorber();

        pdfDocument.Pages.Accept(textAbsorber);

        if (doRenderTextAsBookfold)
        {
            Assert.True(textAbsorber.Text.IndexOf("Heading #1", StringComparison.ORDINAL) < textAbsorber.Text.IndexOf("Heading #2", StringComparison.ORDINAL));
            Assert.True(textAbsorber.Text.IndexOf("Heading #2", StringComparison.ORDINAL) < textAbsorber.Text.IndexOf("Heading #3", StringComparison.ORDINAL));
            Assert.True(textAbsorber.Text.IndexOf("Heading #3", StringComparison.ORDINAL) < textAbsorber.Text.IndexOf("Heading #4", StringComparison.ORDINAL));
            Assert.True(textAbsorber.Text.IndexOf("Heading #4", StringComparison.ORDINAL) < textAbsorber.Text.IndexOf("Heading #5", StringComparison.ORDINAL));
            Assert.True(textAbsorber.Text.IndexOf("Heading #5", StringComparison.ORDINAL) < textAbsorber.Text.IndexOf("Heading #6", StringComparison.ORDINAL));
            Assert.True(textAbsorber.Text.IndexOf("Heading #6", StringComparison.ORDINAL) < textAbsorber.Text.IndexOf("Heading #7", StringComparison.ORDINAL));
            Assert.False(textAbsorber.Text.IndexOf("Heading #7", StringComparison.ORDINAL) < textAbsorber.Text.IndexOf("Heading #8", StringComparison.ORDINAL));
            Assert.True(textAbsorber.Text.IndexOf("Heading #8", StringComparison.ORDINAL) < textAbsorber.Text.IndexOf("Heading #9", StringComparison.ORDINAL));
            Assert.False(textAbsorber.Text.IndexOf("Heading #9", StringComparison.ORDINAL) < textAbsorber.Text.IndexOf("Heading #10", StringComparison.ORDINAL));
        }
        else
        {
            Assert.True(textAbsorber.Text.IndexOf("Heading #1", StringComparison.ORDINAL) < textAbsorber.Text.IndexOf("Heading #2", StringComparison.ORDINAL));
            Assert.True(textAbsorber.Text.IndexOf("Heading #2", StringComparison.ORDINAL) < textAbsorber.Text.IndexOf("Heading #3", StringComparison.ORDINAL));
            Assert.True(textAbsorber.Text.IndexOf("Heading #3", StringComparison.ORDINAL) < textAbsorber.Text.IndexOf("Heading #4", StringComparison.ORDINAL));
            Assert.True(textAbsorber.Text.IndexOf("Heading #4", StringComparison.ORDINAL) < textAbsorber.Text.IndexOf("Heading #5", StringComparison.ORDINAL));
            Assert.True(textAbsorber.Text.IndexOf("Heading #5", StringComparison.ORDINAL) < textAbsorber.Text.IndexOf("Heading #6", StringComparison.ORDINAL));
            Assert.True(textAbsorber.Text.IndexOf("Heading #6", StringComparison.ORDINAL) < textAbsorber.Text.IndexOf("Heading #7", StringComparison.ORDINAL));
            Assert.True(textAbsorber.Text.IndexOf("Heading #7", StringComparison.ORDINAL) < textAbsorber.Text.IndexOf("Heading #8", StringComparison.ORDINAL));
            Assert.True(textAbsorber.Text.IndexOf("Heading #8", StringComparison.ORDINAL) < textAbsorber.Text.IndexOf("Heading #9", StringComparison.ORDINAL));
            Assert.True(textAbsorber.Text.IndexOf("Heading #9", StringComparison.ORDINAL) < textAbsorber.Text.IndexOf("Heading #10", StringComparison.ORDINAL));
        }
            }

	//JAVA-added data provider for test method
	@DataProvider(name = "saveAsPdfBookFoldDataProvider")
	public static Object[][] saveAsPdfBookFoldDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void zoomBehaviour() throws Exception
    {
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

                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.ZoomBehaviour.pdf");
        GoToAction action = (GoToAction)pdfDocument.OpenAction;

        Assert.AreEqual(0.25d, (ms.as(action.Destination, XYZExplicitDestination.class)).Zoom);
            }

    @Test (dataProvider = "pageModeDataProvider")
    public void pageMode(/*PdfPageMode*/int pageMode) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.PageMode
        //ExFor:PdfPageMode
        //ExSummary:Shows how to set instructions for some PDF readers to follow when opening an output document.
        Document doc = new Document(getMyDir() + "Document.docx");

        PdfSaveOptions options = new PdfSaveOptions();
        options.setPageMode(pageMode);

        doc.save(getArtifactsDir() + "PdfSaveOptions.PageMode.pdf", options);
        //ExEnd
        
        String docLocaleName = new msCultureInfo(doc.getStyles().getDefaultFont().getLocaleId()).getName();

        switch (pageMode)
        {
            case PdfPageMode.FULL_SCREEN:
                TestUtil.fileContainsString($"<</Type /Catalog/Pages 3 0 R/PageMode /FullScreen/Lang({docLocaleName})>>\r\n", getArtifactsDir() + "PdfSaveOptions.PageMode.pdf");
                break;
            case PdfPageMode.USE_THUMBS:
                TestUtil.fileContainsString($"<</Type /Catalog/Pages 3 0 R/PageMode /UseThumbs/Lang({docLocaleName})>>", getArtifactsDir() + "PdfSaveOptions.PageMode.pdf");
                break;
            case PdfPageMode.USE_OC:
                TestUtil.fileContainsString($"<</Type /Catalog/Pages 3 0 R/PageMode /UseOC/Lang({docLocaleName})>>\r\n", getArtifactsDir() + "PdfSaveOptions.PageMode.pdf");
                break;
            case PdfPageMode.USE_OUTLINES:
            case PdfPageMode.USE_NONE:
                TestUtil.fileContainsString($"<</Type /Catalog/Pages 3 0 R/Lang({docLocaleName})>>\r\n", getArtifactsDir() + "PdfSaveOptions.PageMode.pdf");
                break;
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "pageModeDataProvider")
	public static Object[][] pageModeDataProvider() throws Exception
	{
		return new Object[][]
		{
			{PdfPageMode.FULL_SCREEN},
			{PdfPageMode.USE_THUMBS},
			{PdfPageMode.USE_OC},
			{PdfPageMode.USE_OUTLINES},
			{PdfPageMode.USE_NONE},
		};
	}

    @Test (dataProvider = "noteHyperlinksDataProvider")
    public void noteHyperlinks(boolean doCreateHyperlinks) throws Exception
    {
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

        if (doCreateHyperlinks)
        {
            TestUtil.fileContainsString("<</Type /Annot/Subtype /Link/Rect [157.80099487 720.90106201 159.35600281 733.55004883]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 85 677 0]>>", 
                getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString("<</Type /Annot/Subtype /Link/Rect [202.16900635 720.90106201 206.06201172 733.55004883]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 85 79 0]>>", 
                getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString("<</Type /Annot/Subtype /Link/Rect [212.23199463 699.2510376 215.34199524 711.90002441]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 85 654 0]>>", 
                getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString("<</Type /Annot/Subtype /Link/Rect [258.15499878 699.2510376 262.04800415 711.90002441]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 85 68 0]>>", 
                getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString("<</Type /Annot/Subtype /Link/Rect [85.05000305 68.19905853 88.66500092 79.69805908]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 202 733 0]>>", 
                getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString("<</Type /Annot/Subtype /Link/Rect [85.05000305 56.70005798 88.66500092 68.19905853]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 258 711 0]>>", 
                getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString("<</Type /Annot/Subtype /Link/Rect [85.05000305 666.10205078 86.4940033 677.60107422]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 157 733 0]>>", 
                getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString("<</Type /Annot/Subtype /Link/Rect [85.05000305 643.10406494 87.93800354 654.60308838]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 212 711 0]>>", 
                getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
        }
        else
        {
            Assert.<AssertionError>Throws(() => TestUtil.fileContainsString("<</Type /Annot/Subtype /Link/Rect", getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf"));
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "noteHyperlinksDataProvider")
	public static Object[][] noteHyperlinksDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "customPropertiesExportDataProvider")
    public void customPropertiesExport(/*PdfCustomPropertiesExport*/int pdfCustomPropertiesExportMode) throws Exception
    {
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

        switch (pdfCustomPropertiesExportMode)
        {
            case PdfCustomPropertiesExport.NONE:
                Assert.<AssertionError>Throws(() => TestUtil.fileContainsString(doc.getCustomDocumentProperties().get(0).getName(), 
                    getArtifactsDir() + "PdfSaveOptions.CustomPropertiesExport.pdf"));
                Assert.<AssertionError>Throws(() => TestUtil.fileContainsString("<</Type /Metadata/Subtype /XML/Length 8 0 R/Filter /FlateDecode>>", 
                    getArtifactsDir() + "PdfSaveOptions.CustomPropertiesExport.pdf"));
                break;
            case PdfCustomPropertiesExport.STANDARD:
                TestUtil.fileContainsString(doc.getCustomDocumentProperties().get(0).getName(), getArtifactsDir() + "PdfSaveOptions.CustomPropertiesExport.pdf");
                break;
            case PdfCustomPropertiesExport.METADATA:
                TestUtil.fileContainsString("<</Type /Metadata/Subtype /XML/Length 8 0 R/Filter /FlateDecode>>", getArtifactsDir() + "PdfSaveOptions.CustomPropertiesExport.pdf");
                break;
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "customPropertiesExportDataProvider")
	public static Object[][] customPropertiesExportDataProvider() throws Exception
	{
		return new Object[][]
		{
			{PdfCustomPropertiesExport.NONE},
			{PdfCustomPropertiesExport.STANDARD},
			{PdfCustomPropertiesExport.METADATA},
		};
	}

    @Test (dataProvider = "drawingMLEffectsDataProvider")
    public void drawingMLEffects(/*DmlEffectsRenderingMode*/int effectsRenderingMode) throws Exception
    {
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

                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.DrawingMLEffects.pdf");

        ImagePlacementAbsorber imb = new ImagePlacementAbsorber();
        imb.Visit(pdfDocument.Pages[1]);

        TableAbsorber ttb = new TableAbsorber();
        ttb.Visit(pdfDocument.Pages[1]);

        switch (effectsRenderingMode)
        {
            case DmlEffectsRenderingMode.NONE:
            case DmlEffectsRenderingMode.SIMPLIFIED:
                TestUtil.fileContainsString("4 0 obj\r\n" +
                                            "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                    getArtifactsDir() + "PdfSaveOptions.DrawingMLEffects.pdf");
                Assert.AreEqual(0, imb.ImagePlacements.Count);
                Assert.AreEqual(28, ttb.TableList.Count);
                break;
            case DmlEffectsRenderingMode.FINE:
                TestUtil.fileContainsString("4 0 obj\r\n<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R>>/XObject<</X1 9 0 R/X2 10 0 R/X3 11 0 R/X4 12 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                    getArtifactsDir() + "PdfSaveOptions.DrawingMLEffects.pdf");
                Assert.AreEqual(21, imb.ImagePlacements.Count);
                Assert.AreEqual(4, ttb.TableList.Count);
                break;
        }
            }

	//JAVA-added data provider for test method
	@DataProvider(name = "drawingMLEffectsDataProvider")
	public static Object[][] drawingMLEffectsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{DmlEffectsRenderingMode.NONE},
			{DmlEffectsRenderingMode.SIMPLIFIED},
			{DmlEffectsRenderingMode.FINE},
		};
	}

    @Test (dataProvider = "drawingMLFallbackDataProvider")
    public void drawingMLFallback(/*DmlRenderingMode*/int dmlRenderingMode) throws Exception
    {
        //ExStart
        //ExFor:DmlRenderingMode
        //ExFor:SaveOptions.DmlRenderingMode
        //ExSummary:Shows how to render fallback shapes when saving to Pdf.
        Document doc = new Document(getMyDir() + "DrawingML shape fallbacks.docx");

        PdfSaveOptions options = new PdfSaveOptions();
        options.setDmlRenderingMode(dmlRenderingMode);

        doc.save(getArtifactsDir() + "PdfSaveOptions.DrawingMLFallback.pdf", options);
        //ExEnd

        switch (dmlRenderingMode)
        {
            case DmlRenderingMode.DRAWING_ML:
                TestUtil.fileContainsString("<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABA 10 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                    getArtifactsDir() + "PdfSaveOptions.DrawingMLFallback.pdf");
                break;
            case DmlRenderingMode.FALLBACK:
                TestUtil.fileContainsString("4 0 obj\r\n<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABC 12 0 R>>/ExtGState<</GS1 9 0 R/GS2 10 0 R>>>>/Group ",
                    getArtifactsDir() + "PdfSaveOptions.DrawingMLFallback.pdf");
                break;
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "drawingMLFallbackDataProvider")
	public static Object[][] drawingMLFallbackDataProvider() throws Exception
	{
		return new Object[][]
		{
			{DmlRenderingMode.FALLBACK},
			{DmlRenderingMode.DRAWING_ML},
		};
	}

    @Test (dataProvider = "exportDocumentStructureDataProvider")
    public void exportDocumentStructure(boolean doExportStructure) throws Exception
    {
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

        if (doExportStructure)
        {
            TestUtil.fileContainsString("4 0 obj\r\n" +
                                        "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABC 12 0 R>>/ExtGState<</GS1 9 0 R/GS2 10 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>/StructParents 0/Tabs /S>>",
                getArtifactsDir() + "PdfSaveOptions.ExportDocumentStructure.pdf");
        }
        else
        {
            TestUtil.fileContainsString("4 0 obj\r\n" +
                                        "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABA 10 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                getArtifactsDir() + "PdfSaveOptions.ExportDocumentStructure.pdf");
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportDocumentStructureDataProvider")
	public static Object[][] exportDocumentStructureDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "preblendImagesDataProvider")
    public void preblendImages(boolean doPreblendImages) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.PreblendImages
        //ExSummary:Shows how to preblend images with transparent backgrounds.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        BufferedImage img = BitmapPal.loadNativeImage(getImageDir() + "Transparent background logo.png");
        builder.insertImage(img);

        // Setting this flag in a SaveOptions object may change the quality and size of the output .pdf
        // because of the way some images are rendered
        PdfSaveOptions options = new PdfSaveOptions();
        options.setPreblendImages(doPreblendImages);

        doc.save(getArtifactsDir() + "PdfSaveOptions.PreblendImages.pdf", options);
        //ExEnd

        testPreblendImages(getArtifactsDir() + "PdfSaveOptions.PreblendImages.pdf", doPreblendImages);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "preblendImagesDataProvider")
	public static Object[][] preblendImagesDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    private void testPreblendImages(String outFileName, boolean doPreblendImages) throws Exception
    {
        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(outFileName);
        XImage image = pdfDocument.Pages[1].Resources.Images[1];

        MemoryStream stream = new MemoryStream();
        try /*JAVA: was using*/
        {
            image.Save(stream);

            if (doPreblendImages)
            {
                TestUtil.fileContainsString("9 0 obj\r\n20849 ", outFileName);
                Assert.assertEquals(17898, stream.getLength());
            }
            else
            {
                TestUtil.fileContainsString("9 0 obj\r\n19289 ", outFileName);
                Assert.assertEquals(19216, stream.getLength());
            }
        }
        finally { if (stream != null) stream.close(); }
    }

    @Test
    public void interpolateImages() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.InterpolateImages
        //ExSummary:Shows how to improve the quality of an image in the rendered documents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        BufferedImage img = BitmapPal.loadNativeImage(getImageDir() + "Transparent background logo.png");
        builder.insertImage(img);
        
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setInterpolateImages(true);
        
        doc.save(getArtifactsDir() + "PdfSaveOptions.InterpolateImages.pdf", saveOptions);
        //ExEnd
    }

    @Test (groups = "SkipMono")
    public void dml3DEffectsRenderingModeTest() throws Exception
    {
        Document doc = new Document(getMyDir() + "DrawingML shape 3D effects.docx");
        
        RenderCallback warningCallback = new RenderCallback();
        doc.setWarningCallback(warningCallback);
        
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setDml3DEffectsRenderingMode(Dml3DEffectsRenderingMode.ADVANCED);
        
        doc.save(getArtifactsDir() + "PdfSaveOptions.Dml3DEffectsRenderingModeTest.pdf", saveOptions);

        Assert.AreEqual(43, warningCallback.Count);
    }

    public static class RenderCallback implements IWarningCallback
    {
        public void warning(WarningInfo info)
        {
            System.out.println("{info.WarningType}: {info.Description}.");
            mWarnings.Add(info);
        }

         !!Autoporter error: Indexer ApiExamples.ExPdfSaveOptions.RenderCallback.Item(int) hasn't both getter and setter!private mWarnings.CountmWarnings;

        /// <summary>
        /// Returns true if a warning with the specified properties has been generated.
        /// </summary>
        public boolean contains(/*WarningSource*/int source, /*WarningType*/int type, String description)
        {
            return mWarnings.Any(warning => warning.Source == source && warning.WarningType == type && warning.Description == description);
        }

        private /*final*/ ArrayList<WarningInfo> mWarnings = new ArrayList<WarningInfo>();
    }


    @Test
    public void pdfDigitalSignature() throws Exception
    {
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
        DateTime signingTime = DateTime.getNow();
        options.setDigitalSignatureDetails(new PdfDigitalSignatureDetails(certificateHolder, "Test Signing", "Aspose Office", signingTime));

        // We can use this attribute to set a different hash algorithm
        options.getDigitalSignatureDetails().setHashAlgorithm(PdfDigitalSignatureHashAlgorithm.SHA_256);

        Assert.assertEquals("Test Signing", options.getDigitalSignatureDetails().getReason());
        Assert.assertEquals("Aspose Office", options.getDigitalSignatureDetails().getLocation());
        Assert.assertEquals(signingTime.toUniversalTime(), options.getDigitalSignatureDetails().getSignatureDateInternal());

        doc.save(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignature.pdf", options);
        //ExEnd
        
        TestUtil.fileContainsString("6 0 obj\r\n" +
                                    "<</Type /Annot/Subtype /Widget/FT /Sig/DR <<>>/F 132/Rect [0 0 0 0]/V 7 0 R/P 4 0 R/T(þÿ\0A\u0000s\u0000p\u0000o\u0000s\0e\0D\u0000i\u0000g\u0000i\u0000t\0a\u0000l\u0000S\u0000i\u0000g\u0000n\0a\u0000t\u0000u\u0000r\0e)/AP <</N 8 0 R>>>>", 
            getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignature.pdf");

        Assert.assertFalse(FileFormatUtil.detectFileFormat(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignature.pdf").hasDigitalSignature());
    }

    @Test
    public void pdfDigitalSignatureTimestamp() throws Exception
    {
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
        options.setDigitalSignatureDetails(new PdfDigitalSignatureDetails(certificateHolder, "Test Signing", "Aspose Office", DateTime.getNow()));

        // We can set a verified timestamp for our signature as well, with a valid timestamp authority
        options.getDigitalSignatureDetails().setTimestampSettings(new PdfDigitalSignatureTimestampSettings("https://freetsa.org/tsr", "JohnDoe", "MyPassword"));

        // The default lifespan of the timestamp is 100 seconds
        Assert.assertEquals(100.0d, options.getDigitalSignatureDetails().getTimestampSettings().getTimeoutInternal().getTotalSeconds());

        // We can set our own timeout period via the constructor
        options.getDigitalSignatureDetails().setTimestampSettings(new PdfDigitalSignatureTimestampSettings("https://freetsa.org/tsr", "JohnDoe", "MyPassword", TimeSpan.fromMinutes(30.0)));

        Assert.assertEquals(1800.0d, options.getDigitalSignatureDetails().getTimestampSettings().getTimeoutInternal().getTotalSeconds());
        Assert.assertEquals("https://freetsa.org/tsr", options.getDigitalSignatureDetails().getTimestampSettings().getServerUrl());
        Assert.assertEquals("JohnDoe", options.getDigitalSignatureDetails().getTimestampSettings().getUserName());
        Assert.assertEquals("MyPassword", options.getDigitalSignatureDetails().getTimestampSettings().getPassword());

        doc.save(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf", options);
        //ExEnd

        Assert.assertFalse(FileFormatUtil.detectFileFormat(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf").hasDigitalSignature());
        TestUtil.fileContainsString("6 0 obj\r\n" +
                                    "<</Type /Annot/Subtype /Widget/FT /Sig/DR <<>>/F 132/Rect [0 0 0 0]/V 7 0 R/P 4 0 R/T(þÿ\0A\u0000s\u0000p\u0000o\u0000s\0e\0D\u0000i\u0000g\u0000i\u0000t\0a\u0000l\u0000S\u0000i\u0000g\u0000n\0a\u0000t\u0000u\u0000r\0e)/AP <</N 8 0 R>>>>", 
        getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf");
    }

    @Test (dataProvider = "renderMetafileDataProvider")
    public void renderMetafile(/*EmfPlusDualRenderingMode*/int renderingMode) throws Exception
    {
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

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.RenderMetafile.pdf");

        switch (renderingMode)
        {
            case EmfPlusDualRenderingMode.EMF:
            case EmfPlusDualRenderingMode.EMF_PLUS_WITH_FALLBACK:
                Assert.AreEqual(0, pdfDocument.Pages[1].Resources.Images.Count);
                TestUtil.fileContainsString("4 0 obj\r\n" +
                                            "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 595.29998779 841.90002441]/Resources<</Font<</FAAAAH 7 0 R/FAAABA 10 0 R/FAAABD 13 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                    getArtifactsDir() + "PdfSaveOptions.RenderMetafile.pdf");
                break;
            case EmfPlusDualRenderingMode.EMF_PLUS:
                Assert.AreEqual(1, pdfDocument.Pages[1].Resources.Images.Count);
                TestUtil.fileContainsString("4 0 obj\r\n" +
                                            "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 595.29998779 841.90002441]/Resources<</Font<</FAAAAH 7 0 R/FAAABB 11 0 R/FAAABE 14 0 R>>/XObject<</X1 9 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                    getArtifactsDir() + "PdfSaveOptions.RenderMetafile.pdf");
                break;
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "renderMetafileDataProvider")
	public static Object[][] renderMetafileDataProvider() throws Exception
	{
		return new Object[][]
		{
			{EmfPlusDualRenderingMode.EMF},
			{EmfPlusDualRenderingMode.EMF_PLUS},
			{EmfPlusDualRenderingMode.EMF_PLUS_WITH_FALLBACK},
		};
	}
}
