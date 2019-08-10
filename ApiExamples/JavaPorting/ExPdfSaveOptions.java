// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.StyleIdentifier;
import org.testng.Assert;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.DmlRenderingMode;
import com.aspose.words.PdfImageCompression;
import com.aspose.words.PdfCompliance;
import com.aspose.words.ColorMode;
import com.aspose.words.SaveOptions;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.MetafileRenderingOptions;
import com.aspose.words.MetafileRenderingMode;
import com.aspose.words.IWarningCallback;
import com.aspose.words.WarningInfo;
import com.aspose.words.WarningType;
import com.aspose.ms.System.msConsole;
import com.aspose.words.WarningInfoCollection;
import com.aspose.words.HeaderFooterBookmarksExportMode;
import com.aspose.words.PdfTextCompression;
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
        //ExSummary:Shows how to create missing outline levels saving the document in PDF
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

        doc.save(getArtifactsDir() + "CreateMissingOutlineLevels.pdf", pdfSaveOptions);
        //ExEnd
        // Bind PDF with Aspose.PDF
        PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
        bookmarkEditor.BindPdf(getArtifactsDir() + "CreateMissingOutlineLevels.pdf");

        // Get all bookmarks from the document
        Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

        Assert.AreEqual(11, bookmarks.Count);
    }

    @Test
    public void allowToAddBookmarksWithWhiteSpaces() throws Exception
    {
        //ExStart
        //ExFor:OutlineOptions.BookmarksOutlineLevels
        //ExFor:BookmarksOutlineLevelCollection.Add(String, Int32)
        //ExSummary:Shows how adding bookmarks outlines with whitespaces(pdf, xps)
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add bookmarks with whitespaces. MS Word formats (like doc, docx) does not support bookmarks with whitespaces by default 
        // and all whitespaces in the bookmarks were replaced with underscores. If you need to use bookmarks in PDF or XPS outlines, you can use them with whitespaces.
        builder.startBookmark("My Bookmark");
        builder.writeln("Text inside a bookmark.");

        builder.startBookmark("Nested Bookmark");
        builder.writeln("Text inside a NestedBookmark.");
        builder.endBookmark("Nested Bookmark");

        builder.writeln("Text after Nested Bookmark.");
        builder.endBookmark("My Bookmark");

        // Specify bookmarks outline level. If you are using xps format, just use XpsSaveOptions.
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        pdfSaveOptions.getOutlineOptions().getBookmarksOutlineLevels().add("My Bookmark", 1);
        pdfSaveOptions.getOutlineOptions().getBookmarksOutlineLevels().add("Nested Bookmark", 2);

        doc.save(getArtifactsDir() + "Bookmarks.WhiteSpaces.pdf", pdfSaveOptions);
        //ExEnd
        // Bind pdf with Aspose.Pdf
        PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
        bookmarkEditor.BindPdf(getArtifactsDir() + "Bookmarks.WhiteSpaces.pdf");

        // Get all bookmarks from the document
        Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

        Assert.AreEqual(2, bookmarks.Count);

        // Assert that all the bookmarks title are with whitespaces
        Assert.AreEqual("My Bookmark", bookmarks[0].Title);
        Assert.AreEqual("Nested Bookmark", bookmarks[1].Title);
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

        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions(); { pdfSaveOptions.setDmlRenderingMode(DmlRenderingMode.DRAWING_ML); }

        doc.save(getArtifactsDir() + "DrawingMl.pdf", pdfSaveOptions);
        //ExEnd
    }

    @Test (groups = "SkipMono")
    public void withoutUpdateFields() throws Exception
    {
        //ExStart
        //ExFor:SaveOptions.UpdateFields
        //ExSummary:Shows how to update fields before saving into a PDF document.
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        {
            pdfSaveOptions.setUpdateFields(false);
        }

        doc.save(getArtifactsDir() + "UpdateFields_False.pdf", pdfSaveOptions);
        //ExEnd
        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "UpdateFields_False.pdf");

        // Get text fragment by search String
        Aspose.Pdf.Text.TextFragmentAbsorber textFragmentAbsorber = new Aspose.Pdf.Text.TextFragmentAbsorber("Page  of");
        pdfDocument.Pages.Accept(textFragmentAbsorber);

        // Assert that fields are not updated
        Assert.AreEqual("Page  of", textFragmentAbsorber.TextFragments[1].Text);
    }

    @Test (groups = "SkipMono")
    public void withUpdateFields() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions(); { pdfSaveOptions.setUpdateFields(true); }

        doc.save(getArtifactsDir() + "UpdateFields_False.pdf", pdfSaveOptions);
        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "UpdateFields_False.pdf");

        // Get text fragment by search String from PDF document
        Aspose.Pdf.Text.TextFragmentAbsorber textFragmentAbsorber = new Aspose.Pdf.Text.TextFragmentAbsorber("Page 1 of 2");
        pdfDocument.Pages.Accept(textFragmentAbsorber);

        // Assert that fields are updated
        Assert.AreEqual("Page 1 of 2", textFragmentAbsorber.TextFragments[1].Text);
    }

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
        //ExSummary:Shows how to save images to PDF using JPEG encoding to decrease file size.
        Document doc = new Document(getMyDir() + "SaveOptions.PdfImageCompression.rtf");
        
        PdfSaveOptions options = new PdfSaveOptions();
        {
            options.setImageCompression(PdfImageCompression.JPEG);
            options.setPreserveFormFields(true);
        }
        doc.save(getArtifactsDir() + "SaveOptions.PdfImageCompression.pdf", options);

        PdfSaveOptions optionsA1B = new PdfSaveOptions();
        optionsA1B.setCompliance(PdfCompliance.PDF_A_1_B);
        optionsA1B.setImageCompression(PdfImageCompression.JPEG);
        optionsA1B.setJpegQuality(100); // Use JPEG compression at 50% quality to reduce file size.

        doc.save(getArtifactsDir() + "SaveOptions.PdfImageComppression PDF_A_1_B.pdf", optionsA1B);
        //ExEnd

        PdfSaveOptions optionsA1A = new PdfSaveOptions();
        {
            optionsA1A.setCompliance(PdfCompliance.PDF_A_1_A);
            optionsA1A.setExportDocumentStructure(true);
            optionsA1A.setImageCompression(PdfImageCompression.JPEG);
        }

        doc.save(getArtifactsDir() + "SaveOptions.PdfImageComppression PDF_A_1_A.pdf", optionsA1A);
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
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions(); { pdfSaveOptions.setColorMode(ColorMode.GRAYSCALE); }

        // Assert that color image in document was grey
        doc.save(getArtifactsDir() + "ColorMode.PdfGrayscaleMode.pdf", pdfSaveOptions);
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
        
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions(); { pdfSaveOptions.setDisplayDocTitle(true); }

        doc.save(getArtifactsDir() + "PdfTitle.pdf", pdfSaveOptions);
        //ExEnd
        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfTitle.pdf");

        Assert.IsTrue(pdfDocument.DisplayDocTitle);
        Assert.AreEqual("Windows bar pdf title", pdfDocument.Info.Title);
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

        doc.save(getArtifactsDir() + "SaveOptions.MemoryOptimization.pdf", saveOptions);
        //ExEnd
    }

    @Test (dataProvider = "escapeUriDataProvider")
    public void escapeUri(String uri, String result, boolean isEscaped) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.EscapeUri
        //ExSummary: Shows how to escape hyperlinks or not in the document.
        DocumentBuilder builder = new DocumentBuilder();
        builder.insertHyperlink("Testlink", uri, false);

        // Set this property to false if you are sure that hyperlinks in document's model are already escaped
        PdfSaveOptions options = new PdfSaveOptions();
        options.setEscapeUri(isEscaped);

        builder.getDocument().save(getArtifactsDir() + "PdfSaveOptions.EscapedUri.pdf", options);
        //ExEnd

        Aspose.Pdf.Document pdfDocument =
            new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.EscapedUri.pdf");

        // Get first page
        Aspose.Pdf.Page page = pdfDocument.Pages[1];
        // Get the first link annotation
        LinkAnnotation linkAnnot = (LinkAnnotation) page.Annotations[1];

        GoToURIAction action = (GoToURIAction) linkAnnot.Action;
        String uriText = action.URI;

        msAssert.areEqual(result, uriText);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "escapeUriDataProvider")
	public static Object[][] escapeUriDataProvider() throws Exception
	{
		return new Object[][]
		{
			{"https://www.google.com/search?q= aspose",  "https://www.google.com/search?q=%20aspose",  true},
			{"https://www.google.com/search?q=%20aspose",  "https://www.google.com/search?q=%20aspose",  true},
			{"https://www.google.com/search?q= aspose",  "https://www.google.com/search?q= aspose",  false},
			{"https://www.google.com/search?q=%20aspose",  "https://www.google.com/search?q=%20aspose",  false},
		};
	}

    @Test (groups = "SkipMono")
    public void handleBinaryRasterWarnings() throws Exception
    {
        //ExStart
        //ExFor:MetafileRenderingMode
        //ExFor:MetafileRenderingOptions
        //ExFor:MetafileRenderingOptions.EmulateRasterOperations
        //ExFor:MetafileRenderingOptions.RenderingMode
        //ExFor:IWarningCallback
        //ExFor:FixedPageSaveOptions.MetafileRenderingOptions
        //ExSummary:Shows added fallback to bitmap rendering and changing type of warnings about unsupported metafile records
        Document doc = new Document(getMyDir() + "PdfSaveOptions.HandleRasterWarnings.doc");

        MetafileRenderingOptions metafileRenderingOptions =
            new MetafileRenderingOptions();
            {
                metafileRenderingOptions.setEmulateRasterOperations(false);
                metafileRenderingOptions.setRenderingMode(MetafileRenderingMode.VECTOR_WITH_FALLBACK);
            }

        // If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words renders this metafile to a bitmap. 
        HandleDocumentWarnings callback = new HandleDocumentWarnings();
        doc.setWarningCallback(callback);

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setMetafileRenderingOptions(metafileRenderingOptions);

        doc.save(getArtifactsDir() + "PdfSaveOptions.HandleRasterWarnings.pdf", saveOptions);

        msAssert.areEqual(1, callback.mWarnings.getCount());
        Assert.assertTrue(callback.mWarnings.get(0).getDescription().contains("R2_XORPEN"));
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
            //For now type of warnings about unsupported metafile records changed from DataLoss/UnexpectedContent to MinorFormattingLoss.
            if (info.getWarningType() == WarningType.MINOR_FORMATTING_LOSS)
            {
                msConsole.writeLine("Unsupported operation: " + info.getDescription());
                mWarnings.warning(info);
            }
        }

        public WarningInfoCollection mWarnings = new WarningInfoCollection();
    }
    //ExEnd

    @Test (dataProvider = "headerFooterBookmarksExportModeDataProvider")
    public void headerFooterBookmarksExportMode(/*HeaderFooterBookmarksExportMode*/int headerFooterBookmarksExportMode) throws Exception
    {
        //ExStart
        //ExFor:HeaderFooterBookmarksExportMode
        //ExFor:OutlineOptions
        //ExFor:OutlineOptions.DefaultBookmarksOutlineLevel
        //ExSummary:Shows how bookmarks in headers/footers are exported to pdf
        Document doc = new Document(getMyDir() + "PdfSaveOption.HeaderFooterBookmarksExportMode.docx");

        // You can specify how bookmarks in headers/footers are exported.
        // There is a several options for this:
        // "None" - Bookmarks in headers/footers are not exported.
        // "First" - Only bookmark in first header/footer of the section is exported.
        // "All" - Bookmarks in all headers/footers are exported.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        {
            saveOptions.setHeaderFooterBookmarksExportMode(headerFooterBookmarksExportMode);
            saveOptions.setOutlineOptions({ saveOptions.getOutlineOptions().setDefaultBookmarksOutlineLevel(1); });
        }
        doc.save(getArtifactsDir() + "PdfSaveOption.HeaderFooterBookmarksExportMode.pdf", saveOptions);
        //ExEnd
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
        Document doc = new Document(getMyDir() + "PdfSaveOptions.TestCorruptedImage.docx");

        SaveWarningCallback saveWarningCallback = new SaveWarningCallback();
        doc.setWarningCallback(saveWarningCallback);

        doc.save(getArtifactsDir() + "PdfSaveOption.HeaderFooterBookmarksExportMode.pdf", SaveFormat.PDF);

        Assert.That(saveWarningCallback.mSaveWarnings.get(0).getDescription(),
            Is.EqualTo("Image can not be processed. Possibly unsupported image format."));
    }

    public static class SaveWarningCallback implements IWarningCallback
    {
        public void warning(WarningInfo info)
        {
            if (info.getWarningType() == WarningType.MINOR_FORMATTING_LOSS)
            {
                msConsole.writeLine($"{info.WarningType}: {info.Description}.");
                mSaveWarnings.warning(info);
            }
        }

        WarningInfoCollection mSaveWarnings = new WarningInfoCollection();
	}
	
	@Test
    public void fontsScaledToMetafileSize() throws Exception
    {
        //ExStart
        //ExFor:MetafileRenderingOptions.ScaleWmfFontsToMetafileSize
        //ExSummary:Shows how to WMF fonts scaling according to metafile size on the page
        Document doc = new Document(getMyDir() + "PdfSaveOptions.FontsScaledToMetafileSize.docx");

        // There is a several options for this:
        // 'True' - Aspose.Words emulates font scaling according to metafile size on the page.
        // 'False' - Aspose.Words displays the fonts as metafile is rendered to its default size.
        // Use 'False' option is used only when metafile is rendered as vector graphics.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.getMetafileRenderingOptions().setScaleWmfFontsToMetafileSize(true);

        doc.save(getArtifactsDir() + "PdfSaveOptions.FontsScaledToMetafileSize.pdf", saveOptions);
        //ExEnd
    }

    @Test
    public void additionalTextPositioning() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.AdditionalTextPositioning
        //ExSummary:Show how to write additional text positioning operators.
        Document doc = new Document(getMyDir() + "PdfSaveOptions.AdditionalTextPositioning.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        // This may help to overcome issues with inaccurate text positioning with some printers
        saveOptions.setAdditionalTextPositioning(true);
        saveOptions.setTextCompression(PdfTextCompression.NONE);

        doc.save(getArtifactsDir() + "PdfSaveOptions.AdditionalTextPositioning.pdf", saveOptions);
        //ExEnd
    }
}
