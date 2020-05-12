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
import com.aspose.words.DocumentBuilder;
import com.aspose.words.StyleIdentifier;
import org.testng.Assert;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.PdfImageCompression;
import com.aspose.words.PdfCompliance;
import com.aspose.words.PdfImageColorSpaceExportMode;
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
import com.aspose.words.PdfTextCompression;
import com.aspose.words.Section;
import com.aspose.words.MultiplePagesType;
import com.aspose.words.PdfZoomBehavior;
import com.aspose.words.PdfPageMode;
import com.aspose.words.PdfCustomPropertiesExport;
import com.aspose.words.DmlEffectsRenderingMode;
import com.aspose.words.DmlRenderingMode;
import java.awt.image.BufferedImage;
import com.aspose.BitmapPal;
import com.aspose.words.CertificateHolder;
import com.aspose.ms.System.DateTime;
import com.aspose.words.PdfDigitalSignatureDetails;
import com.aspose.words.PdfDigitalSignatureHashAlgorithm;
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
    }

    @Test (groups = "SkipMono")
    public void withoutUpdateFields() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.Clone
        //ExFor:SaveOptions.UpdateFields
        //ExSummary:Shows how to update fields before saving into a PDF document.
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        {
            pdfSaveOptions.setUpdateFields(false);
        }

        // PdfSaveOptions objects can be cloned
        Assert.assertNotSame(pdfSaveOptions, pdfSaveOptions.deepClone());

        doc.save(getArtifactsDir() + "PdfSaveOptions.WithoutUpdateFields.pdf", pdfSaveOptions);
        //ExEnd

                        }

    @Test (groups = "SkipMono")
    public void withUpdateFields() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions(); { pdfSaveOptions.setUpdateFields(true); }

        doc.save(getArtifactsDir() + "PdfSaveOptions.WithUpdateFields.pdf", pdfSaveOptions);

                        }

    // For assert this test you need to open "SaveOptions.PdfImageCompression PDF_A_1_B Out.pdf" and "SaveOptions.PdfImageCompression PDF_A_1_A Out.pdf" 
    // and check that header image in this documents are equal header image in the "SaveOptions.PdfImageComppression Out.pdf" 
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
        Document doc = new Document(getMyDir() + "Rendering.docx");
        
        PdfSaveOptions options = new PdfSaveOptions();
        {
            options.setImageCompression(PdfImageCompression.JPEG);
            options.setPreserveFormFields(true);
        }
        doc.save(getArtifactsDir() + "PdfSaveOptions.PdfImageCompression.pdf", options);

        PdfSaveOptions optionsA1B = new PdfSaveOptions();
        {
            optionsA1B.setCompliance(PdfCompliance.PDF_A_1_B);
            optionsA1B.setImageCompression(PdfImageCompression.JPEG);
            optionsA1B.setJpegQuality(100); // Use JPEG compression at 50% quality to reduce file size
            optionsA1B.setImageColorSpaceExportMode(PdfImageColorSpaceExportMode.SIMPLE_CMYK);
        }

        doc.save(getArtifactsDir() + "PdfSaveOptions.ImageCompression.PDF_A_1_B.pdf", optionsA1B);        
        //ExEnd

        PdfSaveOptions optionsA1A = new PdfSaveOptions();
        {
            optionsA1A.setCompliance(PdfCompliance.PDF_A_1_A);
            optionsA1A.setExportDocumentStructure(true);
            optionsA1A.setImageCompression(PdfImageCompression.JPEG);
        }

        doc.save(getArtifactsDir() + "PdfSaveOptions.ImageCompression.PDF_A_1_A.pdf", optionsA1A);
    }

    @Test
    public void colorRendering() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions
        //ExFor:ColorMode
        //ExFor:FixedPageSaveOptions.ColorMode
        //ExSummary:Shows how change image color with save options property
        // Open document with color image
        Document doc = new Document(getMyDir() + "Rendering.docx");
        // Set grayscale mode for document
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions(); { pdfSaveOptions.setColorMode(ColorMode.GRAYSCALE); }
        
        // Assert that color image in document was grey
        doc.save(getArtifactsDir() + "PdfSaveOptions.ColorRendering.pdf", pdfSaveOptions);
        //ExEnd
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
        //ExSummary: Shows how to escape hyperlinks or not in the document.
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
        //ExSummary:Shows added fallback to bitmap rendering and changing type of warnings about unsupported metafile records.
        Document doc = new Document(getMyDir() + "WMF with image.docx");

        MetafileRenderingOptions metafileRenderingOptions =
            new MetafileRenderingOptions();
            {
                metafileRenderingOptions.setEmulateRasterOperations(false);
                metafileRenderingOptions.setRenderingMode(MetafileRenderingMode.VECTOR_WITH_FALLBACK);
            }

        // If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words renders this metafile to a bitmap
        HandleDocumentWarnings callback = new HandleDocumentWarnings();
        doc.setWarningCallback(callback);

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setMetafileRenderingOptions(metafileRenderingOptions);

        doc.save(getArtifactsDir() + "PdfSaveOptions.HandleBinaryRasterWarnings.pdf", saveOptions);

        Assert.assertEquals(1, callback.Warnings.getCount());
        Assert.assertTrue(callback.Warnings.get(0).getDescription().contains("R2_XORPEN"));
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
        //ExFor:PdfSaveOptions.HeaderFooterBookmarksExportMode
        //ExFor:OutlineOptions
        //ExFor:OutlineOptions.DefaultBookmarksOutlineLevel
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
        }
        doc.save(getArtifactsDir() + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf", saveOptions);
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
	
	@Test
    public void fontsScaledToMetafileSize() throws Exception
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
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        // This may help to overcome issues with inaccurate text positioning with some printers, even if the PDF looks fine,
        // but the file size will increase due to higher text positioning precision used
        saveOptions.setAdditionalTextPositioning(true);
        saveOptions.setTextCompression(PdfTextCompression.NONE);

        doc.save(getArtifactsDir() + "PdfSaveOptions.AdditionalTextPositioning.pdf", saveOptions);
        //ExEnd
    }

    @Test
    public void saveAsPdfBookFold() throws Exception
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
        options.setUseBookFoldPrintingSettings(true);

        // In order to make a booklet, we will need to print this document, stack the pages
        // in the order they come out of the printer and then fold down the middle
        doc.save(getArtifactsDir() + "PdfSaveOptions.SaveAsPdfBookFold.pdf", options);
        //ExEnd
    }

    @Test
    public void zoomBehaviour() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.PageMode
        //ExFor:PdfSaveOptions.ZoomBehavior
        //ExFor:PdfSaveOptions.ZoomFactor
        //ExFor:PdfPageMode
        //ExFor:PdfZoomBehavior
        //ExSummary:Shows how to set the default zooming of an output PDF to 1/4 of default size.
        // Open a document with multiple paragraphs
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions options = new PdfSaveOptions();
        options.setZoomBehavior(PdfZoomBehavior.ZOOM_FACTOR);
        options.setZoomFactor(25);
        options.setPageMode(PdfPageMode.USE_THUMBS);

        // When opening the .pdf with a viewer such as Adobe Acrobat Pro, the zoom level will be at 25% by default,
        // with thumbnails for each page to the left
        doc.save(getArtifactsDir() + "PdfSaveOptions.ZoomBehaviour.pdf", options);
        //ExEnd
    }

    @Test
    public void noteHyperlinks() throws Exception
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
        options.setCreateNoteHyperlinks(true);

        doc.save(getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf", options);
        //ExEnd
    }

    @Test
    public void customPropertiesExport() throws Exception
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
        options.setCustomPropertiesExport(PdfCustomPropertiesExport.STANDARD);

        doc.save(getArtifactsDir() + "PdfSaveOptions.CustomPropertiesExport.pdf", options);
        //ExEnd
    }

    @Test
    public void drawingML() throws Exception
    {
        //ExStart
        //ExFor:DmlRenderingMode
        //ExFor:DmlEffectsRenderingMode
        //ExFor:PdfSaveOptions.DmlEffectsRenderingMode
        //ExFor:SaveOptions.DmlEffectsRenderingMode
        //ExFor:SaveOptions.DmlRenderingMode
        //ExSummary:Shows how to configure DrawingML rendering quality with PdfSaveOptions.
        Document doc = new Document(getMyDir() + "DrawingML shape effects.docx");

        // Creating a new PdfSaveOptions object and setting its DmlEffectsRenderingMode to "None" will
        // strip the shapes of all their shading effects in the output pdf
        PdfSaveOptions options = new PdfSaveOptions();
        options.setDmlEffectsRenderingMode(DmlEffectsRenderingMode.NONE);
        options.setDmlRenderingMode(DmlRenderingMode.FALLBACK); 

        doc.save(getArtifactsDir() + "PdfSaveOptions.DrawingML.pdf", options);
        //ExEnd
    }

    @Test
    public void exportDocumentStructure() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.ExportDocumentStructure
        //ExSummary:Shows how to convert a .docx to .pdf while preserving the document structure.
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        // Create a PdfSaveOptions object and configure it to preserve the logical structure that's in the input document
        // The file size will be increased and the structure will be visible in the "Content" navigation pane
        // of Adobe Acrobat Pro, while editing the .pdf
        PdfSaveOptions options = new PdfSaveOptions();
        options.setExportDocumentStructure(true);

        doc.save(getArtifactsDir() + "PdfSaveOptions.ExportDocumentStructure.pdf", options);
        //ExEnd
    }

    @Test
    public void preblendImages() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.PreblendImages
        //ExSummary:Shows how to preblend images with transparent backgrounds.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        BufferedImage img = BitmapPal.loadNativeImage(getImageDir() + "Transparent background logo.png");
        builder.insertImage(img);

        // Create a PdfSaveOptions object and setting this flag may change the quality and size of the output .pdf
        // because of the way some images are rendered
        PdfSaveOptions options = new PdfSaveOptions();
        options.setPreblendImages(true);

        doc.save(getArtifactsDir() + "PdfSaveOptions.PreblendImagest.pdf", options);
        //ExEnd
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

        doc.save(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignature.pdf");
        //ExEnd
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

        doc.save(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf");
        //ExEnd
    }

    @Test
    public void renderMetafile() throws Exception
    {
        //ExStart
        //ExFor:EmfPlusDualRenderingMode
        //ExFor:MetafileRenderingOptions.EmfPlusDualRenderingMode
        //ExFor:MetafileRenderingOptions.UseEmfEmbeddedToWmf
        //ExSummary:Shows how to adjust EMF (Enhanced Windows Metafile) rendering options when saving to PDF.
        Document doc = new Document(getMyDir() + "EMF.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.getMetafileRenderingOptions().setEmfPlusDualRenderingMode(EmfPlusDualRenderingMode.EMF_PLUS);
        saveOptions.getMetafileRenderingOptions().setUseEmfEmbeddedToWmf(false);

        doc.save(getArtifactsDir() + "PdfSaveOptions.RenderMetafile.pdf", saveOptions);
        //ExEnd
    }
}
