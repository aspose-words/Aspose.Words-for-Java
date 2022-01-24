// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.words.BreakType;
import com.aspose.ms.System.IO.Stream;
import com.aspose.ms.System.IO.File;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.PageSet;
import org.testng.Assert;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.SaveFormat;
import com.aspose.words.PdfCompliance;
import com.aspose.words.PdfTextCompression;
import com.aspose.ms.System.IO.FileInfo;
import com.aspose.words.PdfImageCompression;
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
import com.aspose.words.PdfPageMode;
import com.aspose.ms.System.Globalization.msCultureInfo;
import com.aspose.words.FontSourceBase;
import com.aspose.words.FontSettings;
import com.aspose.words.FolderFontSource;
import com.aspose.words.PdfFontEmbeddingMode;
import com.aspose.words.Section;
import com.aspose.words.MultiplePagesType;
import com.aspose.ms.System.StringComparison;
import com.aspose.words.PdfZoomBehavior;
import java.util.ArrayList;
import com.aspose.words.PdfCustomPropertiesExport;
import com.aspose.words.DmlEffectsRenderingMode;
import com.aspose.words.DmlRenderingMode;
import java.awt.image.BufferedImage;
import javax.imageio.ImageIO;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.words.Dml3DEffectsRenderingMode;
import com.aspose.words.WarningSource;
import com.aspose.words.CertificateHolder;
import com.aspose.ms.System.DateTime;
import java.util.Date;
import com.aspose.words.PdfDigitalSignatureDetails;
import com.aspose.words.PdfDigitalSignatureHashAlgorithm;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.PdfDigitalSignatureTimestampSettings;
import com.aspose.ms.System.TimeSpan;
import com.aspose.words.EmfPlusDualRenderingMode;
import com.aspose.words.PdfEncryptionDetails;
import com.aspose.words.PdfEncryptionAlgorithm;
import com.aspose.words.PdfPermissions;
import com.aspose.words.NumeralFormat;
import org.testng.annotations.DataProvider;


@Test
class ExPdfSaveOptions !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void onePage() throws Exception
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.PageSet
        //ExFor:Document.Save(Stream, SaveOptions)
        //ExSummary:Shows how to convert only some of the pages in a document to PDF.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Page 1.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 2.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 3.");

        Stream stream = File.create(getArtifactsDir() + "PdfSaveOptions.OnePage.pdf");
        try /*JAVA: was using*/
        {
            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();
            
            // Set the "PageIndex" to "1" to render a portion of the document starting from the second page.
            options.setPageSet(new PageSet(1));

            // This document will contain one page starting from page two, which will only contain the second page.
            doc.save(stream, options);
        }
        finally { if (stream != null) stream.close(); }
        //ExEnd

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.OnePage.pdf");

        Assert.AreEqual(1, pdfDocument.Pages.Count);

        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
        pdfDocument.Pages.Accept(textFragmentAbsorber);

        Assert.AreEqual("Page 2.", textFragmentAbsorber.Text);
    }

    @Test
    public void headingsOutlineLevels() throws Exception
    {
        //ExStart
        //ExFor:ParagraphFormat.IsHeading
        //ExFor:PdfSaveOptions.OutlineOptions
        //ExFor:PdfSaveOptions.SaveFormat
        //ExSummary:Shows how to limit the headings' level that will appear in the outline of a saved PDF document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert headings that can serve as TOC entries of levels 1, 2, and then 3.
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);

        Assert.assertTrue(builder.getParagraphFormat().isHeading());

        builder.writeln("Heading 1");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);

        builder.writeln("Heading 1.1");
        builder.writeln("Heading 1.2");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_3);

        builder.writeln("Heading 1.2.1");
        builder.writeln("Heading 1.2.2");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setSaveFormat(SaveFormat.PDF);
        
        // The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
        // Clicking on an entry in this outline will take us to the location of its respective heading.
        // Set the "HeadingsOutlineLevels" property to "2" to exclude all headings whose levels are above 2 from the outline.
        // The last two headings we have inserted above will not appear.
        saveOptions.getOutlineOptions().setHeadingsOutlineLevels(2);

        doc.save(getArtifactsDir() + "PdfSaveOptions.HeadingsOutlineLevels.pdf", saveOptions);
        //ExEnd

                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
        bookmarkEditor.BindPdf(getArtifactsDir() + "PdfSaveOptions.HeadingsOutlineLevels.pdf");

        Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

        Assert.AreEqual(3, bookmarks.Count);
            }

    @Test (dataProvider = "createMissingOutlineLevelsDataProvider")
    public void createMissingOutlineLevels(boolean createMissingOutlineLevels) throws Exception
    {
        //ExStart
        //ExFor:OutlineOptions.CreateMissingOutlineLevels
        //ExFor:PdfSaveOptions.OutlineOptions
        //ExSummary:Shows how to work with outline levels that do not contain any corresponding headings when saving a PDF document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert headings that can serve as TOC entries of levels 1 and 5.
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);

        Assert.assertTrue(builder.getParagraphFormat().isHeading());

        builder.writeln("Heading 1");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_5);

        builder.writeln("Heading 1.1.1.1.1");
        builder.writeln("Heading 1.1.1.1.2");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
        // Clicking on an entry in this outline will take us to the location of its respective heading.
        // Set the "HeadingsOutlineLevels" property to "5" to include all headings of levels 5 and below in the outline.
        saveOptions.getOutlineOptions().setHeadingsOutlineLevels(5);

        // This document contains headings of levels 1 and 5, and no headings with levels of 2, 3, and 4. 
        // The output PDF document will treat outline levels 2, 3, and 4 as "missing".
        // Set the "CreateMissingOutlineLevels" property to "true" to include all missing levels in the outline,
        // leaving blank outline entries since there are no usable headings.
        // Set the "CreateMissingOutlineLevels" property to "false" to ignore missing outline levels,
        // and treat the outline level 5 headings as level 2.
        saveOptions.getOutlineOptions().setCreateMissingOutlineLevels(createMissingOutlineLevels);

        doc.save(getArtifactsDir() + "PdfSaveOptions.CreateMissingOutlineLevels.pdf", saveOptions);
        //ExEnd

        PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
        bookmarkEditor.BindPdf(getArtifactsDir() + "PdfSaveOptions.CreateMissingOutlineLevels.pdf");

        Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

        Assert.AreEqual(createMissingOutlineLevels ? 6 : 3, bookmarks.Count);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "createMissingOutlineLevelsDataProvider")
	public static Object[][] createMissingOutlineLevelsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "tableHeadingOutlinesDataProvider")
    public void tableHeadingOutlines(boolean createOutlinesForHeadingsInTables) throws Exception
    {
        //ExStart
        //ExFor:OutlineOptions.CreateOutlinesForHeadingsInTables
        //ExSummary:Shows how to create PDF document outline entries for headings inside tables.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a table with three rows. The first row,
        // whose text we will format in a heading-type style, will serve as the column header.
        builder.startTable();
        builder.insertCell();
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.write("Customers");
        builder.endRow();
        builder.insertCell();
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.NORMAL);
        builder.write("John Doe");
        builder.endRow();
        builder.insertCell();
        builder.write("Jane Doe");
        builder.endTable();

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
        // Clicking on an entry in this outline will take us to the location of its respective heading.
        // Set the "HeadingsOutlineLevels" property to "1" to get the outline
        // to only register headings with heading levels that are no larger than 1.
        pdfSaveOptions.getOutlineOptions().setHeadingsOutlineLevels(1);

        // Set the "CreateOutlinesForHeadingsInTables" property to "false" to exclude all headings within tables,
        // such as the one we have created above from the outline.
        // Set the "CreateOutlinesForHeadingsInTables" property to "true" to include all headings within tables
        // in the outline, provided that they have a heading level that is no larger than the value of the "HeadingsOutlineLevels" property.
        pdfSaveOptions.getOutlineOptions().setCreateOutlinesForHeadingsInTables(createOutlinesForHeadingsInTables);

        doc.save(getArtifactsDir() + "PdfSaveOptions.TableHeadingOutlines.pdf", pdfSaveOptions);
        //ExEnd

                Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.TableHeadingOutlines.pdf");

        if (createOutlinesForHeadingsInTables)
        {
            Assert.AreEqual(1, pdfDoc.Outlines.Count);
            Assert.AreEqual("Customers", pdfDoc.Outlines[1].Title);
        } else
            Assert.AreEqual(0, pdfDoc.Outlines.Count);

        TableAbsorber tableAbsorber = new TableAbsorber();
        tableAbsorber.Visit(pdfDoc.Pages[1]);

        Assert.AreEqual("Customers", tableAbsorber.TableList[0].RowList[0].CellList[0].TextFragments[1].Text);
        Assert.AreEqual("John Doe", tableAbsorber.TableList[0].RowList[1].CellList[0].TextFragments[1].Text);
        Assert.AreEqual("Jane Doe", tableAbsorber.TableList[0].RowList[2].CellList[0].TextFragments[1].Text);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "tableHeadingOutlinesDataProvider")
	public static Object[][] tableHeadingOutlinesDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void expandedOutlineLevels() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(String, SaveOptions)
        //ExFor:PdfSaveOptions
        //ExFor:OutlineOptions.HeadingsOutlineLevels
        //ExFor:OutlineOptions.ExpandedOutlineLevels
        //ExSummary:Shows how to convert a whole document to PDF with three levels in the document outline.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert headings of levels 1 to 5.
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);

        Assert.assertTrue(builder.getParagraphFormat().isHeading());

        builder.writeln("Heading 1");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);

        builder.writeln("Heading 1.1");
        builder.writeln("Heading 1.2");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_3);

        builder.writeln("Heading 1.2.1");
        builder.writeln("Heading 1.2.2");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_4);

        builder.writeln("Heading 1.2.2.1");
        builder.writeln("Heading 1.2.2.2");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_5);

        builder.writeln("Heading 1.2.2.2.1");
        builder.writeln("Heading 1.2.2.2.2");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
        // Clicking on an entry in this outline will take us to the location of its respective heading.
        // Set the "HeadingsOutlineLevels" property to "4" to exclude all headings whose levels are above 4 from the outline.
        options.getOutlineOptions().setHeadingsOutlineLevels(4);

        // If an outline entry has subsequent entries of a higher level inbetween itself and the next entry of the same or lower level,
        // an arrow will appear to the left of the entry. This entry is the "owner" of several such "sub-entries".
        // In our document, the outline entries from the 5th heading level are sub-entries of the second 4th level outline entry,
        // the 4th and 5th heading level entries are sub-entries of the second 3rd level entry, and so on. 
        // In the outline, we can click on the arrow of the "owner" entry to collapse/expand all its sub-entries.
        // Set the "ExpandedOutlineLevels" property to "2" to automatically expand all heading level 2 and lower outline entries
        // and collapse all level and 3 and higher entries when we open the document. 
        options.getOutlineOptions().setExpandedOutlineLevels(2);

        doc.save(getArtifactsDir() + "PdfSaveOptions.ExpandedOutlineLevels.pdf", options);
        //ExEnd

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.ExpandedOutlineLevels.pdf");

        Assert.AreEqual(1, pdfDocument.Outlines.Count);
        Assert.AreEqual(5, pdfDocument.Outlines.VisibleCount);

        Assert.True(pdfDocument.Outlines[1].Open);
        Assert.AreEqual(1, pdfDocument.Outlines[1].Level);

        Assert.False(pdfDocument.Outlines[1][1].Open);
        Assert.AreEqual(2, pdfDocument.Outlines[1][1].Level);

        Assert.True(pdfDocument.Outlines[1][2].Open);
        Assert.AreEqual(2, pdfDocument.Outlines[1][2].Level);
    }

    @Test (dataProvider = "updateFieldsDataProvider")
    public void updateFields(boolean updateFields) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.Clone
        //ExFor:SaveOptions.UpdateFields
        //ExSummary:Shows how to update all the fields in a document immediately before saving it to PDF.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert text with PAGE and NUMPAGES fields. These fields do not display the correct value in real time.
        // We will need to manually update them using updating methods such as "Field.Update()", and "Document.UpdateFields()"
        // each time we need them to display accurate values.
        builder.write("Page ");
        builder.insertField("PAGE", "");
        builder.write(" of ");
        builder.insertField("NUMPAGES", "");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Hello World!");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Set the "UpdateFields" property to "false" to not update all the fields in a document right before a save operation.
        // This is the preferable option if we know that all our fields will be up to date before saving.
        // Set the "UpdateFields" property to "true" to iterate through all the document
        // fields and update them before we save it as a PDF. This will make sure that all the fields will display
        // the most accurate values in the PDF.
        options.setUpdateFields(updateFields);
        
        // We can clone PdfSaveOptions objects.
        Assert.assertNotSame(options, options.deepClone());

        doc.save(getArtifactsDir() + "PdfSaveOptions.UpdateFields.pdf", options);
        //ExEnd

                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.UpdateFields.pdf");

        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
        pdfDocument.Pages.Accept(textFragmentAbsorber);

        Assert.AreEqual(updateFields ? "Page 1 of 2" : "Page  of ", textFragmentAbsorber.TextFragments[1].Text);
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

    @Test (dataProvider = "preserveFormFieldsDataProvider")
    public void preserveFormFields(boolean preserveFormFields) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.PreserveFormFields
        //ExSummary:Shows how to save a document to the PDF format using the Save method and the PdfSaveOptions class.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Please select a fruit: ");

        // Insert a combo box which will allow a user to choose an option from a collection of strings.
        builder.insertComboBox("MyComboBox", new String[] { "Apple", "Banana", "Cherry" }, 0);

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set the "PreserveFormFields" property to "true" to save form fields as interactive objects in the output PDF.
        // Set the "PreserveFormFields" property to "false" to freeze all form fields in the document at
        // their current values and display them as plain text in the output PDF.
        pdfOptions.setPreserveFormFields(preserveFormFields);

        doc.save(getArtifactsDir() + "PdfSaveOptions.PreserveFormFields.pdf", pdfOptions);
        //ExEnd

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.PreserveFormFields.pdf");

        Assert.AreEqual(1, pdfDocument.Pages.Count);

        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
        pdfDocument.Pages.Accept(textFragmentAbsorber);

        if (preserveFormFields)
        {
            Assert.AreEqual("Please select a fruit: ", textFragmentAbsorber.Text);
            TestUtil.fileContainsString("11 0 obj\r\n" +
                                        "<</Type /Annot/Subtype /Widget/P 5 0 R/FT /Ch/F 4/Rect [168.39199829 707.35101318 217.87442017 722.64007568]/Ff 131072/T(þÿ\u0000M\u0000y\0C\u0000o\u0000m\0b\u0000o\0B\u0000o\u0000x)/Opt " +
                                        "[(þÿ\0A\u0000p\u0000p\u0000l\0e) (þÿ\0B\0a\u0000n\0a\u0000n\0a) (þÿ\0C\u0000h\0e\u0000r\u0000r\u0000y) ]/V(þÿ\0A\u0000p\u0000p\u0000l\0e)/DA(0 g /FAAABD 12 Tf )/AP<</N 12 0 R>>>>",
                getArtifactsDir() + "PdfSaveOptions.PreserveFormFields.pdf");

            Aspose.Pdf.Forms.Form form = pdfDocument.Form;
            Assert.AreEqual(1, pdfDocument.Form.Count);

            ComboBoxField field = (ComboBoxField)form.Fields[0];
            
            Assert.AreEqual("MyComboBox", field.FullName);
            Assert.AreEqual(3, field.Options.Count);
            Assert.AreEqual("Apple", field.Value);
        }
        else
        {
            Assert.AreEqual("Please select a fruit: Apple", textFragmentAbsorber.Text);
            Assert.<AssertionError>Throws(() =>
            {
                TestUtil.fileContainsString("/Widget",
                    getArtifactsDir() + "PdfSaveOptions.PreserveFormFields.pdf");
            });

            Assert.AreEqual(0, pdfDocument.Form.Count);
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "preserveFormFieldsDataProvider")
	public static Object[][] preserveFormFieldsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "complianceDataProvider")
    public void compliance(/*PdfCompliance*/int pdfCompliance) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.Compliance
        //ExFor:PdfCompliance
        //ExSummary:Shows how to set the PDF standards compliance level of saved PDF documents.
        Document doc = new Document(getMyDir() + "Images.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Set the "Compliance" property to "PdfCompliance.PdfA1b" to comply with the "PDF/A-1b" standard,
        // which aims to preserve the visual appearance of the document as Aspose.Words convert it to PDF.
        // Set the "Compliance" property to "PdfCompliance.Pdf17" to comply with the "1.7" standard.
        // Set the "Compliance" property to "PdfCompliance.PdfA1a" to comply with the "PDF/A-1a" standard,
        // which complies with "PDF/A-1b" as well as preserving the document structure of the original document.
        // This helps with making documents searchable but may significantly increase the size of already large documents.
        saveOptions.setCompliance(pdfCompliance);

        doc.save(getArtifactsDir() + "PdfSaveOptions.Compliance.pdf", saveOptions);
        //ExEnd

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.Compliance.pdf");

        switch (pdfCompliance)
        {
            case PdfCompliance.PDF_17:
                Assert.AreEqual(PdfFormat.v_1_7, pdfDocument.PdfFormat);
                Assert.AreEqual("1.7", pdfDocument.Version);
                break;
            case PdfCompliance.PDF_A_2_A:
                Assert.AreEqual(PdfFormat.PDF_A_2A, pdfDocument.PdfFormat);
                Assert.AreEqual("1.7", pdfDocument.Version);
                break;
            case PdfCompliance.PDF_A_2_U:
                Assert.AreEqual(PdfFormat.PDF_A_2U, pdfDocument.PdfFormat);
                Assert.AreEqual("1.7", pdfDocument.Version);
                break;
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "complianceDataProvider")
	public static Object[][] complianceDataProvider() throws Exception
	{
		return new Object[][]
		{
			{PdfCompliance.PDF_A_2_U},
			{PdfCompliance.PDF_17},
			{PdfCompliance.PDF_A_2_A},
		};
	}

    @Test (dataProvider = "textCompressionDataProvider")
    public void textCompression(/*PdfTextCompression*/int pdfTextCompression) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions
        //ExFor:PdfSaveOptions.TextCompression
        //ExFor:PdfTextCompression
        //ExSummary:Shows how to apply text compression when saving a document to PDF.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < 100; i++)
            builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                            "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Set the "TextCompression" property to "PdfTextCompression.None" to not apply any
        // compression to text when we save the document to PDF.
        // Set the "TextCompression" property to "PdfTextCompression.Flate" to apply ZIP compression
        // to text when we save the document to PDF. The larger the document, the bigger the impact that this will have.
        options.setTextCompression(pdfTextCompression);

        doc.save(getArtifactsDir() + "PdfSaveOptions.TextCompression.pdf", options);
        //ExEnd

        switch (pdfTextCompression)
        {
            case PdfTextCompression.NONE:
                Assert.That(60000,
                    Is.LessThan(new FileInfo(getArtifactsDir() + "PdfSaveOptions.TextCompression.pdf").getLength()));
                TestUtil.fileContainsString("12 0 obj\r\n<</Length 13 0 R>>stream",
                    getArtifactsDir() + "PdfSaveOptions.TextCompression.pdf");
                break;
            case PdfTextCompression.FLATE:
                Assert.That(30000,
                    Is.AtLeast(new FileInfo(getArtifactsDir() + "PdfSaveOptions.TextCompression.pdf").getLength()));
                TestUtil.fileContainsString("12 0 obj\r\n<</Length 13 0 R/Filter /FlateDecode>>stream",
                    getArtifactsDir() + "PdfSaveOptions.TextCompression.pdf");
                break;
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "textCompressionDataProvider")
	public static Object[][] textCompressionDataProvider() throws Exception
	{
		return new Object[][]
		{
			{PdfTextCompression.NONE},
			{PdfTextCompression.FLATE},
		};
	}

    @Test (dataProvider = "imageCompressionDataProvider")
    public void imageCompression(/*PdfImageCompression*/int pdfImageCompression) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.ImageCompression
        //ExFor:PdfSaveOptions.JpegQuality
        //ExFor:PdfImageCompression
        //ExSummary:Shows how to specify a compression type for all images in a document that we are converting to PDF.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Jpeg image:");
        builder.insertImage(getImageDir() + "Logo.jpg");
        builder.insertParagraph();
        builder.writeln("Png image:");
        builder.insertImage(getImageDir() + "Transparent background logo.png");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Set the "ImageCompression" property to "PdfImageCompression.Auto" to use the
        // "ImageCompression" property to control the quality of the Jpeg images that end up in the output PDF.
        // Set the "ImageCompression" property to "PdfImageCompression.Jpeg" to use the
        // "ImageCompression" property to control the quality of all images that end up in the output PDF.
        pdfSaveOptions.setImageCompression(pdfImageCompression);

        // Set the "JpegQuality" property to "10" to strengthen compression at the cost of image quality.
        pdfSaveOptions.setJpegQuality(10);

        doc.save(getArtifactsDir() + "PdfSaveOptions.ImageCompression.pdf", pdfSaveOptions);
        //ExEnd

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.ImageCompression.pdf");
        Stream pdfDocImageStream = pdfDocument.Pages[1].Resources.Images[1].ToStream();

        try /*JAVA: was using*/
        {
            TestUtil.verifyImage(400, 400, pdfDocImageStream);
        }
        finally { if (pdfDocImageStream != null) pdfDocImageStream.close(); }

        pdfDocImageStream = pdfDocument.Pages[1].Resources.Images[2].ToStream();

        try /*JAVA: was using*/
        {
            switch (pdfImageCompression)
            {
                case PdfImageCompression.AUTO:
                    Assert.That(50000, 
                        Is.LessThan(new FileInfo(getArtifactsDir() + "PdfSaveOptions.ImageCompression.pdf").getLength()));
                    break;
                case PdfImageCompression.JPEG:
                    Assert.That(42000, 
                        Is.AtLeast(new FileInfo(getArtifactsDir() + "PdfSaveOptions.ImageCompression.pdf").getLength()));
                    TestUtil.verifyImage(400, 400, pdfDocImageStream);
                    break;
            }
        }
        finally { if (pdfDocImageStream != null) pdfDocImageStream.close(); }
            }

	//JAVA-added data provider for test method
	@DataProvider(name = "imageCompressionDataProvider")
	public static Object[][] imageCompressionDataProvider() throws Exception
	{
		return new Object[][]
		{
			{PdfImageCompression.AUTO},
			{PdfImageCompression.JPEG},
		};
	}

    @Test (dataProvider = "imageColorSpaceExportModeDataProvider")
    public void imageColorSpaceExportMode(/*PdfImageColorSpaceExportMode*/int pdfImageColorSpaceExportMode) throws Exception
    {
        //ExStart
        //ExFor:PdfImageColorSpaceExportMode
        //ExFor:PdfSaveOptions.ImageColorSpaceExportMode
        //ExSummary:Shows how to set a different color space for images in a document as we export it to PDF.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Jpeg image:");
        builder.insertImage(getImageDir() + "Logo.jpg");
        builder.insertParagraph();
        builder.writeln("Png image:");
        builder.insertImage(getImageDir() + "Transparent background logo.png");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Set the "ImageColorSpaceExportMode" property to "PdfImageColorSpaceExportMode.Auto" to get Aspose.Words to
        // automatically select the color space for images in the document that it converts to PDF.
        // In most cases, the color space will be RGB.
        // Set the "ImageColorSpaceExportMode" property to "PdfImageColorSpaceExportMode.SimpleCmyk"
        // to use the CMYK color space for all images in the saved PDF.
        // Aspose.Words will also apply Flate compression to all images and ignore the "ImageCompression" property's value.
        pdfSaveOptions.setImageColorSpaceExportMode(pdfImageColorSpaceExportMode);

        doc.save(getArtifactsDir() + "PdfSaveOptions.ImageColorSpaceExportMode.pdf", pdfSaveOptions);
        //ExEnd

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.ImageColorSpaceExportMode.pdf");
        XImage pdfDocImage = pdfDocument.Pages[1].Resources.Images[1];

        switch (pdfImageColorSpaceExportMode)
        {
            case PdfImageColorSpaceExportMode.AUTO:
                Assert.That(20000, Is.LessThan(pdfDocImage.ToStream().Length));
                break;
            case PdfImageColorSpaceExportMode.SIMPLE_CMYK:
                Assert.That(100000, Is.LessThan(pdfDocImage.ToStream().Length));
                break;
        }

        Assert.AreEqual(400, pdfDocImage.Width);
        Assert.AreEqual(400, pdfDocImage.Height);
        Assert.AreEqual(ColorType.Rgb, pdfDocImage.GetColorType());

        pdfDocImage = pdfDocument.Pages[1].Resources.Images[2];

        switch (pdfImageColorSpaceExportMode)
        {
            case PdfImageColorSpaceExportMode.AUTO:
                Assert.That(25000, Is.AtLeast(pdfDocImage.ToStream().Length));
                break;
            case PdfImageColorSpaceExportMode.SIMPLE_CMYK:
                Assert.That(18000, Is.LessThan(pdfDocImage.ToStream().Length));
                break;
        }

        Assert.AreEqual(400, pdfDocImage.Width);
        Assert.AreEqual(400, pdfDocImage.Height);
        Assert.AreEqual(ColorType.Rgb, pdfDocImage.GetColorType());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "imageColorSpaceExportModeDataProvider")
	public static Object[][] imageColorSpaceExportModeDataProvider() throws Exception
	{
		return new Object[][]
		{
			{PdfImageColorSpaceExportMode.AUTO},
			{PdfImageColorSpaceExportMode.SIMPLE_CMYK},
		};
	}

    @Test
    public void downsampleOptions() throws Exception
    {
        //ExStart
        //ExFor:DownsampleOptions
        //ExFor:DownsampleOptions.DownsampleImages
        //ExFor:DownsampleOptions.Resolution
        //ExFor:DownsampleOptions.ResolutionThreshold
        //ExFor:PdfSaveOptions.DownsampleOptions
        //ExSummary:Shows how to change the resolution of images in the PDF document.
        Document doc = new Document(getMyDir() + "Images.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // By default, Aspose.Words downsample all images in a document that we save to PDF to 220 ppi.
        Assert.assertTrue(options.getDownsampleOptions().getDownsampleImages());
        Assert.assertEquals(220, options.getDownsampleOptions().getResolution());
        Assert.assertEquals(0, options.getDownsampleOptions().getResolutionThreshold());

        doc.save(getArtifactsDir() + "PdfSaveOptions.DownsampleOptions.Default.pdf", options);

        // Set the "Resolution" property to "36" to downsample all images to 36 ppi.
        options.getDownsampleOptions().setResolution(36);

        // Set the "ResolutionThreshold" property to only apply the downsampling to
        // images with a resolution that is above 128 ppi.
        options.getDownsampleOptions().setResolutionThreshold(128);

        // Only the first two images from the document will be downsampled at this stage.
        doc.save(getArtifactsDir() + "PdfSaveOptions.DownsampleOptions.LowerResolution.pdf", options);
        //ExEnd

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.DownsampleOptions.Default.pdf");
        XImage pdfDocImage = pdfDocument.Pages[1].Resources.Images[1];

        Assert.That(300000, Is.LessThan(pdfDocImage.ToStream().Length));
        Assert.AreEqual(ColorType.Rgb, pdfDocImage.GetColorType());
    }

    @Test (dataProvider = "colorRenderingDataProvider")
    public void colorRendering(/*ColorMode*/int colorMode) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions
        //ExFor:ColorMode
        //ExFor:FixedPageSaveOptions.ColorMode
        //ExSummary:Shows how to change image color with saving options property.
        Document doc = new Document(getMyDir() + "Images.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        // Set the "ColorMode" property to "Grayscale" to render all images from the document in black and white.
        // The size of the output document may be larger with this setting.
        // Set the "ColorMode" property to "Normal" to render all images in color.
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions(); { pdfSaveOptions.setColorMode(colorMode); }
        
        doc.save(getArtifactsDir() + "PdfSaveOptions.ColorRendering.pdf", pdfSaveOptions);
        //ExEnd

                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.ColorRendering.pdf");
        XImage pdfDocImage = pdfDocument.Pages[1].Resources.Images[1];

        switch (colorMode)
        {
            case ColorMode.NORMAL:
                Assert.That(300000, Is.LessThan(pdfDocImage.ToStream().Length));
                Assert.AreEqual(ColorType.Rgb, pdfDocImage.GetColorType());
                break;
            case ColorMode.GRAYSCALE:
                Assert.That(1000000, Is.LessThan(pdfDocImage.ToStream().Length));
                Assert.AreEqual(ColorType.Grayscale, pdfDocImage.GetColorType());
                break;
        }
            }

	//JAVA-added data provider for test method
	@DataProvider(name = "colorRenderingDataProvider")
	public static Object[][] colorRenderingDataProvider() throws Exception
	{
		return new Object[][]
		{
			{ColorMode.GRAYSCALE},
			{ColorMode.NORMAL},
		};
	}

    @Test (dataProvider = "docTitleDataProvider")
    public void docTitle(boolean displayDocTitle) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.DisplayDocTitle
        //ExSummary:Shows how to display the title of the document as the title bar.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        doc.getBuiltInDocumentProperties().setTitle("Windows bar pdf title");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        // Set the "DisplayDocTitle" to "true" to get some PDF readers, such as Adobe Acrobat Pro,
        // to display the value of the document's "Title" built-in property in the tab that belongs to this document.
        // Set the "DisplayDocTitle" to "false" to get such readers to display the document's filename.
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions(); { pdfSaveOptions.setDisplayDocTitle(displayDocTitle); }

        doc.save(getArtifactsDir() + "PdfSaveOptions.DocTitle.pdf", pdfSaveOptions);
        //ExEnd

                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.DocTitle.pdf");

        Assert.AreEqual(displayDocTitle, pdfDocument.DisplayDocTitle);
        Assert.AreEqual("Windows bar pdf title", pdfDocument.Info.Title);
            }

	//JAVA-added data provider for test method
	@DataProvider(name = "docTitleDataProvider")
	public static Object[][] docTitleDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "memoryOptimizationDataProvider")
    public void memoryOptimization(boolean memoryOptimization) throws Exception
    {
        //ExStart
        //ExFor:SaveOptions.CreateSaveOptions(SaveFormat)
        //ExFor:SaveOptions.MemoryOptimization
        //ExSummary:Shows an option to optimize memory consumption when rendering large documents to PDF.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        SaveOptions saveOptions = SaveOptions.createSaveOptions(SaveFormat.PDF);

        // Set the "MemoryOptimization" property to "true" to lower the memory footprint of large documents' saving operations
        // at the cost of increasing the duration of the operation.
        // Set the "MemoryOptimization" property to "false" to save the document as a PDF normally.
        saveOptions.setMemoryOptimization(memoryOptimization);

        doc.save(getArtifactsDir() + "PdfSaveOptions.MemoryOptimization.pdf", saveOptions);
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "memoryOptimizationDataProvider")
	public static Object[][] memoryOptimizationDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "escapeUriDataProvider")
    public void escapeUri(String uri, String result) throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertHyperlink("Testlink", uri, false);

        doc.save(getArtifactsDir() + "PdfSaveOptions.EscapedUri.pdf");

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.EscapedUri.pdf");

        Page page = pdfDocument.Pages[1];
        LinkAnnotation linkAnnot = (LinkAnnotation)page.Annotations[1];

        GoToURIAction action = (GoToURIAction)linkAnnot.Action;

        Assert.AreEqual(result, action.URI);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "escapeUriDataProvider")
	public static Object[][] escapeUriDataProvider() throws Exception
	{
		return new Object[][]
		{
			{"https://www.google.com/search?q= aspose",  "https://www.google.com/search?q=%20aspose"},
			{"https://www.google.com/search?q=%20aspose",  "https://www.google.com/search?q=%20aspose"},
		};
	}

    @Test (dataProvider = "openHyperlinksInNewWindowDataProvider")
    public void openHyperlinksInNewWindow(boolean openHyperlinksInNewWindow) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.OpenHyperlinksInNewWindow
        //ExSummary:Shows how to save hyperlinks in a document we convert to PDF so that they open new pages when we click on them.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertHyperlink("Testlink", "https://www.google.com/search?q=%20aspose", false);

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Set the "OpenHyperlinksInNewWindow" property to "true" to save all hyperlinks using Javascript code
        // that forces readers to open these links in new windows/browser tabs.
        // Set the "OpenHyperlinksInNewWindow" property to "false" to save all hyperlinks normally.
        options.setOpenHyperlinksInNewWindow(openHyperlinksInNewWindow);

        doc.save(getArtifactsDir() + "PdfSaveOptions.OpenHyperlinksInNewWindow.pdf", options);
        //ExEnd

        if (openHyperlinksInNewWindow)
            TestUtil.fileContainsString(
                "<</Type /Annot/Subtype /Link/Rect [70.84999847 707.35101318 110.17799377 721.15002441]/BS " +
                "<</Type/Border/S/S/W 0>>/A<</Type /Action/S /JavaScript/JS(app.launchURL\\(\"https://www.google.com/search?q=%20aspose\", true\\);)>>>>",
                getArtifactsDir() + "PdfSaveOptions.OpenHyperlinksInNewWindow.pdf");
        else
            TestUtil.fileContainsString(
                "<</Type /Annot/Subtype /Link/Rect [70.84999847 707.35101318 110.17799377 721.15002441]/BS " +
                "<</Type/Border/S/S/W 0>>/A<</Type /Action/S /URI/URI(https://www.google.com/search?q=%20aspose)>>>>",
                getArtifactsDir() + "PdfSaveOptions.OpenHyperlinksInNewWindow.pdf");

        Aspose.Pdf.Document pdfDocument =
            new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.OpenHyperlinksInNewWindow.pdf");

        Page page = pdfDocument.Pages[1];
        LinkAnnotation linkAnnot = (LinkAnnotation) page.Annotations[1];

        Assert.AreEqual(openHyperlinksInNewWindow ? JavascriptAction.class : GoToURIAction.class,
            linkAnnot.Action.GetType());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "openHyperlinksInNewWindowDataProvider")
	public static Object[][] openHyperlinksInNewWindowDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    //ExStart
    //ExFor:MetafileRenderingMode
    //ExFor:MetafileRenderingOptions
    //ExFor:MetafileRenderingOptions.EmulateRasterOperations
    //ExFor:MetafileRenderingOptions.RenderingMode
    //ExFor:IWarningCallback
    //ExFor:FixedPageSaveOptions.MetafileRenderingOptions
    //ExSummary:Shows added a fallback to bitmap rendering and changing type of warnings about unsupported metafile records.
    @Test (groups = "SkipMono") //ExSkip
    public void handleBinaryRasterWarnings() throws Exception
    {
        Document doc = new Document(getMyDir() + "WMF with image.docx");

        MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions();

        // Set the "EmulateRasterOperations" property to "false" to fall back to bitmap when
        // it encounters a metafile, which will require raster operations to render in the output PDF.
        metafileRenderingOptions.setEmulateRasterOperations(false);

        // Set the "RenderingMode" property to "VectorWithFallback" to try to render every metafile using vector graphics.
        metafileRenderingOptions.setRenderingMode(MetafileRenderingMode.VECTOR_WITH_FALLBACK);

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF and applies the configuration
        // in our MetafileRenderingOptions object to the saving operation.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setMetafileRenderingOptions(metafileRenderingOptions);

        HandleDocumentWarnings callback = new HandleDocumentWarnings();
        doc.setWarningCallback(callback);

        doc.save(getArtifactsDir() + "PdfSaveOptions.HandleBinaryRasterWarnings.pdf", saveOptions);

        Assert.assertEquals(1, callback.Warnings.getCount());
        Assert.assertEquals("'R2_XORPEN' binary raster operation is partly supported.",
            callback.Warnings.get(0).getDescription());
    }

    /// <summary>
    /// Prints and collects formatting loss-related warnings that occur upon saving a document.
    /// </summary>
    public static class HandleDocumentWarnings implements IWarningCallback
    {
        public void warning(WarningInfo info)
        {
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
        //ExSummary:Shows to process bookmarks in headers/footers in a document that we are rendering to PDF.
        Document doc = new Document(getMyDir() + "Bookmarks in headers and footers.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Set the "PageMode" property to "PdfPageMode.UseOutlines" to display the outline navigation pane in the output PDF.
        saveOptions.setPageMode(PdfPageMode.USE_OUTLINES);

        // Set the "DefaultBookmarksOutlineLevel" property to "1" to display all
        // bookmarks at the first level of the outline in the output PDF.
        saveOptions.getOutlineOptions().setDefaultBookmarksOutlineLevel(1);

        // Set the "HeaderFooterBookmarksExportMode" property to "HeaderFooterBookmarksExportMode.None" to
        // not export any bookmarks that are inside headers/footers.
        // Set the "HeaderFooterBookmarksExportMode" property to "HeaderFooterBookmarksExportMode.First" to
        // only export bookmarks in the first section's header/footers.
        // Set the "HeaderFooterBookmarksExportMode" property to "HeaderFooterBookmarksExportMode.All" to
        // export bookmarks that are in all headers/footers.
        saveOptions.setHeaderFooterBookmarksExportMode(headerFooterBookmarksExportMode);

        doc.save(getArtifactsDir() + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf", saveOptions);
        //ExEnd

        Aspose.Pdf.Document pdfDoc =
            new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf");
        String inputDocLocaleName = new msCultureInfo(doc.getStyles().getDefaultFont().getLocaleId()).getName();

        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
        pdfDoc.Pages.Accept(textFragmentAbsorber);
        switch (headerFooterBookmarksExportMode)
        {
            case com.aspose.words.HeaderFooterBookmarksExportMode.NONE:
                TestUtil.fileContainsString($"<</Type /Catalog/Pages 3 0 R/Lang({inputDocLocaleName})/Metadata 4 0 R>>\r\n",
                    getArtifactsDir() + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf");

                Assert.AreEqual(0, pdfDoc.Outlines.Count);
                break;
            case com.aspose.words.HeaderFooterBookmarksExportMode.FIRST:
            case com.aspose.words.HeaderFooterBookmarksExportMode.ALL:
                TestUtil.fileContainsString(
                    $"<</Type /Catalog/Pages 3 0 R/Outlines 14 0 R/PageMode /UseOutlines/Lang({inputDocLocaleName})/Metadata 4 0 R>>",
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
    public void fontsScaledToMetafileSize(boolean scaleWmfFonts) throws Exception
    {
        //ExStart
        //ExFor:MetafileRenderingOptions.ScaleWmfFontsToMetafileSize
        //ExSummary:Shows how to WMF fonts scaling according to metafile size on the page.
        Document doc = new Document(getMyDir() + "WMF with text.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Set the "ScaleWmfFontsToMetafileSize" property to "true" to scale fonts
        // that format text within WMF images according to the size of the metafile on the page.
        // Set the "ScaleWmfFontsToMetafileSize" property to "false" to
        // preserve the default scale of these fonts.
        saveOptions.getMetafileRenderingOptions().setScaleWmfFontsToMetafileSize(scaleWmfFonts);

        doc.save(getArtifactsDir() + "PdfSaveOptions.FontsScaledToMetafileSize.pdf", saveOptions);
        //ExEnd

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.FontsScaledToMetafileSize.pdf");
        TextFragmentAbsorber textAbsorber = new TextFragmentAbsorber();

        pdfDocument.Pages[1].Accept(textAbsorber);
        Rectangle textFragmentRectangle = textAbsorber.TextFragments[3].Rectangle;

        Assert.AreEqual(scaleWmfFonts ? 1.589d : 5.045d, textFragmentRectangle.Width, 0.001d);
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

    @Test (dataProvider = "embedFullFontsDataProvider")
    public void embedFullFonts(boolean embedFullFonts) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.#ctor
        //ExFor:PdfSaveOptions.EmbedFullFonts
        //ExSummary:Shows how to enable or disable subsetting when embedding fonts while rendering a document to PDF.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Arial");
        builder.writeln("Hello world!");
        builder.getFont().setName("Arvo");
        builder.writeln("The quick brown fox jumps over the lazy dog.");

        // Configure our font sources to ensure that we have access to both the fonts in this document.
        FontSourceBase[] originalFontsSources = FontSettings.getDefaultInstance().getFontsSources();
        FolderFontSource folderFontSource = new FolderFontSource(getFontsDir(), true);
        FontSettings.getDefaultInstance().setFontsSources(new FontSourceBase[] { originalFontsSources[0], folderFontSource });

        FontSourceBase[] fontSources = FontSettings.getDefaultInstance().getFontsSources();
        Assert.True(fontSources[0].getAvailableFonts().Any(f => f.FullFontName == "Arial"));
        Assert.True(fontSources[1].getAvailableFonts().Any(f => f.FullFontName == "Arvo"));

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Since our document contains a custom font, embedding in the output document may be desirable.
        // Set the "EmbedFullFonts" property to "true" to embed every glyph of every embedded font in the output PDF.
        // The document's size may become very large, but we will have full use of all fonts if we edit the PDF.
        // Set the "EmbedFullFonts" property to "false" to apply subsetting to fonts, saving only the glyphs
        // that the document is using. The file will be considerably smaller,
        // but we may need access to any custom fonts if we edit the document.
        options.setEmbedFullFonts(embedFullFonts);

        doc.save(getArtifactsDir() + "PdfSaveOptions.EmbedFullFonts.pdf", options);

        if (embedFullFonts)
            Assert.That(500000, Is.LessThan(new FileInfo(getArtifactsDir() + "PdfSaveOptions.EmbedFullFonts.pdf").getLength()));
        else
            Assert.That(25000, Is.AtLeast(new FileInfo(getArtifactsDir() + "PdfSaveOptions.EmbedFullFonts.pdf").getLength()));

        // Restore the original font sources.
        FontSettings.getDefaultInstance().setFontsSources(originalFontsSources);
        //ExEnd

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.EmbedFullFonts.pdf");

        Aspose.Pdf.Text.Font[] pdfDocFonts = pdfDocument.FontUtilities.GetAllFonts();

        Assert.AreEqual("ArialMT", pdfDocFonts[0].FontName);
        Assert.AreNotEqual(embedFullFonts, pdfDocFonts[0].IsSubset);

        Assert.AreEqual("Arvo", pdfDocFonts[1].FontName);
        Assert.AreNotEqual(embedFullFonts, pdfDocFonts[1].IsSubset);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "embedFullFontsDataProvider")
	public static Object[][] embedFullFontsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "embedWindowsFontsDataProvider")
    public void embedWindowsFonts(/*PdfFontEmbeddingMode*/int pdfFontEmbeddingMode) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.FontEmbeddingMode
        //ExFor:PdfFontEmbeddingMode
        //ExSummary:Shows how to set Aspose.Words to skip embedding Arial and Times New Roman fonts into a PDF document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // "Arial" is a standard font, and "Courier New" is a nonstandard font.
        builder.getFont().setName("Arial");
        builder.writeln("Hello world!");
        builder.getFont().setName("Courier New");
        builder.writeln("The quick brown fox jumps over the lazy dog.");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Set the "EmbedFullFonts" property to "true" to embed every glyph of every embedded font in the output PDF.
        options.setEmbedFullFonts(true);

        // Set the "FontEmbeddingMode" property to "EmbedAll" to embed all fonts in the output PDF.
        // Set the "FontEmbeddingMode" property to "EmbedNonstandard" to only allow nonstandard fonts' embedding in the output PDF.
        // Set the "FontEmbeddingMode" property to "EmbedNone" to not embed any fonts in the output PDF.
        options.setFontEmbeddingMode(pdfFontEmbeddingMode);

        doc.save(getArtifactsDir() + "PdfSaveOptions.EmbedWindowsFonts.pdf", options);

        switch (pdfFontEmbeddingMode)
        {
            case PdfFontEmbeddingMode.EMBED_ALL:
                Assert.That(1000000, Is.LessThan(new FileInfo(getArtifactsDir() + "PdfSaveOptions.EmbedWindowsFonts.pdf").getLength()));
                break;
            case PdfFontEmbeddingMode.EMBED_NONSTANDARD:
                Assert.That(480000, Is.LessThan(new FileInfo(getArtifactsDir() + "PdfSaveOptions.EmbedWindowsFonts.pdf").getLength()));
                break;
            case PdfFontEmbeddingMode.EMBED_NONE:
                Assert.That(4217, Is.AtLeast(new FileInfo(getArtifactsDir() + "PdfSaveOptions.EmbedWindowsFonts.pdf").getLength()));
                break;
        }
        //ExEnd

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.EmbedWindowsFonts.pdf");

        Aspose.Pdf.Text.Font[] pdfDocFonts = pdfDocument.FontUtilities.GetAllFonts();

        Assert.AreEqual("ArialMT", pdfDocFonts[0].FontName);
        Assert.AreEqual(pdfFontEmbeddingMode == PdfFontEmbeddingMode.EMBED_ALL, 
            pdfDocFonts[0].IsEmbedded);

        Assert.AreEqual("CourierNewPSMT", pdfDocFonts[1].FontName);
        Assert.AreEqual(pdfFontEmbeddingMode == PdfFontEmbeddingMode.EMBED_ALL || pdfFontEmbeddingMode == PdfFontEmbeddingMode.EMBED_NONSTANDARD, 
            pdfDocFonts[1].IsEmbedded);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "embedWindowsFontsDataProvider")
	public static Object[][] embedWindowsFontsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{PdfFontEmbeddingMode.EMBED_ALL},
			{PdfFontEmbeddingMode.EMBED_NONE},
			{PdfFontEmbeddingMode.EMBED_NONSTANDARD},
		};
	}

    @Test (dataProvider = "embedCoreFontsDataProvider")
    public void embedCoreFonts(boolean useCoreFonts) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.UseCoreFonts
        //ExSummary:Shows how enable/disable PDF Type 1 font substitution.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Arial");
        builder.writeln("Hello world!");
        builder.getFont().setName("Courier New");
        builder.writeln("The quick brown fox jumps over the lazy dog.");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Set the "UseCoreFonts" property to "true" to replace some fonts,
        // including the two fonts in our document, with their PDF Type 1 equivalents.
        // Set the "UseCoreFonts" property to "false" to not apply PDF Type 1 fonts.
        options.setUseCoreFonts(useCoreFonts);

        doc.save(getArtifactsDir() + "PdfSaveOptions.EmbedCoreFonts.pdf", options);

        if (useCoreFonts)
            Assert.That(3000, Is.AtLeast(new FileInfo(getArtifactsDir() + "PdfSaveOptions.EmbedCoreFonts.pdf").getLength()));
        else
            Assert.That(30000, Is.LessThan(new FileInfo(getArtifactsDir() + "PdfSaveOptions.EmbedCoreFonts.pdf").getLength()));
        //ExEnd

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.EmbedCoreFonts.pdf");

        Aspose.Pdf.Text.Font[] pdfDocFonts = pdfDocument.FontUtilities.GetAllFonts();

        if (useCoreFonts)
        {
            Assert.AreEqual("Helvetica", pdfDocFonts[0].FontName);
            Assert.AreEqual("Courier", pdfDocFonts[1].FontName);
        }
        else
        {
            Assert.AreEqual("ArialMT", pdfDocFonts[0].FontName);
            Assert.AreEqual("CourierNewPSMT", pdfDocFonts[1].FontName);
        }

        Assert.AreNotEqual(useCoreFonts, pdfDocFonts[0].IsEmbedded);
        Assert.AreNotEqual(useCoreFonts, pdfDocFonts[1].IsEmbedded);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "embedCoreFontsDataProvider")
	public static Object[][] embedCoreFontsDataProvider() throws Exception
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
        Document doc = new Document(getMyDir() + "Text positioning operators.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        {
            saveOptions.setTextCompression(PdfTextCompression.NONE);

            // Set the "AdditionalTextPositioning" property to "true" to attempt to fix incorrect
            // element positioning in the output PDF, should there be any, at the cost of increased file size.
            // Set the "AdditionalTextPositioning" property to "false" to render the document as usual.
            saveOptions.setAdditionalTextPositioning(applyAdditionalTextPositioning);
        }

        doc.save(getArtifactsDir() + "PdfSaveOptions.AdditionalTextPositioning.pdf", saveOptions);
        //ExEnd

        Aspose.Pdf.Document pdfDocument =
            new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.AdditionalTextPositioning.pdf");
        TextFragmentAbsorber textAbsorber = new TextFragmentAbsorber();

        pdfDocument.Pages[1].Accept(textAbsorber);

        SetGlyphsPositionShowText tjOperator =
            (SetGlyphsPositionShowText) textAbsorber.TextFragments[1].Page.Contents[85];

        if (applyAdditionalTextPositioning)
        {
            Assert.That(100000,
                Is.LessThan(new FileInfo(getArtifactsDir() + "PdfSaveOptions.AdditionalTextPositioning.pdf").getLength()));
            Assert.AreEqual(
                "[0 (S) 0 (a) 0 (m) 0 (s) 0 (t) 0 (a) -1 (g) 1 (,) 0 ( ) 0 (1) 0 (0) 0 (.) 0 ( ) 0 (N) 0 (o) 0 (v) 0 (e) 0 (m) 0 (b) 0 (e) 0 (r) -1 ( ) 1 (2) -1 (0) 0 (1) 0 (8)] TJ",
                tjOperator.ToString());
        }
        else
        {
            Assert.That(97000,
                Is.LessThan(new FileInfo(getArtifactsDir() + "PdfSaveOptions.AdditionalTextPositioning.pdf").getLength()));
            Assert.AreEqual("[(Samsta) -1 (g) 1 (, 10. November) -1 ( ) 1 (2) -1 (018)] TJ", tjOperator.ToString());
        }
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
    public void saveAsPdfBookFold(boolean renderTextAsBookfold) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.UseBookFoldPrintingSettings
        //ExSummary:Shows how to save a document to the PDF format in the form of a book fold.
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Set the "UseBookFoldPrintingSettings" property to "true" to arrange the contents
        // in the output PDF in a way that helps us use it to make a booklet.
        // Set the "UseBookFoldPrintingSettings" property to "false" to render the PDF normally.
        options.setUseBookFoldPrintingSettings(renderTextAsBookfold);

        // If we are rendering the document as a booklet, we must set the "MultiplePages"
        // properties of the page setup objects of all sections to "MultiplePagesType.BookFoldPrinting".
        if (renderTextAsBookfold)
            for (Section s : (Iterable<Section>) doc.getSections())
            {
                s.getPageSetup().setMultiplePages(MultiplePagesType.BOOK_FOLD_PRINTING);
            }

        // Once we print this document on both sides of the pages, we can fold all the pages down the middle at once,
        // and the contents will line up in a way that creates a booklet.
        doc.save(getArtifactsDir() + "PdfSaveOptions.SaveAsPdfBookFold.pdf", options);
        //ExEnd

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.SaveAsPdfBookFold.pdf");
        TextAbsorber textAbsorber = new TextAbsorber();

        pdfDocument.Pages.Accept(textAbsorber);

        if (renderTextAsBookfold)
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
        //ExSummary:Shows how to set the default zooming that a reader applies when opening a rendered PDF document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        // Set the "ZoomBehavior" property to "PdfZoomBehavior.ZoomFactor" to get a PDF reader to
        // apply a percentage-based zoom factor when we open the document with it.
        // Set the "ZoomFactor" property to "25" to give the zoom factor a value of 25%.
        PdfSaveOptions options = new PdfSaveOptions();
        {
            options.setZoomBehavior(PdfZoomBehavior.ZOOM_FACTOR);
            options.setZoomFactor(25);
        }

        // When we open this document using a reader such as Adobe Acrobat, we will see the document scaled at 1/4 of its actual size.
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
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Set the "PageMode" property to "PdfPageMode.FullScreen" to get the PDF reader to open the saved
        // document in full-screen mode, which takes over the monitor's display and has no controls visible.
        // Set the "PageMode" property to "PdfPageMode.UseThumbs" to get the PDF reader to display a separate panel
        // with a thumbnail for each page in the document.
        // Set the "PageMode" property to "PdfPageMode.UseOC" to get the PDF reader to display a separate panel
        // that allows us to work with any layers present in the document.
        // Set the "PageMode" property to "PdfPageMode.UseOutlines" to get the PDF reader
        // also to display the outline, if possible.
        // Set the "PageMode" property to "PdfPageMode.UseNone" to get the PDF reader to display just the document itself.
        options.setPageMode(pageMode);

        doc.save(getArtifactsDir() + "PdfSaveOptions.PageMode.pdf", options);
        //ExEnd
        
        String docLocaleName = new msCultureInfo(doc.getStyles().getDefaultFont().getLocaleId()).getName();

        switch (pageMode)
        {
            case PdfPageMode.FULL_SCREEN:
                TestUtil.fileContainsString(
                    $"<</Type /Catalog/Pages 3 0 R/PageMode /FullScreen/Lang({docLocaleName})/Metadata 4 0 R>>\r\n",
                    getArtifactsDir() + "PdfSaveOptions.PageMode.pdf");
                break;
            case PdfPageMode.USE_THUMBS:
                TestUtil.fileContainsString(
                    $"<</Type /Catalog/Pages 3 0 R/PageMode /UseThumbs/Lang({docLocaleName})/Metadata 4 0 R>>",
                    getArtifactsDir() + "PdfSaveOptions.PageMode.pdf");
                break;
            case PdfPageMode.USE_OC:
                TestUtil.fileContainsString(
                    $"<</Type /Catalog/Pages 3 0 R/PageMode /UseOC/Lang({docLocaleName})/Metadata 4 0 R>>\r\n",
                    getArtifactsDir() + "PdfSaveOptions.PageMode.pdf");
                break;
            case PdfPageMode.USE_OUTLINES:
            case PdfPageMode.USE_NONE:
                TestUtil.fileContainsString($"<</Type /Catalog/Pages 3 0 R/Lang({docLocaleName})/Metadata 4 0 R>>\r\n",
                    getArtifactsDir() + "PdfSaveOptions.PageMode.pdf");
                break;
        }

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.PageMode.pdf");

        switch (pageMode)
        {
            case PdfPageMode.USE_NONE:
            case PdfPageMode.USE_OUTLINES:
                Assert.AreEqual(Aspose.Pdf.PageMode.UseNone, pdfDocument.PageMode);
                break;
            case PdfPageMode.USE_THUMBS:
                Assert.AreEqual(Aspose.Pdf.PageMode.UseThumbs, pdfDocument.PageMode);
                break;
            case PdfPageMode.FULL_SCREEN:
                Assert.AreEqual(Aspose.Pdf.PageMode.FullScreen, pdfDocument.PageMode);
                break;
            case PdfPageMode.USE_OC:
                Assert.AreEqual(Aspose.Pdf.PageMode.UseOC, pdfDocument.PageMode);
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
    public void noteHyperlinks(boolean createNoteHyperlinks) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.CreateNoteHyperlinks
        //ExSummary:Shows how to make footnotes and endnotes function as hyperlinks.
        Document doc = new Document(getMyDir() + "Footnotes and endnotes.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Set the "CreateNoteHyperlinks" property to "true" to turn all footnote/endnote symbols
        // in the text act as links that, upon clicking, take us to their respective footnotes/endnotes.
        // Set the "CreateNoteHyperlinks" property to "false" not to have footnote/endnote symbols link to anything.
        options.setCreateNoteHyperlinks(createNoteHyperlinks);

        doc.save(getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf", options);
        //ExEnd

        if (createNoteHyperlinks)
        {
            TestUtil.fileContainsString(
                "<</Type /Annot/Subtype /Link/Rect [157.80099487 720.90106201 159.35600281 733.55004883]/BS <</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 85 677 0]>>",
                getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString(
                "<</Type /Annot/Subtype /Link/Rect [202.16900635 720.90106201 206.06201172 733.55004883]/BS <</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 85 79 0]>>",
                getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString(
                "<</Type /Annot/Subtype /Link/Rect [212.23199463 699.2510376 215.34199524 711.90002441]/BS <</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 85 654 0]>>",
                getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString(
                "<</Type /Annot/Subtype /Link/Rect [258.15499878 699.2510376 262.04800415 711.90002441]/BS <</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 85 68 0]>>",
                getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString(
                "<</Type /Annot/Subtype /Link/Rect [85.05000305 68.19904327 88.66500092 79.69804382]/BS <</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 202 733 0]>>",
                getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString(
                "<</Type /Annot/Subtype /Link/Rect [85.05000305 56.70004272 88.66500092 68.19904327]/BS <</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 258 711 0]>>",
                getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString(
                "<</Type /Annot/Subtype /Link/Rect [85.05000305 666.10205078 86.4940033 677.60107422]/BS <</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 157 733 0]>>",
                getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString(
                "<</Type /Annot/Subtype /Link/Rect [85.05000305 643.10406494 87.93800354 654.60308838]/BS <</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 212 711 0]>>",
                getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
        }
        else
        {
            if (!isRunningOnMono())
                Assert.<AssertionError>Throws(() =>
                    TestUtil.fileContainsString("<</Type /Annot/Subtype /Link/Rect",
                        getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf"));
        }

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
        Page page = pdfDocument.Pages[1];
        AnnotationSelector annotationSelector = new AnnotationSelector(new LinkAnnotation(page, Rectangle.Trivial));

        page.Accept(annotationSelector);

        ArrayList</* unknown Type use JavaGenericArguments */> linkAnnotations = annotationSelector.Selected.<LinkAnnotation>Cast().ToList();

        if (createNoteHyperlinks)
        {
            Assert.AreEqual(8, linkAnnotations.Count(a => a.AnnotationType == AnnotationType.Link));

            Assert.AreEqual("1 XYZ 85 677 0", linkAnnotations.get(0).Destination.ToString());
            Assert.AreEqual("1 XYZ 85 79 0", linkAnnotations.get(1).Destination.ToString());
            Assert.AreEqual("1 XYZ 85 654 0", linkAnnotations.get(2).Destination.ToString());
            Assert.AreEqual("1 XYZ 85 68 0", linkAnnotations.get(3).Destination.ToString());
            Assert.AreEqual("1 XYZ 202 733 0", linkAnnotations.get(4).Destination.ToString());
            Assert.AreEqual("1 XYZ 258 711 0", linkAnnotations.get(5).Destination.ToString());
            Assert.AreEqual("1 XYZ 157 733 0", linkAnnotations.get(6).Destination.ToString());
            Assert.AreEqual("1 XYZ 212 711 0", linkAnnotations.get(7).Destination.ToString());
        }
        else
        {
            Assert.AreEqual(0, annotationSelector.Selected.Count);
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
        //ExSummary:Shows how to export custom properties while converting a document to PDF.
        Document doc = new Document();

        doc.getCustomDocumentProperties().add("Company", "My value");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Set the "CustomPropertiesExport" property to "PdfCustomPropertiesExport.None" to discard
        // custom document properties as we save the document to .PDF. 
        // Set the "CustomPropertiesExport" property to "PdfCustomPropertiesExport.Standard"
        // to preserve custom properties within the output PDF document.
        // Set the "CustomPropertiesExport" property to "PdfCustomPropertiesExport.Metadata"
        // to preserve custom properties in an XMP packet.
        options.setCustomPropertiesExport(pdfCustomPropertiesExportMode);

        doc.save(getArtifactsDir() + "PdfSaveOptions.CustomPropertiesExport.pdf", options);
        //ExEnd

        switch (pdfCustomPropertiesExportMode)
        {
            case PdfCustomPropertiesExport.NONE:
                if (!isRunningOnMono())
                {
                    Assert.<AssertionError>Throws(() => TestUtil.fileContainsString(
                        doc.getCustomDocumentProperties().get(0).getName(),
                        getArtifactsDir() + "PdfSaveOptions.CustomPropertiesExport.pdf"));
                    Assert.<AssertionError>Throws(() => TestUtil.fileContainsString(
                        "<</Type /Metadata/Subtype /XML/Length 8 0 R/Filter /FlateDecode>>",
                        getArtifactsDir() + "PdfSaveOptions.CustomPropertiesExport.pdf"));
                }

                break;
            case PdfCustomPropertiesExport.STANDARD:
                TestUtil.fileContainsString(
                    "<</Creator(þÿ\0A\u0000s\u0000p\u0000o\u0000s\0e\u0000.\u0000W\u0000o\u0000r\0d\u0000s)/Producer(þÿ\0A\u0000s\u0000p\u0000o\u0000s\0e\u0000.\u0000W\u0000o\u0000r\0d\u0000s\u0000 \0f\u0000o\u0000r\u0000",
                    getArtifactsDir() + "PdfSaveOptions.CustomPropertiesExport.pdf");
                TestUtil.fileContainsString("/Company (þÿ\u0000M\u0000y\u0000 \u0000v\0a\u0000l\u0000u\0e)>>",
                    getArtifactsDir() + "PdfSaveOptions.CustomPropertiesExport.pdf");
                break;
            case PdfCustomPropertiesExport.METADATA:
                TestUtil.fileContainsString("<</Type /Metadata/Subtype /XML/Length 8 0 R/Filter /FlateDecode>>",
                    getArtifactsDir() + "PdfSaveOptions.CustomPropertiesExport.pdf");
                break;
        }

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.CustomPropertiesExport.pdf");

        Assert.AreEqual("Aspose.Words", pdfDocument.Info.Creator);
        Assert.True(pdfDocument.Info.Producer.StartsWith("Aspose.Words"));
        
        switch (pdfCustomPropertiesExportMode)
        {
            case PdfCustomPropertiesExport.NONE:
                Assert.AreEqual(2, pdfDocument.Info.Count);
                Assert.AreEqual(3, pdfDocument.Metadata.Count);
                break;
            case PdfCustomPropertiesExport.METADATA:
                Assert.AreEqual(2, pdfDocument.Info.Count);
                Assert.AreEqual(4, pdfDocument.Metadata.Count);

                Assert.AreEqual("Aspose.Words", pdfDocument.Metadata["xmp:CreatorTool"].ToString());
                Assert.AreEqual("Company", pdfDocument.Metadata["custprops:Property1"].ToString());
                break;
            case PdfCustomPropertiesExport.STANDARD:
                Assert.AreEqual(3, pdfDocument.Info.Count);
                Assert.AreEqual(3, pdfDocument.Metadata.Count);

                Assert.AreEqual("My value", pdfDocument.Info["Company"]);
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
        //ExSummary:Shows how to configure the rendering quality of DrawingML effects in a document as we save it to PDF.
        Document doc = new Document(getMyDir() + "DrawingML shape effects.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Set the "DmlEffectsRenderingMode" property to "DmlEffectsRenderingMode.None" to discard all DrawingML effects.
        // Set the "DmlEffectsRenderingMode" property to "DmlEffectsRenderingMode.Simplified"
        // to render a simplified version of DrawingML effects.
        // Set the "DmlEffectsRenderingMode" property to "DmlEffectsRenderingMode.Fine" to
        // render DrawingML effects with more accuracy and also with more processing cost.
        options.setDmlEffectsRenderingMode(effectsRenderingMode);

        Assert.assertEquals(DmlRenderingMode.DRAWING_ML, options.getDmlRenderingMode());

        doc.save(getArtifactsDir() + "PdfSaveOptions.DrawingMLEffects.pdf", options);
        //ExEnd

        Aspose.Pdf.Document pdfDocument =
            new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.DrawingMLEffects.pdf");

        ImagePlacementAbsorber imagePlacementAbsorber = new ImagePlacementAbsorber();
        imagePlacementAbsorber.Visit(pdfDocument.Pages[1]);

        TableAbsorber tableAbsorber = new TableAbsorber();
        tableAbsorber.Visit(pdfDocument.Pages[1]);

        switch (effectsRenderingMode)
        {
            case DmlEffectsRenderingMode.NONE:
            case DmlEffectsRenderingMode.SIMPLIFIED:
                TestUtil.fileContainsString("5 0 obj\r\n" +
                                            "<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                    getArtifactsDir() + "PdfSaveOptions.DrawingMLEffects.pdf");
                Assert.AreEqual(0, imagePlacementAbsorber.ImagePlacements.Count);
                Assert.AreEqual(28, tableAbsorber.TableList.Count);
                break;
            case DmlEffectsRenderingMode.FINE:
                TestUtil.fileContainsString(
                    "5 0 obj\r\n<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R>>/XObject<</X1 10 0 R/X2 11 0 R/X3 12 0 R/X4 13 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                    getArtifactsDir() + "PdfSaveOptions.DrawingMLEffects.pdf");
                Assert.AreEqual(21, imagePlacementAbsorber.ImagePlacements.Count);
                Assert.AreEqual(4, tableAbsorber.TableList.Count);
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
        //ExSummary:Shows how to render fallback shapes when saving to PDF.
        Document doc = new Document(getMyDir() + "DrawingML shape fallbacks.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Set the "DmlRenderingMode" property to "DmlRenderingMode.Fallback"
        // to substitute DML shapes with their fallback shapes.
        // Set the "DmlRenderingMode" property to "DmlRenderingMode.DrawingML"
        // to render the DML shapes themselves.
        options.setDmlRenderingMode(dmlRenderingMode);

        doc.save(getArtifactsDir() + "PdfSaveOptions.DrawingMLFallback.pdf", options);
        //ExEnd

        switch (dmlRenderingMode)
        {
            case DmlRenderingMode.DRAWING_ML:
                TestUtil.fileContainsString(
                    "<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R/FAAABB 11 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                    getArtifactsDir() + "PdfSaveOptions.DrawingMLFallback.pdf");
                break;
            case DmlRenderingMode.FALLBACK:
                TestUtil.fileContainsString(
                    "5 0 obj\r\n<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R/FAAABD 13 0 R>>/ExtGState<</GS1 10 0 R/GS2 11 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                    getArtifactsDir() + "PdfSaveOptions.DrawingMLFallback.pdf");
                break;
        }

        Aspose.Pdf.Document pdfDocument =
            new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.DrawingMLFallback.pdf");

        ImagePlacementAbsorber imagePlacementAbsorber = new ImagePlacementAbsorber();
        imagePlacementAbsorber.Visit(pdfDocument.Pages[1]);

        TableAbsorber tableAbsorber = new TableAbsorber();
        tableAbsorber.Visit(pdfDocument.Pages[1]);

        switch (dmlRenderingMode)
        {
            case DmlRenderingMode.DRAWING_ML:
                Assert.AreEqual(6, tableAbsorber.TableList.Count);
                break;
            case DmlRenderingMode.FALLBACK:
                Assert.AreEqual(15, tableAbsorber.TableList.Count);
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
    public void exportDocumentStructure(boolean exportDocumentStructure) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.ExportDocumentStructure
        //ExSummary:Shows how to preserve document structure elements, which can assist in programmatically interpreting our document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 1"));
        builder.writeln("Hello world!");
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Normal"));
        builder.write(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Set the "ExportDocumentStructure" property to "true" to make the document structure, such tags, available via the
        // "Content" navigation pane of Adobe Acrobat at the cost of increased file size.
        // Set the "ExportDocumentStructure" property to "false" to not export the document structure.
        options.setExportDocumentStructure(exportDocumentStructure);

        // Suppose we export document structure while saving this document. In that case,
        // we can open it using Adobe Acrobat and find tags for elements such as the heading
        // and the next paragraph via "View" -> "Show/Hide" -> "Navigation panes" -> "Tags".
        doc.save(getArtifactsDir() + "PdfSaveOptions.ExportDocumentStructure.pdf", options);
        //ExEnd

        if (exportDocumentStructure)
        {
            TestUtil.fileContainsString("5 0 obj\r\n" +
                                        "<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R/FAAABC 12 0 R>>/ExtGState<</GS1 10 0 R/GS2 14 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>/StructParents 0/Tabs /S>>",
                getArtifactsDir() + "PdfSaveOptions.ExportDocumentStructure.pdf");
        }
        else
        {
            TestUtil.fileContainsString("5 0 obj\r\n" +
                                        "<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R/FAAABB 11 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
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
    public void preblendImages(boolean preblendImages) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.PreblendImages
        //ExSummary:Shows how to preblend images with transparent backgrounds while saving a document to PDF.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        BufferedImage img = ImageIO.read(getImageDir() + "Transparent background logo.png");
        builder.insertImage(img);

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Set the "PreblendImages" property to "true" to preblend transparent images
        // with a background, which may reduce artifacts.
        // Set the "PreblendImages" property to "false" to render transparent images normally.
        options.setPreblendImages(preblendImages);

        doc.save(getArtifactsDir() + "PdfSaveOptions.PreblendImages.pdf", options);
        //ExEnd

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.PreblendImages.pdf");
        XImage image = pdfDocument.Pages[1].Resources.Images[1];

        MemoryStream stream = new MemoryStream();
        try /*JAVA: was using*/
        {
            image.Save(stream);

            if (preblendImages)
            {
                TestUtil.fileContainsString("11 0 obj\r\n20849 ", getArtifactsDir() + "PdfSaveOptions.PreblendImages.pdf");
                Assert.assertEquals(17898, stream.getLength());
            }
            else
            {
                TestUtil.fileContainsString("11 0 obj\r\n19289 ", getArtifactsDir() + "PdfSaveOptions.PreblendImages.pdf");
                Assert.assertEquals(19216, stream.getLength());
            }
        }
        finally { if (stream != null) stream.close(); }
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

    @Test (dataProvider = "interpolateImagesDataProvider")
    public void interpolateImages(boolean interpolateImages) throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.InterpolateImages
        //ExSummary:Shows how to perform interpolation on images while saving a document to PDF.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        BufferedImage img = ImageIO.read(getImageDir() + "Transparent background logo.png");
        builder.insertImage(img);

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Set the "InterpolateImages" property to "true" to get the reader that opens this document to interpolate images.
        // Their resolution should be lower than that of the device that is displaying the document.
        // Set the "InterpolateImages" property to "false" to make it so that the reader does not apply any interpolation.
        saveOptions.setInterpolateImages(interpolateImages);

        // When we open this document with a reader such as Adobe Acrobat, we will need to zoom in on the image
        // to see the interpolation effect if we saved the document with it enabled.
        doc.save(getArtifactsDir() + "PdfSaveOptions.InterpolateImages.pdf", saveOptions);
        //ExEnd

        if (interpolateImages)
        {
            TestUtil.fileContainsString("7 0 obj\r\n" +
                                        "<</Type /XObject/Subtype /Image/Width 400/Height 400/ColorSpace /DeviceRGB/BitsPerComponent 8/SMask 10 0 R/Interpolate true/Length 11 0 R/Filter /FlateDecode>>",
                getArtifactsDir() + "PdfSaveOptions.InterpolateImages.pdf");
        }
        else
        {
            TestUtil.fileContainsString("7 0 obj\r\n" +
                                        "<</Type /XObject/Subtype /Image/Width 400/Height 400/ColorSpace /DeviceRGB/BitsPerComponent 8/SMask 10 0 R/Length 11 0 R/Filter /FlateDecode>>",
                getArtifactsDir() + "PdfSaveOptions.InterpolateImages.pdf");
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "interpolateImagesDataProvider")
	public static Object[][] interpolateImagesDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
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

        Assert.AreEqual(38, warningCallback.Count);
    }

    public static class RenderCallback implements IWarningCallback
    {
        public void warning(WarningInfo info)
        {
            System.out.println("{info.WarningType}: {info.Description}.");
            mWarnings.Add(info);
        }

         !!Autoporter error: Indexer ApiExamples.ExPdfSaveOptions.RenderCallback.Item(int) hasn't both getter and setter!
            mWarnings.Clear();
        }

        public int Count => private mWarnings.CountmWarnings;

        /// <summary>
        /// Returns true if a warning with the specified properties has been generated.
        /// </summary>
        public boolean contains(/*WarningSource*/int source, /*WarningType*/int type, String description)
        {
            return mWarnings.Any(warning =>
                warning.Source == source && warning.WarningType == type && warning.Description == description);
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
        //ExSummary:Shows how to sign a generated PDF document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Contents of signed PDF.");

        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Configure the "DigitalSignatureDetails" object of the "SaveOptions" object to
        // digitally sign the document as we render it with the "Save" method.
        DateTime signingTime = new Date();
        options.setDigitalSignatureDetails(new PdfDigitalSignatureDetails(certificateHolder, "Test Signing", "My Office", signingTime));
        options.getDigitalSignatureDetails().setHashAlgorithm(PdfDigitalSignatureHashAlgorithm.SHA_256);

        Assert.assertEquals("Test Signing", options.getDigitalSignatureDetails().getReason());
        Assert.assertEquals("My Office", options.getDigitalSignatureDetails().getLocation());
        Assert.assertEquals(signingTime.toUniversalTime(), options.getDigitalSignatureDetails().getSignatureDateInternal().toUniversalTime());

        doc.save(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignature.pdf", options);
        //ExEnd

        TestUtil.fileContainsString("7 0 obj\r\n" +
                                    "<</Type /Annot/Subtype /Widget/Rect [0 0 0 0]/FT /Sig/T",
            getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignature.pdf");

        Assert.assertFalse(FileFormatUtil.detectFileFormat(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignature.pdf")
            .hasDigitalSignature());

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignature.pdf");

        Assert.True(pdfDocument.Form.SignaturesExist);

        SignatureField signatureField = (SignatureField)pdfDocument.Form[1];

        Assert.AreEqual("AsposeDigitalSignature", signatureField.FullName);
        Assert.AreEqual("AsposeDigitalSignature", signatureField.PartialName);
        Assert.AreEqual(Aspose.Pdf.Forms.PKCS7.class, signatureField.Signature.GetType());
        Assert.AreEqual(DateTime.getToday(), signatureField.Signature.Date.Date);
        Assert.AreEqual("þÿ\u0000M\u0000o\u0000r\u0000z\0a\u0000l\u0000.\u0000M\0e", signatureField.Signature.Authority);
        Assert.AreEqual("þÿ\u0000M\u0000y\u0000 \u0000O\0f\0f\u0000i\0c\0e", signatureField.Signature.Location);
        Assert.AreEqual("þÿ\u0000T\0e\u0000s\u0000t\u0000 \u0000S\u0000i\u0000g\u0000n\u0000i\u0000n\u0000g", signatureField.Signature.Reason);
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
        //ExSummary:Shows how to sign a saved PDF document digitally and timestamp it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Signed PDF contents.");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Create a digital signature and assign it to our SaveOptions object to sign the document when we save it to PDF. 
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");
        options.setDigitalSignatureDetails(new PdfDigitalSignatureDetails(certificateHolder, "Test Signing", "Aspose Office", new Date()));

        // Create a timestamp authority-verified timestamp.
        options.getDigitalSignatureDetails().setTimestampSettings(new PdfDigitalSignatureTimestampSettings("https://freetsa.org/tsr", "JohnDoe", "MyPassword"));

        // The default lifespan of the timestamp is 100 seconds.
        Assert.assertEquals(100.0d, options.getDigitalSignatureDetails().getTimestampSettings().getTimeoutInternal().getTotalSeconds());

        // We can set our timeout period via the constructor.
        options.getDigitalSignatureDetails().setTimestampSettings(new PdfDigitalSignatureTimestampSettings("https://freetsa.org/tsr", "JohnDoe", "MyPassword", TimeSpan.fromMinutes(30.0)));

        Assert.assertEquals(1800.0d, options.getDigitalSignatureDetails().getTimestampSettings().getTimeoutInternal().getTotalSeconds());
        Assert.assertEquals("https://freetsa.org/tsr", options.getDigitalSignatureDetails().getTimestampSettings().getServerUrl());
        Assert.assertEquals("JohnDoe", options.getDigitalSignatureDetails().getTimestampSettings().getUserName());
        Assert.assertEquals("MyPassword", options.getDigitalSignatureDetails().getTimestampSettings().getPassword());

        // The "Save" method will apply our signature to the output document at this time.
        doc.save(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf", options);
        //ExEnd

        Assert.assertFalse(FileFormatUtil.detectFileFormat(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf").hasDigitalSignature());
        TestUtil.fileContainsString("7 0 obj\r\n" +
                                    "<</Type /Annot/Subtype /Widget/Rect [0 0 0 0]/FT /Sig/T", 
        getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf");

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf");

        Assert.True(pdfDocument.Form.SignaturesExist);

        SignatureField signatureField = (SignatureField)pdfDocument.Form[1];

        Assert.AreEqual("AsposeDigitalSignature", signatureField.FullName);
        Assert.AreEqual("AsposeDigitalSignature", signatureField.PartialName);
        Assert.AreEqual(Aspose.Pdf.Forms.PKCS7.class, signatureField.Signature.GetType());
        Assert.AreEqual(new DateTime(1, 1, 1, 0, 0, 0), signatureField.Signature.Date);
        Assert.AreEqual("þÿ\u0000M\u0000o\u0000r\u0000z\0a\u0000l\u0000.\u0000M\0e", signatureField.Signature.Authority);
        Assert.AreEqual("þÿ\0A\u0000s\u0000p\u0000o\u0000s\0e\u0000 \u0000O\0f\0f\u0000i\0c\0e", signatureField.Signature.Location);
        Assert.AreEqual("þÿ\u0000T\0e\u0000s\u0000t\u0000 \u0000S\u0000i\u0000g\u0000n\u0000i\u0000n\u0000g", signatureField.Signature.Reason);
        Assert.Null(signatureField.Signature.TimestampSettings);
    }

    @Test (dataProvider = "renderMetafileDataProvider")
    public void renderMetafile(/*EmfPlusDualRenderingMode*/int renderingMode) throws Exception
    {
        //ExStart
        //ExFor:EmfPlusDualRenderingMode
        //ExFor:MetafileRenderingOptions.EmfPlusDualRenderingMode
        //ExFor:MetafileRenderingOptions.UseEmfEmbeddedToWmf
        //ExSummary:Shows how to configure Enhanced Windows Metafile-related rendering options when saving to PDF.
        Document doc = new Document(getMyDir() + "EMF.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Set the "EmfPlusDualRenderingMode" property to "EmfPlusDualRenderingMode.Emf"
        // to only render the EMF part of an EMF+ dual metafile.
        // Set the "EmfPlusDualRenderingMode" property to "EmfPlusDualRenderingMode.EmfPlus" to
        // to render the EMF+ part of an EMF+ dual metafile.
        // Set the "EmfPlusDualRenderingMode" property to "EmfPlusDualRenderingMode.EmfPlusWithFallback"
        // to render the EMF+ part of an EMF+ dual metafile if all of the EMF+ records are supported.
        // Otherwise, Aspose.Words will render the EMF part.
        saveOptions.getMetafileRenderingOptions().setEmfPlusDualRenderingMode(renderingMode);

        // Set the "UseEmfEmbeddedToWmf" property to "true" to render embedded EMF data
        // for metafiles that we can render as vector graphics.
        saveOptions.getMetafileRenderingOptions().setUseEmfEmbeddedToWmf(true);

        doc.save(getArtifactsDir() + "PdfSaveOptions.RenderMetafile.pdf", saveOptions);
        //ExEnd

        Aspose.Pdf.Document pdfDocument =
            new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.RenderMetafile.pdf");

        switch (renderingMode)
        {
            case EmfPlusDualRenderingMode.EMF:
            case EmfPlusDualRenderingMode.EMF_PLUS_WITH_FALLBACK:
                Assert.AreEqual(0, pdfDocument.Pages[1].Resources.Images.Count);
                TestUtil.fileContainsString("5 0 obj\r\n" +
                                            "<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 595.29998779 841.90002441]/Resources<</Font<</FAAAAI 8 0 R/FAAABB 11 0 R/FAAABE 14 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                    getArtifactsDir() + "PdfSaveOptions.RenderMetafile.pdf");
                break;
            case EmfPlusDualRenderingMode.EMF_PLUS:
                Assert.AreEqual(1, pdfDocument.Pages[1].Resources.Images.Count);
                TestUtil.fileContainsString("5 0 obj\r\n" +
                                            "<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 595.29998779 841.90002441]/Resources<</Font<</FAAAAI 8 0 R/FAAABC 12 0 R/FAAABF 15 0 R>>/XObject<</X1 10 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
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

    @Test
    public void encryptionPermissions() throws Exception
    {
        //ExStart
        //ExFor:PdfEncryptionDetails.#ctor
        //ExFor:PdfSaveOptions.EncryptionDetails
        //ExFor:PdfEncryptionDetails.Permissions
        //ExFor:PdfEncryptionDetails.EncryptionAlgorithm
        //ExFor:PdfEncryptionDetails.OwnerPassword
        //ExFor:PdfEncryptionDetails.UserPassword
        //ExFor:PdfEncryptionAlgorithm
        //ExFor:PdfPermissions
        //ExFor:PdfEncryptionDetails
        //ExSummary:Shows how to set permissions on a saved PDF document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello world!");

        PdfEncryptionDetails encryptionDetails =
            new PdfEncryptionDetails("password", "", PdfEncryptionAlgorithm.RC_4_128);

        // Start by disallowing all permissions.
        encryptionDetails.setPermissions(PdfPermissions.DISALLOW_ALL);

        // Extend permissions to allow the editing of annotations.
        encryptionDetails.setPermissions(PdfPermissions.MODIFY_ANNOTATIONS | PdfPermissions.DOCUMENT_ASSEMBLY);

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Enable encryption via the "EncryptionDetails" property.
        saveOptions.setEncryptionDetails(encryptionDetails);

        // When we open this document, we will need to provide the password before accessing its contents.
        doc.save(getArtifactsDir() + "PdfSaveOptions.EncryptionPermissions.pdf", saveOptions);
        //ExEnd

        Aspose.Pdf.Document pdfDocument;

        Assert.<InvalidPasswordException>Throws(() => 
            pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.EncryptionPermissions.pdf"));

        pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.EncryptionPermissions.pdf", "password");
        TextFragmentAbsorber textAbsorber = new TextFragmentAbsorber();

        pdfDocument.Pages[1].Accept(textAbsorber);
        
        Assert.AreEqual("Hello world!", textAbsorber.Text);
    }

    @Test (dataProvider = "setNumeralFormatDataProvider")
    public void setNumeralFormat(/*NumeralFormat*/int numeralFormat) throws Exception
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.NumeralFormat
        //ExFor:NumeralFormat
        //ExSummary:Shows how to set the numeral format used when saving to PDF.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setLocaleId(new msCultureInfo("ar-AR").getLCID());
        builder.writeln("1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 50, 100");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Set the "NumeralFormat" property to "NumeralFormat.ArabicIndic" to
        // use glyphs from the U+0660 to U+0669 range as numbers.
        // Set the "NumeralFormat" property to "NumeralFormat.Context" to
        // look up the locale to determine what number of glyphs to use.
        // Set the "NumeralFormat" property to "NumeralFormat.EasternArabicIndic" to
        // use glyphs from the U+06F0 to U+06F9 range as numbers.
        // Set the "NumeralFormat" property to "NumeralFormat.European" to use european numerals.
        // Set the "NumeralFormat" property to "NumeralFormat.System" to determine the symbol set from regional settings.
        options.setNumeralFormat(numeralFormat);

        doc.save(getArtifactsDir() + "PdfSaveOptions.SetNumeralFormat.pdf", options);
        //ExEnd

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.SetNumeralFormat.pdf");
        TextFragmentAbsorber textAbsorber = new TextFragmentAbsorber();

        pdfDocument.Pages[1].Accept(textAbsorber);

        switch (numeralFormat)
        {
            case NumeralFormat.EUROPEAN:
                Assert.AreEqual("1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 50, 100", textAbsorber.Text);
                break;
            case NumeralFormat.ARABIC_INDIC:
                Assert.AreEqual(", ٢, ٣, ٤, ٥, ٦, ٧, ٨, ٩, ١٠, ٥٠, ١١٠٠", textAbsorber.Text);
                break;
            case NumeralFormat.EASTERN_ARABIC_INDIC:
                Assert.AreEqual("۱۰۰ ,۵۰ ,۱۰ ,۹ ,۸ ,۷ ,۶ ,۵ ,۴ ,۳ ,۲ ,۱", textAbsorber.Text);
                break;
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "setNumeralFormatDataProvider")
	public static Object[][] setNumeralFormatDataProvider() throws Exception
	{
		return new Object[][]
		{
			{NumeralFormat.ARABIC_INDIC},
			{NumeralFormat.CONTEXT},
			{NumeralFormat.EASTERN_ARABIC_INDIC},
			{NumeralFormat.EUROPEAN},
			{NumeralFormat.SYSTEM},
		};
	}

    @Test
    public void exportPageSet() throws Exception
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.PageSet
        //ExSummary:Shows how to export Odd pages from the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < 5; i++)
        {
            builder.writeln($"Page {i + 1} ({(i % 2 == 0 ? "odd" : "even")})");
            if (i < 4)
                builder.insertBreak(BreakType.PAGE_BREAK);
        }

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Below are three PageSet properties that we can use to filter out a set of pages from
        // our document to save in an output PDF document based on the parity of their page numbers.
        // 1 -  Save only the even-numbered pages:
        options.setPageSet(PageSet.getEven());

        doc.save(getArtifactsDir() + "PdfSaveOptions.ExportPageSet.Even.pdf", options);

        // 2 -  Save only the odd-numbered pages:
        options.setPageSet(PageSet.getOdd());

        doc.save(getArtifactsDir() + "PdfSaveOptions.ExportPageSet.Odd.pdf", options);

        // 3 -  Save every page:
        options.setPageSet(PageSet.getAll());

        doc.save(getArtifactsDir() + "PdfSaveOptions.ExportPageSet.All.pdf", options);
        //ExEnd

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.ExportPageSet.Even.pdf");
        TextAbsorber textAbsorber = new TextAbsorber();
        pdfDocument.Pages.Accept(textAbsorber);

        Assert.AreEqual("Page 2 (even)\r\n" +
                        "Page 4 (even)", textAbsorber.Text);

        pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.ExportPageSet.Odd.pdf");
        textAbsorber = new TextAbsorber();
        pdfDocument.Pages.Accept(textAbsorber);

        Assert.AreEqual("Page 1 (odd)\r\n" +
                        "Page 3 (odd)\r\n" +
                        "Page 5 (odd)", textAbsorber.Text);

        pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "PdfSaveOptions.ExportPageSet.All.pdf");
        textAbsorber = new TextAbsorber();
        pdfDocument.Pages.Accept(textAbsorber);

        Assert.AreEqual("Page 1 (odd)\r\n" +
                        "Page 2 (even)\r\n" +
                        "Page 3 (odd)\r\n" +
                        "Page 4 (even)\r\n" +
                        "Page 5 (odd)", textAbsorber.Text);
    }

    @Test
    public void exportLanguageToSpanTag() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.ExportLanguageToSpanTag
        //ExSummary:Shows how to create a "Span" tag in the document structure to export the text language.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello world!");
        builder.writeln("Hola mundo!");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        {
            // Note, when "ExportDocumentStructure" is false, "ExportLanguageToSpanTag" is ignored.
            saveOptions.setExportDocumentStructure(true); saveOptions.setExportLanguageToSpanTag(true);
        }

        doc.save(getArtifactsDir() + "PdfSaveOptions.ExportLanguageToSpanTag.pdf", saveOptions);
        //ExEnd
    }
}
