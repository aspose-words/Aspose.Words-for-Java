package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.pdf.Font;
import com.aspose.pdf.*;
import com.aspose.pdf.exceptions.InvalidPasswordException;
import com.aspose.pdf.facades.Bookmarks;
import com.aspose.pdf.facades.PdfBookmarkEditor;
import com.aspose.pdf.operators.SetGlyphsPositionShowText;
import com.aspose.words.Document;
import com.aspose.words.FolderFontSource;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.SaveOptions;
import com.aspose.words.WarningInfo;
import com.aspose.words.WarningType;
import com.aspose.words.*;
import org.apache.commons.collections4.IterableUtils;
import org.apache.poi.util.LocaleUtil;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExPdfSaveOptions extends ApiExampleBase {
    @Test
    public void onePage() throws Exception {
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

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Set the "PageIndex" to "1" to render a portion of the document starting from the second page.
        options.setPageSet(new PageSet(1));

        // This document will contain one page starting from page two, which will only contain the second page.
        doc.save(new FileOutputStream(getArtifactsDir() + "PdfSaveOptions.OnePage.pdf"), options);
        //ExEnd

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.OnePage.pdf");

        Assert.assertEquals(1, pdfDocument.getPages().size());

        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
        pdfDocument.getPages().accept(textFragmentAbsorber);

        Assert.assertEquals("Page 2.", textFragmentAbsorber.getText());

        pdfDocument.close();
    }

    @Test
    public void headingsOutlineLevels() throws Exception {
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
        bookmarkEditor.bindPdf(getArtifactsDir() + "PdfSaveOptions.HeadingsOutlineLevels.pdf");

        Bookmarks bookmarks = bookmarkEditor.extractBookmarks();

        Assert.assertEquals(3, bookmarks.size());

        bookmarkEditor.close();
    }

    @Test(dataProvider = "createMissingOutlineLevelsDataProvider")
    public void createMissingOutlineLevels(boolean createMissingOutlineLevels) throws Exception {
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
        bookmarkEditor.bindPdf(getArtifactsDir() + "PdfSaveOptions.CreateMissingOutlineLevels.pdf");

        Bookmarks bookmarks = bookmarkEditor.extractBookmarks();

        Assert.assertEquals(createMissingOutlineLevels ? 6 : 3, bookmarks.size());

        bookmarkEditor.close();
    }

    @DataProvider(name = "createMissingOutlineLevelsDataProvider")
    public static Object[][] createMissingOutlineLevelsDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "tableHeadingOutlinesDataProvider")
    public void tableHeadingOutlines(boolean createOutlinesForHeadingsInTables) throws Exception {
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

        com.aspose.pdf.Document pdfDoc = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.TableHeadingOutlines.pdf");

        if (createOutlinesForHeadingsInTables) {
            Assert.assertEquals(1, pdfDoc.getOutlines().size());
            Assert.assertEquals("Customers", pdfDoc.getOutlines().get_Item(1).getTitle());
        } else
            Assert.assertEquals(0, pdfDoc.getOutlines().size());

        TableAbsorber tableAbsorber = new TableAbsorber();
        tableAbsorber.visit(pdfDoc.getPages().get_Item(1));

        Assert.assertEquals("Customers", tableAbsorber.getTableList().get(0).getRowList().get(0).getCellList().get(0).getTextFragments().get_Item(1).getText());
        Assert.assertEquals("John Doe", tableAbsorber.getTableList().get(0).getRowList().get(1).getCellList().get(0).getTextFragments().get_Item(1).getText());
        Assert.assertEquals("Jane Doe", tableAbsorber.getTableList().get(0).getRowList().get(2).getCellList().get(0).getTextFragments().get_Item(1).getText());

        pdfDoc.close();
    }

    @DataProvider(name = "tableHeadingOutlinesDataProvider")
    public static Object[][] tableHeadingOutlinesDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void expandedOutlineLevels() throws Exception {
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

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.ExpandedOutlineLevels.pdf");

        Assert.assertEquals(1, pdfDocument.getOutlines().size());
        Assert.assertEquals(5, pdfDocument.getOutlines().getVisibleCount());

        Assert.assertTrue(pdfDocument.getOutlines().get_Item(1).getOpen());
        Assert.assertEquals(1, pdfDocument.getOutlines().get_Item(1).getLevel());

        Assert.assertFalse(pdfDocument.getOutlines().get_Item(1).get_Item(1).getOpen());
        Assert.assertEquals(2, pdfDocument.getOutlines().get_Item(1).get_Item(1).getLevel());

        Assert.assertTrue(pdfDocument.getOutlines().get_Item(1).get_Item(2).getOpen());
        Assert.assertEquals(2, pdfDocument.getOutlines().get_Item(1).get_Item(2).getLevel());

        pdfDocument.close();
    }

    @Test(dataProvider = "updateFieldsDataProvider")
    public void updateFields(boolean updateFields) throws Exception {
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

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.UpdateFields.pdf");

        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
        pdfDocument.getPages().accept(textFragmentAbsorber);

        Assert.assertEquals(updateFields ? "Page 1 of 2" : "Page  of ", textFragmentAbsorber.getTextFragments().get_Item(1).getText());

        pdfDocument.close();
    }

    @DataProvider(name = "updateFieldsDataProvider")
    public static Object[][] updateFieldsDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "preserveFormFieldsDataProvider")
    public void preserveFormFields(boolean preserveFormFields) throws Exception {
        //ExStart
        //ExFor:PdfSaveOptions.PreserveFormFields
        //ExSummary:Shows how to save a document to the PDF format using the Save method and the PdfSaveOptions class.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Please select a fruit: ");

        // Insert a combo box which will allow a user to choose an option from a collection of strings.
        builder.insertComboBox("MyComboBox", new String[]{"Apple", "Banana", "Cherry"}, 0);

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set the "PreserveFormFields" property to "true" to save form fields as interactive objects in the output PDF.
        // Set the "PreserveFormFields" property to "false" to freeze all form fields in the document at
        // their current values and display them as plain text in the output PDF.
        pdfOptions.setPreserveFormFields(preserveFormFields);

        doc.save(getArtifactsDir() + "PdfSaveOptions.PreserveFormFields.pdf", pdfOptions);
        //ExEnd

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.PreserveFormFields.pdf");

        Assert.assertEquals(1, pdfDocument.getPages().size());

        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
        pdfDocument.getPages().accept(textFragmentAbsorber);

        if (preserveFormFields) {
            Assert.assertEquals("Please select a fruit: ", textFragmentAbsorber.getText());
            TestUtil.fileContainsString("10 0 obj\r\n" +
                            "<</Type /Annot/Subtype /Widget/P 4 0 R/FT /Ch/F 4/Rect [169.54200745 706.20098877 219.02442932 721.49005127]/Ff 131072/T(��\u0000M\u0000y\0C\u0000o\u0000m\0b\u0000o\0B\u0000o\u0000x)/Opt " +
                            "[(��\0A\u0000p\u0000p\u0000l\0e) (��\0B\0a\u0000n\0a\u0000n\0a) (��\0C\u0000h\0e\u0000r\u0000r\u0000y) ]/V(��\0A\u0000p\u0000p\u0000l\0e)/DA(0 g /FAAABC 12 Tf )/AP<</N 11 0 R>>>>",
                    getArtifactsDir() + "PdfSaveOptions.PreserveFormFields.pdf");

            com.aspose.pdf.Form form = pdfDocument.getForm();
            Assert.assertEquals(1, form.size());

            ComboBoxField field = (ComboBoxField) form.getFields()[0];

            Assert.assertEquals("MyComboBox", field.getFullName());
            Assert.assertEquals(3, field.getOptions().size());
            Assert.assertEquals("Apple", field.getValue());
        } else {
            Assert.assertEquals("Please select a fruit: Apple", textFragmentAbsorber.getText());
            Assert.assertThrows(AssertionError.class, () -> TestUtil.fileContainsString("/Widget",
                    getArtifactsDir() + "PdfSaveOptions.PreserveFormFields.pdf"));

            Assert.assertEquals(0, pdfDocument.getForm().size());
        }

        pdfDocument.close();
    }

    @DataProvider(name = "preserveFormFieldsDataProvider")
    public static Object[][] preserveFormFieldsDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "complianceDataProvider")
    public void compliance(int pdfCompliance) throws Exception {
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

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.Compliance.pdf");

        switch (pdfCompliance) {
            case PdfCompliance.PDF_17:
                Assert.assertEquals(PdfFormat.v_1_7, pdfDocument.getPdfFormat());
                Assert.assertEquals("1.7", pdfDocument.getVersion());
                break;
            case PdfCompliance.PDF_A_1_A:
                Assert.assertEquals(PdfFormat.PDF_A_1A, pdfDocument.getPdfFormat());
                Assert.assertEquals("1.4", pdfDocument.getVersion());
                break;
            case PdfCompliance.PDF_A_1_B:
                Assert.assertEquals(PdfFormat.PDF_A_1B, pdfDocument.getPdfFormat());
                Assert.assertEquals("1.4", pdfDocument.getVersion());
                break;
        }

        pdfDocument.close();
    }

    @DataProvider(name = "complianceDataProvider")
    public static Object[][] complianceDataProvider() {
        return new Object[][]
                {
                        {PdfCompliance.PDF_A_1_B},
                        {PdfCompliance.PDF_17},
                        {PdfCompliance.PDF_A_1_A},
                };
    }

    @Test(dataProvider = "textCompressionDataProvider")
    public void textCompression(int pdfTextCompression) throws Exception {
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

        switch (pdfTextCompression) {
            case PdfTextCompression.NONE:
                Assert.assertTrue(new File(getArtifactsDir() + "PdfSaveOptions.TextCompression.pdf").length() < 68000);
                TestUtil.fileContainsString("5 0 obj\r\n<</Length 9 0 R>>stream",
                        getArtifactsDir() + "PdfSaveOptions.TextCompression.pdf");
                break;
            case PdfTextCompression.FLATE:
                Assert.assertTrue(new File(getArtifactsDir() + "PdfSaveOptions.TextCompression.pdf").length() < 30000);
                TestUtil.fileContainsString("5 0 obj\r\n<</Length 9 0 R/Filter /FlateDecode>>stream",
                        getArtifactsDir() + "PdfSaveOptions.TextCompression.pdf");
                break;
        }
    }

    @DataProvider(name = "textCompressionDataProvider")
    public static Object[][] textCompressionDataProvider() {
        return new Object[][]
                {
                        {PdfTextCompression.NONE},
                        {PdfTextCompression.FLATE},
                };
    }

    @Test(enabled = false, dataProvider = "imageCompressionDataProvider")
    public void imageCompression(int pdfImageCompression) throws Exception {
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

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.ImageCompression.pdf");

        try (InputStream firstImage = pdfDocument.getPages().get_Item(1).getResources().getImages().get_Item(1).toStream()) {
            TestUtil.verifyImage(400, 400, firstImage);
        }

        switch (pdfImageCompression) {
            case PdfImageCompression.AUTO:
                Assert.assertTrue(new File(getArtifactsDir() + "PdfSaveOptions.ImageCompression.pdf").length() < 51500);
                break;
            case PdfImageCompression.JPEG:
                Assert.assertTrue(new File(getArtifactsDir() + "PdfSaveOptions.ImageCompression.pdf").length() <= 42000);
                try (InputStream secondImage = pdfDocument.getPages().get_Item(1).getResources().getImages().get_Item(2).toStream()) {
                    TestUtil.verifyImage(400, 400, secondImage);
                }
                break;
        }

        pdfDocument.close();
    }

    @DataProvider(name = "imageCompressionDataProvider")
    public static Object[][] imageCompressionDataProvider() {
        return new Object[][]
                {
                        {PdfImageCompression.AUTO},
                        {PdfImageCompression.JPEG},
                };
    }

    @Test(dataProvider = "imageColorSpaceExportModeDataProvider")
    public void imageColorSpaceExportMode(int pdfImageColorSpaceExportMode) throws Exception {
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

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.ImageColorSpaceExportMode.pdf");
        XImage pdfDocImage = pdfDocument.getPages().get_Item(1).getResources().getImages().get_Item(1);

        Assert.assertEquals(400, pdfDocImage.getWidth());
        Assert.assertEquals(400, pdfDocImage.getHeight());
        Assert.assertEquals(ColorType.Rgb, pdfDocImage.getColorType());

        pdfDocImage = pdfDocument.getPages().get_Item(1).getResources().getImages().get_Item(2);

        Assert.assertEquals(400, pdfDocImage.getWidth());
        Assert.assertEquals(400, pdfDocImage.getHeight());
        Assert.assertEquals(ColorType.Rgb, pdfDocImage.getColorType());

        pdfDocument.close();
    }

    @DataProvider(name = "imageColorSpaceExportModeDataProvider")
    public static Object[][] imageColorSpaceExportModeDataProvider() {
        return new Object[][]
                {
                        {PdfImageColorSpaceExportMode.AUTO},
                        {PdfImageColorSpaceExportMode.SIMPLE_CMYK},
                };
    }

    @Test
    public void downsampleOptions() throws Exception {
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

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.DownsampleOptions.Default.pdf");
        XImage pdfDocImage = pdfDocument.getPages().get_Item(1).getResources().getImages().get_Item(1);

        Assert.assertEquals(ColorType.Rgb, pdfDocImage.getColorType());

        pdfDocument.close();
    }

    @Test(dataProvider = "colorRenderingDataProvider")
    public void colorRendering(int colorMode) throws Exception {
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
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        {
            pdfSaveOptions.setColorMode(colorMode);
        }

        doc.save(getArtifactsDir() + "PdfSaveOptions.ColorRendering.pdf", pdfSaveOptions);
        //ExEnd

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.ColorRendering.pdf");
        XImage pdfDocImage = pdfDocument.getPages().get_Item(1).getResources().getImages().get_Item(1);

        switch (colorMode) {
            case ColorMode.NORMAL:
                Assert.assertEquals(ColorType.Rgb, pdfDocImage.getColorType());
                break;
            case ColorMode.GRAYSCALE:
                Assert.assertEquals(ColorType.Grayscale, pdfDocImage.getColorType());
                break;
        }

        pdfDocument.close();
    }

    @DataProvider(name = "colorRenderingDataProvider")
    public static Object[][] colorRenderingDataProvider() {
        return new Object[][]
                {
                        {ColorMode.GRAYSCALE},
                        {ColorMode.NORMAL},
                };
    }

    @Test(dataProvider = "docTitleDataProvider")
    public void docTitle(boolean displayDocTitle) throws Exception {
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
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        {
            pdfSaveOptions.setDisplayDocTitle(displayDocTitle);
        }

        doc.save(getArtifactsDir() + "PdfSaveOptions.DocTitle.pdf", pdfSaveOptions);
        //ExEnd

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.DocTitle.pdf");

        Assert.assertEquals(displayDocTitle, pdfDocument.isDisplayDocTitle());
        Assert.assertEquals("Windows bar pdf title", pdfDocument.getInfo().getTitle());

        pdfDocument.close();
    }

    @DataProvider(name = "docTitleDataProvider")
    public static Object[][] docTitleDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "memoryOptimizationDataProvider")
    public void memoryOptimization(boolean memoryOptimization) throws Exception {
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

    @DataProvider(name = "memoryOptimizationDataProvider")
    public static Object[][] memoryOptimizationDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "escapeUriDataProvider")
    public void escapeUri(String uri, String result) throws Exception {
        //ExStart
        //ExFor:PdfSaveOptions
        //ExSummary:Shows how to escape hyperlinks in the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertHyperlink("Testlink", uri, false);

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        doc.save(getArtifactsDir() + "PdfSaveOptions.EscapedUri.pdf", options);
        //ExEnd

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.EscapedUri.pdf");

        Page page = pdfDocument.getPages().get_Item(1);
        LinkAnnotation linkAnnot = (LinkAnnotation) page.getAnnotations().get_Item(1);

        GoToURIAction action = (GoToURIAction) linkAnnot.getAction();

        Assert.assertEquals(result, action.getURI());

        pdfDocument.close();
    }

    @DataProvider(name = "escapeUriDataProvider")
    public static Object[][] escapeUriDataProvider() {
        return new Object[][]
                {
                        {"https://www.google.com/search?q= aspose",  "https://www.google.com/search?q=%20aspose"},
                        {"https://www.google.com/search?q=%20aspose",  "https://www.google.com/search?q=%20aspose"},
                };
    }

    @Test(dataProvider = "openHyperlinksInNewWindowDataProvider")
    public void openHyperlinksInNewWindow(boolean openHyperlinksInNewWindow) throws Exception {
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
                    "<</Type/Border/S/S/W 0>>/A<</Type /Action/S /JavaScript/JS(app.launchURL\\(\"https://www.google.com/search?q=%20aspose\", true\\);)>>>>",
                    getArtifactsDir() + "PdfSaveOptions.OpenHyperlinksInNewWindow.pdf");
        else
            TestUtil.fileContainsString(
                    "<</Type/Border/S/S/W 0>>/A<</Type /Action/S /URI/URI(https://www.google.com/search?q=%20aspose)>>>>",
                    getArtifactsDir() + "PdfSaveOptions.OpenHyperlinksInNewWindow.pdf");

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.OpenHyperlinksInNewWindow.pdf");

        Page page = pdfDocument.getPages().get_Item(1);
        LinkAnnotation linkAnnot = (LinkAnnotation) page.getAnnotations().get_Item(1);

        Assert.assertEquals(openHyperlinksInNewWindow ? JavascriptAction.class : GoToURIAction.class,
                linkAnnot.getAction().getClass());

        pdfDocument.close();
    }

    @DataProvider(name = "openHyperlinksInNewWindowDataProvider")
    public static Object[][] openHyperlinksInNewWindowDataProvider() {
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
    @Test(groups = "SkipMono") //ExSkip
    public void handleBinaryRasterWarnings() throws Exception {
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

        Assert.assertEquals(1, callback.mWarnings.getCount());
        Assert.assertEquals("'R2_XORPEN' binary raster operation is partly supported.",
                callback.mWarnings.get(0).getDescription());
    }

    /// <summary>
    /// Prints and collects formatting loss-related warnings that occur upon saving a document.
    /// </summary>
    public static class HandleDocumentWarnings implements IWarningCallback {
        public void warning(WarningInfo info) {
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

        com.aspose.pdf.Document pdfDoc = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf");
        String inputDocLocaleName = LocaleUtil.getLocaleFromLCID(doc.getStyles().getDefaultFont().getLocaleId());

        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
        pdfDoc.getPages().accept(textFragmentAbsorber);
        switch (headerFooterBookmarksExportMode) {
            case com.aspose.words.HeaderFooterBookmarksExportMode.NONE:
                TestUtil.fileContainsString(MessageFormat.format("<</Type /Catalog/Pages 3 0 R/Lang({0})>>\r\n", inputDocLocaleName),
                        getArtifactsDir() + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf");

                Assert.assertEquals(0, pdfDoc.getOutlines().size());
                break;
            case com.aspose.words.HeaderFooterBookmarksExportMode.FIRST:
            case com.aspose.words.HeaderFooterBookmarksExportMode.ALL:
                TestUtil.fileContainsString(
                        MessageFormat.format("<</Type /Catalog/Pages 3 0 R/Outlines 13 0 R/PageMode /UseOutlines/Lang({0})>>", inputDocLocaleName),
                        getArtifactsDir() + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf");

                OutlineCollection outlineItemCollection = pdfDoc.getOutlines();

                Assert.assertEquals(4, outlineItemCollection.size());
                Assert.assertEquals("Bookmark_1", outlineItemCollection.get_Item(1).getTitle());
                Assert.assertEquals("1 XYZ 233 806 0", outlineItemCollection.get_Item(1).getDestination().toString());

                Assert.assertEquals("Bookmark_2", outlineItemCollection.get_Item(2).getTitle());
                Assert.assertEquals("1 XYZ 84 47 0", outlineItemCollection.get_Item(2).getDestination().toString());

                Assert.assertEquals("Bookmark_3", outlineItemCollection.get_Item(3).getTitle());
                Assert.assertEquals("2 XYZ 85 806 0", outlineItemCollection.get_Item(3).getDestination().toString());

                Assert.assertEquals("Bookmark_4", outlineItemCollection.get_Item(4).getTitle());
                Assert.assertEquals("2 XYZ 85 48 0", outlineItemCollection.get_Item(4).getDestination().toString());
                break;
        }
        pdfDoc.close();
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
    public void fontsScaledToMetafileSize(boolean scaleWmfFonts) throws Exception {
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

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.FontsScaledToMetafileSize.pdf");
        TextFragmentAbsorber textAbsorber = new TextFragmentAbsorber();

        pdfDocument.getPages().get_Item(1).accept(textAbsorber);
        Rectangle textFragmentRectangle = textAbsorber.getTextFragments().get_Item(3).getRectangle();

        Assert.assertEquals(scaleWmfFonts ? 1.589d : 5.045d, textFragmentRectangle.getWidth(), 0.001d);

        pdfDocument.close();
    }

    @DataProvider(name = "fontsScaledToMetafileSizeDataProvider")
    public static Object[][] fontsScaledToMetafileSizeDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "embedFullFontsDataProvider")
    public void embedFullFonts(boolean embedFullFonts) throws Exception {
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
        FontSettings.getDefaultInstance().setFontsSources(new FontSourceBase[]{originalFontsSources[0], folderFontSource});

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
            Assert.assertTrue(new File(getArtifactsDir() + "PdfSaveOptions.EmbedFullFonts.pdf").length() < 571000);
        else
            Assert.assertTrue(new File(getArtifactsDir() + "PdfSaveOptions.EmbedFullFonts.pdf").length() < 25000);

        // Restore the original font sources.
        FontSettings.getDefaultInstance().setFontsSources(originalFontsSources);
        //ExEnd

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.EmbedFullFonts.pdf");

        Font[] pdfDocFonts = pdfDocument.getFontUtilities().getAllFonts();

        Assert.assertEquals("ArialMT", pdfDocFonts[0].getFontName());
        Assert.assertNotEquals(embedFullFonts, pdfDocFonts[0].isSubset());

        Assert.assertEquals("Arvo", pdfDocFonts[1].getFontName());
        Assert.assertNotEquals(embedFullFonts, pdfDocFonts[1].isSubset());

        pdfDocument.close();
    }

    @DataProvider(name = "embedFullFontsDataProvider")
    public static Object[][] embedFullFontsDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "embedWindowsFontsDataProvider")
    public void embedWindowsFonts(int pdfFontEmbeddingMode) throws Exception {
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

        switch (pdfFontEmbeddingMode) {
            case PdfFontEmbeddingMode.EMBED_ALL:
                Assert.assertTrue(new File(getArtifactsDir() + "PdfSaveOptions.EmbedWindowsFonts.pdf").length() < 1030500);
                break;
            case PdfFontEmbeddingMode.EMBED_NONSTANDARD:
                Assert.assertTrue(new File(getArtifactsDir() + "PdfSaveOptions.EmbedWindowsFonts.pdf").length() < 491000);
                break;
            case PdfFontEmbeddingMode.EMBED_NONE:
                Assert.assertTrue(new File(getArtifactsDir() + "PdfSaveOptions.EmbedWindowsFonts.pdf").length() <= 4000);
                break;
        }
        //ExEnd

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.EmbedWindowsFonts.pdf");

        com.aspose.pdf.Font[] pdfDocFonts = pdfDocument.getFontUtilities().getAllFonts();

        Assert.assertEquals("ArialMT", pdfDocFonts[0].getFontName());
        Assert.assertEquals(pdfFontEmbeddingMode == PdfFontEmbeddingMode.EMBED_ALL,
                pdfDocFonts[0].isEmbedded());

        Assert.assertEquals("CourierNewPSMT", pdfDocFonts[1].getFontName());
        Assert.assertEquals(pdfFontEmbeddingMode == PdfFontEmbeddingMode.EMBED_ALL || pdfFontEmbeddingMode == PdfFontEmbeddingMode.EMBED_NONSTANDARD,
                pdfDocFonts[1].isEmbedded());

        pdfDocument.close();
    }

    @DataProvider(name = "embedWindowsFontsDataProvider")
    public static Object[][] embedWindowsFontsDataProvider() {
        return new Object[][]
                {
                        {PdfFontEmbeddingMode.EMBED_ALL},
                        {PdfFontEmbeddingMode.EMBED_NONE},
                        {PdfFontEmbeddingMode.EMBED_NONSTANDARD},
                };
    }

    @Test(dataProvider = "embedCoreFontsDataProvider")
    public void embedCoreFonts(boolean useCoreFonts) throws Exception {
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
            Assert.assertTrue(new File(getArtifactsDir() + "PdfSaveOptions.EmbedCoreFonts.pdf").length() < 3000);
        else
            Assert.assertTrue(new File(getArtifactsDir() + "PdfSaveOptions.EmbedCoreFonts.pdf").length() < 32500);
        //ExEnd

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.EmbedCoreFonts.pdf");

        Font[] pdfDocFonts = pdfDocument.getFontUtilities().getAllFonts();

        if (useCoreFonts) {
            Assert.assertEquals("Helvetica", pdfDocFonts[0].getFontName());
            Assert.assertEquals("Courier", pdfDocFonts[1].getFontName());
        } else {
            Assert.assertEquals("ArialMT", pdfDocFonts[0].getFontName());
            Assert.assertEquals("CourierNewPSMT", pdfDocFonts[1].getFontName());
        }

        Assert.assertNotEquals(useCoreFonts, pdfDocFonts[0].isEmbedded());
        Assert.assertNotEquals(useCoreFonts, pdfDocFonts[1].isEmbedded());

        pdfDocument.close();
    }

    @DataProvider(name = "embedCoreFontsDataProvider")
    public static Object[][] embedCoreFontsDataProvider() {
        return new Object[][]
                {
                        {true},
                        {false}
                };
    }

    @Test(dataProvider = "additionalTextPositioningDataProvider")
    public void additionalTextPositioning(boolean applyAdditionalTextPositioning) throws Exception {
        //ExStart
        //ExFor:PdfSaveOptions.AdditionalTextPositioning
        //ExSummary:Show how to write additional text positioning operators.
        Document doc = new Document(getMyDir() + "Text positioning operators.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setTextCompression(PdfTextCompression.NONE);

        // Set the "AdditionalTextPositioning" property to "true" to attempt to fix incorrect
        // element positioning in the output PDF, should there be any, at the cost of increased file size.
        // Set the "AdditionalTextPositioning" property to "false" to render the document as usual.
        saveOptions.setAdditionalTextPositioning(applyAdditionalTextPositioning);

        doc.save(getArtifactsDir() + "PdfSaveOptions.AdditionalTextPositioning.pdf", saveOptions);
        //ExEnd

        com.aspose.pdf.Document pdfDocument =
                new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.AdditionalTextPositioning.pdf");
        TextFragmentAbsorber textAbsorber = new TextFragmentAbsorber();

        pdfDocument.getPages().get_Item(1).accept(textAbsorber);

        SetGlyphsPositionShowText tjOperator =
                (SetGlyphsPositionShowText) textAbsorber.getTextFragments().get_Item(1).getPage().getContents().get_Item(85);

        if (applyAdditionalTextPositioning) {
            Assert.assertTrue(new File(getArtifactsDir() + "PdfSaveOptions.AdditionalTextPositioning.pdf").length() < 101000);
            Assert.assertEquals(
                    "[0 (S) 0 (a) 0 (m) 0 (s) 0 (t) 0 (a) -1 (g) 1 (,) 0 ( ) 0 (1) 0 (0) 0 (.) 0 ( ) 0 (N) 0 (o) 0 (v) 0 (e) 0 (m) 0 (b) 0 (e) 0 (r) -1 ( ) 1 (2) -1 (0) 0 (1) 0 (8)] TJ",
                    tjOperator.toString());
        } else {
            Assert.assertTrue(new File(getArtifactsDir() + "PdfSaveOptions.AdditionalTextPositioning.pdf").length() < 98200);
            Assert.assertEquals("[(Samsta) -1 (g) 1 (, 10. November) -1 ( ) 1 (2) -1 (018)] TJ", tjOperator.toString());
        }

        pdfDocument.close();
    }

    @DataProvider(name = "additionalTextPositioningDataProvider")
    public static Object[][] additionalTextPositioningDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "saveAsPdfBookFoldDataProvider")
    public void saveAsPdfBookFold(boolean renderTextAsBookfold) throws Exception {
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
            for (Section s : doc.getSections()) {
                s.getPageSetup().setMultiplePages(MultiplePagesType.BOOK_FOLD_PRINTING);
            }

        // Once we print this document on both sides of the pages, we can fold all the pages down the middle at once,
        // and the contents will line up in a way that creates a booklet.
        doc.save(getArtifactsDir() + "PdfSaveOptions.SaveAsPdfBookFold.pdf", options);
        //ExEnd

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.SaveAsPdfBookFold.pdf");
        TextAbsorber textAbsorber = new TextAbsorber();

        pdfDocument.getPages().accept(textAbsorber);

        if (renderTextAsBookfold) {
            Assert.assertTrue(textAbsorber.getText().indexOf("Heading #1") < textAbsorber.getText().indexOf("Heading #2"));
            Assert.assertTrue(textAbsorber.getText().indexOf("Heading #2") < textAbsorber.getText().indexOf("Heading #3"));
            Assert.assertTrue(textAbsorber.getText().indexOf("Heading #3") < textAbsorber.getText().indexOf("Heading #4"));
            Assert.assertTrue(textAbsorber.getText().indexOf("Heading #4") < textAbsorber.getText().indexOf("Heading #5"));
            Assert.assertTrue(textAbsorber.getText().indexOf("Heading #5") < textAbsorber.getText().indexOf("Heading #6"));
            Assert.assertTrue(textAbsorber.getText().indexOf("Heading #6") < textAbsorber.getText().indexOf("Heading #7"));
            Assert.assertFalse(textAbsorber.getText().indexOf("Heading #7") < textAbsorber.getText().indexOf("Heading #8"));
            Assert.assertTrue(textAbsorber.getText().indexOf("Heading #8") < textAbsorber.getText().indexOf("Heading #9"));
            Assert.assertFalse(textAbsorber.getText().indexOf("Heading #9") < textAbsorber.getText().indexOf("Heading #10"));
        } else {
            Assert.assertTrue(textAbsorber.getText().indexOf("Heading #1") < textAbsorber.getText().indexOf("Heading #2"));
            Assert.assertTrue(textAbsorber.getText().indexOf("Heading #2") < textAbsorber.getText().indexOf("Heading #3"));
            Assert.assertTrue(textAbsorber.getText().indexOf("Heading #3") < textAbsorber.getText().indexOf("Heading #4"));
            Assert.assertTrue(textAbsorber.getText().indexOf("Heading #4") < textAbsorber.getText().indexOf("Heading #5"));
            Assert.assertTrue(textAbsorber.getText().indexOf("Heading #5") < textAbsorber.getText().indexOf("Heading #6"));
            Assert.assertTrue(textAbsorber.getText().indexOf("Heading #6") < textAbsorber.getText().indexOf("Heading #7"));
            Assert.assertTrue(textAbsorber.getText().indexOf("Heading #7") < textAbsorber.getText().indexOf("Heading #8"));
            Assert.assertTrue(textAbsorber.getText().indexOf("Heading #8") < textAbsorber.getText().indexOf("Heading #9"));
            Assert.assertTrue(textAbsorber.getText().indexOf("Heading #9") < textAbsorber.getText().indexOf("Heading #10"));
        }

        pdfDocument.close();
    }

    @DataProvider(name = "saveAsPdfBookFoldDataProvider")
    public static Object[][] saveAsPdfBookFoldDataProvider() {
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

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.ZoomBehaviour.pdf");
        GoToAction action = (GoToAction) pdfDocument.getOpenAction();

        Assert.assertEquals(0.25d, ((XYZExplicitDestination) action.getDestination()).getZoom());

        pdfDocument.close();
    }

    @Test(dataProvider = "pageModeDataProvider")
    public void pageMode(int pageMode) throws Exception {
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

        String docLocaleName = LocaleUtil.getLocaleFromLCID(doc.getStyles().getDefaultFont().getLocaleId());

        switch (pageMode) {
            case PdfPageMode.FULL_SCREEN:
                TestUtil.fileContainsString(
                        MessageFormat.format("<</Type /Catalog/Pages 3 0 R/PageMode /FullScreen/Lang({0})>>\r\n", docLocaleName),
                        getArtifactsDir() + "PdfSaveOptions.PageMode.pdf");
                break;
            case PdfPageMode.USE_THUMBS:
                TestUtil.fileContainsString(
                        MessageFormat.format("<</Type /Catalog/Pages 3 0 R/PageMode /UseThumbs/Lang({0})>>", docLocaleName),
                        getArtifactsDir() + "PdfSaveOptions.PageMode.pdf");
                break;
            case PdfPageMode.USE_OC:
                TestUtil.fileContainsString(
                        MessageFormat.format("<</Type /Catalog/Pages 3 0 R/PageMode /UseOC/Lang({0})>>\r\n", docLocaleName),
                        getArtifactsDir() + "PdfSaveOptions.PageMode.pdf");
                break;
            case PdfPageMode.USE_OUTLINES:
            case PdfPageMode.USE_NONE:
                TestUtil.fileContainsString(MessageFormat.format("<</Type /Catalog/Pages 3 0 R/Lang({0})>>\r\n", docLocaleName),
                        getArtifactsDir() + "PdfSaveOptions.PageMode.pdf");
                break;
        }

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.PageMode.pdf");

        switch (pageMode) {
            case PdfPageMode.USE_NONE:
            case PdfPageMode.USE_OUTLINES:
                Assert.assertEquals(PageMode.UseNone, pdfDocument.getPageMode());
                break;
            case PdfPageMode.USE_THUMBS:
                Assert.assertEquals(PageMode.UseThumbs, pdfDocument.getPageMode());
                break;
            case PdfPageMode.FULL_SCREEN:
                Assert.assertEquals(PageMode.FullScreen, pdfDocument.getPageMode());
                break;
            case PdfPageMode.USE_OC:
                Assert.assertEquals(PageMode.UseOC, pdfDocument.getPageMode());
                break;
        }

        pdfDocument.close();
    }

    @DataProvider(name = "pageModeDataProvider")
    public static Object[][] pageModeDataProvider() {
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
    public void noteHyperlinks(boolean createNoteHyperlinks) throws Exception {
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

        if (createNoteHyperlinks) {
            TestUtil.fileContainsString(
                    "<</Type /Annot/Subtype /Link/Rect [157.80099487 720.90106201 159.35600281 733.55004883]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 85 677 0]>>",
                    getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString(
                    "<</Type /Annot/Subtype /Link/Rect [202.16900635 720.90106201 206.06201172 733.55004883]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 85 79 0]>>",
                    getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString(
                    "<</Type /Annot/Subtype /Link/Rect [212.23199463 699.2510376 215.34199524 711.90002441]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 85 654 0]>>",
                    getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString(
                    "<</Type /Annot/Subtype /Link/Rect [258.15499878 699.2510376 262.04800415 711.90002441]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 85 68 0]>>",
                    getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString(
                    "<</Type /Annot/Subtype /Link/Rect [85.05000305 68.19904327 88.66500092 79.69804382]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 202 733 0]>>",
                    getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString(
                    "<</Type /Annot/Subtype /Link/Rect [85.05000305 56.70004272 88.66500092 68.19904327]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 258 711 0]>>",
                    getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString(
                    "<</Type /Annot/Subtype /Link/Rect [85.05000305 666.10205078 86.4940033 677.60107422]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 157 733 0]>>",
                    getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil.fileContainsString(
                    "<</Type /Annot/Subtype /Link/Rect [85.05000305 643.10406494 87.93800354 654.60308838]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 212 711 0]>>",
                    getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
        } else {
            Assert.assertThrows(AssertionError.class, () -> TestUtil.fileContainsString("<</Type /Annot/Subtype /Link/Rect",
                    getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf"));
        }

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.NoteHyperlinks.pdf");
        Page page = pdfDocument.getPages().get_Item(1);
        AnnotationSelector annotationSelector = new AnnotationSelector(new LinkAnnotation(page, Rectangle.getTrivial()));

        page.accept(annotationSelector);

        List<LinkAnnotation> linkAnnotations = (List<LinkAnnotation>) (List<?>) annotationSelector.getSelected();

        if (createNoteHyperlinks) {
            Assert.assertEquals(8, IterableUtils.countMatches(linkAnnotations, a -> a.getAnnotationType() == AnnotationType.Link));

            Assert.assertEquals("1 XYZ 85 677 0", linkAnnotations.get(0).getDestination().toString());
            Assert.assertEquals("1 XYZ 85 79 0", linkAnnotations.get(1).getDestination().toString());
            Assert.assertEquals("1 XYZ 85 654 0", linkAnnotations.get(2).getDestination().toString());
            Assert.assertEquals("1 XYZ 85 68 0", linkAnnotations.get(3).getDestination().toString());
            Assert.assertEquals("1 XYZ 202 733 0", linkAnnotations.get(4).getDestination().toString());
            Assert.assertEquals("1 XYZ 258 711 0", linkAnnotations.get(5).getDestination().toString());
            Assert.assertEquals("1 XYZ 157 733 0", linkAnnotations.get(6).getDestination().toString());
            Assert.assertEquals("1 XYZ 212 711 0", linkAnnotations.get(7).getDestination().toString());
        } else {
            Assert.assertEquals(0, annotationSelector.getSelected().size());
        }

        pdfDocument.close();
    }

    @DataProvider(name = "noteHyperlinksDataProvider")
    public static Object[][] noteHyperlinksDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "customPropertiesExportDataProvider")
    public void customPropertiesExport(int pdfCustomPropertiesExportMode) throws Exception {
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

        switch (pdfCustomPropertiesExportMode) {
            case PdfCustomPropertiesExport.NONE:
                Assert.assertThrows(AssertionError.class, () -> TestUtil.fileContainsString(
                        doc.getCustomDocumentProperties().get(0).getName(),
                        getArtifactsDir() + "PdfSaveOptions.CustomPropertiesExport.pdf"));
                Assert.assertThrows(AssertionError.class, () -> TestUtil.fileContainsString(
                        "<</Type /Metadata/Subtype /XML/Length 8 0 R/Filter /FlateDecode>>",
                        getArtifactsDir() + "PdfSaveOptions.CustomPropertiesExport.pdf"));
                break;
            case PdfCustomPropertiesExport.STANDARD:
                TestUtil.fileContainsString(
                        "<</Creator(��\0A\u0000s\u0000p\u0000o\u0000s\0e\u0000.\u0000W\u0000o\u0000r\0d\u0000s)/Producer(��\0A\u0000s\u0000p\u0000o\u0000s\0e\u0000.\u0000W\u0000o\u0000r\0d\u0000s\u0000 \0f\u0000o\u0000r\u0000",
                        getArtifactsDir() + "PdfSaveOptions.CustomPropertiesExport.pdf");
                TestUtil.fileContainsString("/Company (��\u0000M\u0000y\u0000 \u0000v\0a\u0000l\u0000u\0e)>>",
                        getArtifactsDir() + "PdfSaveOptions.CustomPropertiesExport.pdf");
                break;
            case PdfCustomPropertiesExport.METADATA:
                TestUtil.fileContainsString("<</Type /Metadata/Subtype /XML/Length 8 0 R/Filter /FlateDecode>>",
                        getArtifactsDir() + "PdfSaveOptions.CustomPropertiesExport.pdf");
                break;
        }

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.CustomPropertiesExport.pdf");

        Assert.assertEquals("Aspose.Words", pdfDocument.getInfo().getCreator());
        Assert.assertTrue(pdfDocument.getInfo().getProducer().startsWith("Aspose.Words"));

        switch (pdfCustomPropertiesExportMode) {
            case PdfCustomPropertiesExport.NONE:
                Assert.assertEquals(2, pdfDocument.getInfo().size());
                Assert.assertEquals(0, pdfDocument.getMetadata().size());
                break;
            case PdfCustomPropertiesExport.METADATA:
                Assert.assertEquals(2, pdfDocument.getInfo().size());
                Assert.assertEquals(2, pdfDocument.getMetadata().size());

                Assert.assertEquals("Aspose.Words", pdfDocument.getMetadata().get_Item("xmp:CreatorTool").toString());
                Assert.assertEquals("Company", pdfDocument.getMetadata().get_Item("custprops:Property1").toString());
                break;
            case PdfCustomPropertiesExport.STANDARD:
                Assert.assertEquals(3, pdfDocument.getInfo().size());
                Assert.assertEquals(0, pdfDocument.getMetadata().size());

                Assert.assertEquals("My value", pdfDocument.getInfo().get_Item("Company"));
                break;
        }

        pdfDocument.close();
    }

    @DataProvider(name = "customPropertiesExportDataProvider")
    public static Object[][] customPropertiesExportDataProvider() {
        return new Object[][]
                {
                        {PdfCustomPropertiesExport.NONE},
                        {PdfCustomPropertiesExport.STANDARD},
                        {PdfCustomPropertiesExport.METADATA},
                };
    }

    @Test(dataProvider = "drawingMLEffectsDataProvider")
    public void drawingMLEffects(int effectsRenderingMode) throws Exception {
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

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.DrawingMLEffects.pdf");

        ImagePlacementAbsorber imb = new ImagePlacementAbsorber();
        imb.visit(pdfDocument.getPages().get_Item(1));

        TableAbsorber ttb = new TableAbsorber();
        ttb.visit(pdfDocument.getPages().get_Item(1));

        switch (effectsRenderingMode) {
            case DmlEffectsRenderingMode.NONE:
            case DmlEffectsRenderingMode.SIMPLIFIED:
                TestUtil.fileContainsString("4 0 obj\r\n" +
                                "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                        getArtifactsDir() + "PdfSaveOptions.DrawingMLEffects.pdf");
                Assert.assertEquals(0, imb.getImagePlacements().size());
                Assert.assertEquals(28, ttb.getTableList().size());
                break;
            case DmlEffectsRenderingMode.FINE:
                TestUtil.fileContainsString(
                        "4 0 obj\r\n<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R>>/XObject<</X1 9 0 R/X2 10 0 R/X3 11 0 R/X4 12 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                        getArtifactsDir() + "PdfSaveOptions.DrawingMLEffects.pdf");
                Assert.assertEquals(21, imb.getImagePlacements().size());
                Assert.assertEquals(4, ttb.getTableList().size());
                break;
        }

        pdfDocument.close();
    }

    @DataProvider(name = "drawingMLEffectsDataProvider")
    public static Object[][] drawingMLEffectsDataProvider() {
        return new Object[][]
                {
                        {DmlEffectsRenderingMode.NONE},
                        {DmlEffectsRenderingMode.SIMPLIFIED},
                        {DmlEffectsRenderingMode.FINE},
                };
    }

    @Test(dataProvider = "drawingMLFallbackDataProvider")
    public void drawingMLFallback(int dmlRenderingMode) throws Exception {
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

        switch (dmlRenderingMode) {
            case DmlRenderingMode.DRAWING_ML:
                TestUtil.fileContainsString(
                        "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABA 10 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                        getArtifactsDir() + "PdfSaveOptions.DrawingMLFallback.pdf");
                break;
            case DmlRenderingMode.FALLBACK:
                TestUtil.fileContainsString(
                        "4 0 obj\r\n<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABC 12 0 R>>/ExtGState<</GS1 9 0 R/GS2 10 0 R>>>>/Group ",
                        getArtifactsDir() + "PdfSaveOptions.DrawingMLFallback.pdf");
                break;
        }

        com.aspose.pdf.Document pdfDocument =
                new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.DrawingMLFallback.pdf");

        ImagePlacementAbsorber imagePlacementAbsorber = new ImagePlacementAbsorber();
        imagePlacementAbsorber.visit(pdfDocument.getPages().get_Item(1));

        TableAbsorber tableAbsorber = new TableAbsorber();
        tableAbsorber.visit(pdfDocument.getPages().get_Item(1));

        switch (dmlRenderingMode) {
            case DmlRenderingMode.DRAWING_ML:
                Assert.assertEquals(6, tableAbsorber.getTableList().size());
                break;
            case DmlRenderingMode.FALLBACK:
                Assert.assertEquals(15, tableAbsorber.getTableList().size());
                break;
        }

        pdfDocument.close();
    }

    @DataProvider(name = "drawingMLFallbackDataProvider")
    public static Object[][] drawingMLFallbackDataProvider() {
        return new Object[][]
                {
                        {DmlRenderingMode.FALLBACK},
                        {DmlRenderingMode.DRAWING_ML},
                };
    }

    @Test(dataProvider = "exportDocumentStructureDataProvider")
    public void exportDocumentStructure(boolean exportDocumentStructure) throws Exception {
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

        if (exportDocumentStructure) {
            TestUtil.fileContainsString("4 0 obj\r\n" +
                            "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABB 11 0 R>>/ExtGState<</GS1 9 0 R/GS2 13 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>/StructParents 0/Tabs /S>>",
                    getArtifactsDir() + "PdfSaveOptions.ExportDocumentStructure.pdf");
        } else {
            TestUtil.fileContainsString("4 0 obj\r\n" +
                            "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABA 10 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                    getArtifactsDir() + "PdfSaveOptions.ExportDocumentStructure.pdf");
        }
    }

    @DataProvider(name = "exportDocumentStructureDataProvider")
    public static Object[][] exportDocumentStructureDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "preblendImagesDataProvider")
    public void preblendImages(boolean preblendImages) throws Exception {
        //ExStart
        //ExFor:PdfSaveOptions.PreblendImages
        //ExSummary:Shows how to preblend images with transparent backgrounds while saving a document to PDF.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertImage(getImageDir() + "Transparent background logo.png");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Set the "PreblendImages" property to "true" to preblend transparent images
        // with a background, which may reduce artifacts.
        // Set the "PreblendImages" property to "false" to render transparent images normally.
        options.setPreblendImages(preblendImages);

        doc.save(getArtifactsDir() + "PdfSaveOptions.PreblendImages.pdf", options);
        //ExEnd
    }

    @DataProvider(name = "preblendImagesDataProvider")
    public static Object[][] preblendImagesDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "interpolateImagesDataProvider")
    public void interpolateImages(boolean interpolateImages) throws Exception {
        //ExStart
        //ExFor:PdfSaveOptions.InterpolateImages
        //ExSummary:Shows how to perform interpolation on images while saving a document to PDF.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        BufferedImage img = ImageIO.read(new File(getImageDir() + "Transparent background logo.png"));
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

        if (interpolateImages) {
            TestUtil.fileContainsString("6 0 obj\r\n" +
                            "<</Type /XObject/Subtype /Image/Width 400/Height 400/ColorSpace [/Indexed/DeviceRGB 255 8 0 R]/BitsPerComponent 8/SMask 9 0 R/Interpolate true/Length 10 0 R/Filter /FlateDecode>>",
                    getArtifactsDir() + "PdfSaveOptions.InterpolateImages.pdf");
        } else {
            TestUtil.fileContainsString("6 0 obj\r\n" +
                            "<</Type /XObject/Subtype /Image/Width 400/Height 400/ColorSpace [/Indexed/DeviceRGB 255 8 0 R]/BitsPerComponent 8/SMask 9 0 R/Length 10 0 R/Filter /FlateDecode>>",
                    getArtifactsDir() + "PdfSaveOptions.InterpolateImages.pdf");
        }
    }

    @DataProvider(name = "interpolateImagesDataProvider")
    public static Object[][] interpolateImagesDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
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

        Assert.assertEquals(43, warningCallback.getCount());
    }

    public static class RenderCallback implements IWarningCallback {
        public void warning(WarningInfo info) {
            System.out.println(MessageFormat.format("{0}: {1}.", info.getWarningType(), info.getDescription()));
            mWarnings.add(info);
        }

        public int getCount() {
            return mWarnings.size();
        }

        private final ArrayList<WarningInfo> mWarnings = new ArrayList<>();
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
        Date signingTime = new Date();
        options.setDigitalSignatureDetails(new PdfDigitalSignatureDetails(certificateHolder, "Test Signing", "My Office", signingTime));
        options.getDigitalSignatureDetails().setHashAlgorithm(PdfDigitalSignatureHashAlgorithm.SHA_256);

        Assert.assertEquals(options.getDigitalSignatureDetails().getReason(), "Test Signing");
        Assert.assertEquals(options.getDigitalSignatureDetails().getLocation(), "My Office");
        Assert.assertEquals(options.getDigitalSignatureDetails().getSignatureDate(), signingTime);

        doc.save(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignature.pdf", options);
        //ExEnd

        TestUtil.fileContainsString("6 0 obj\r\n" +
                        "<</Type /Annot/Subtype /Widget/Rect [0 0 0 0]/FT /Sig/DR <<>>/F 132/V 7 0 R/P 4 0 R/T(��\0A\u0000s\u0000p\u0000o\u0000s\0e\0D\u0000i\u0000g\u0000i\u0000t\0a\u0000l\u0000S\u0000i\u0000g\u0000n\0a\u0000t\u0000u\u0000r\0e)/AP <</N 8 0 R>>>>",
                getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignature.pdf");

        Assert.assertFalse(FileFormatUtil.detectFileFormat(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignature.pdf")
                .hasDigitalSignature());

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignature.pdf");

        Assert.assertFalse(pdfDocument.getForm().getSignaturesExist());

        SignatureField signatureField = (SignatureField) pdfDocument.getForm().get(1);

        Assert.assertEquals("AsposeDigitalSignature", signatureField.getFullName());
        Assert.assertEquals("AsposeDigitalSignature", signatureField.getPartialName());
        Assert.assertEquals(com.aspose.pdf.PKCS7.class.getName(), signatureField.getSignature().getClass().getName());
        Assert.assertEquals("þÿ\u0000M\u0000o\u0000r\u0000z\0a\u0000l\u0000.\u0000M\0e", signatureField.getSignature().getAuthority());
        Assert.assertEquals("þÿ\u0000M\u0000y\u0000 \u0000O\0f\0f\u0000i\0c\0e", signatureField.getSignature().getLocation());
        Assert.assertEquals("þÿ\u0000T\0e\u0000s\u0000t\u0000 \u0000S\u0000i\u0000g\u0000n\u0000i\u0000n\u0000g", signatureField.getSignature().getReason());

        pdfDocument.close();
    }

    @Test
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
        //ExSummary:Shows how to sign a saved PDF document digitally and timestamp it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Signed PDF contents.");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Create a digital signature, and assign it to our SaveOptions object to sign the document when we save it to PDF. 
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");
        options.setDigitalSignatureDetails(new PdfDigitalSignatureDetails(certificateHolder, "Test Signing", "Aspose Office", new Date()));

        // Create a timestamp authority-verified timestamp.
        options.getDigitalSignatureDetails().setTimestampSettings(new PdfDigitalSignatureTimestampSettings("https://freetsa.org/tsr", "JohnDoe", "MyPassword"));

        // The default lifespan of the timestamp is 100 seconds.
        Assert.assertEquals(options.getDigitalSignatureDetails().getTimestampSettings().getTimeout(), 100000);

        // We can set our own timeout period via the constructor.
        options.getDigitalSignatureDetails().setTimestampSettings(new PdfDigitalSignatureTimestampSettings("https://freetsa.org/tsr", "JohnDoe", "MyPassword", (long) 1800.0));

        Assert.assertEquals(options.getDigitalSignatureDetails().getTimestampSettings().getTimeout(), 1800);
        Assert.assertEquals(options.getDigitalSignatureDetails().getTimestampSettings().getServerUrl(), "https://freetsa.org/tsr");
        Assert.assertEquals(options.getDigitalSignatureDetails().getTimestampSettings().getUserName(), "JohnDoe");
        Assert.assertEquals(options.getDigitalSignatureDetails().getTimestampSettings().getPassword(), "MyPassword");

        // The "Save" method will apply our signature to the output document at this time.
        doc.save(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf", options);
        //ExEnd

        Assert.assertFalse(FileFormatUtil.detectFileFormat(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf").hasDigitalSignature());
        TestUtil.fileContainsString("6 0 obj\r\n" +
                        "<</Type /Annot/Subtype /Widget/Rect [0 0 0 0]/FT /Sig/DR <<>>/F 132/V 7 0 R/P 4 0 R/T(��\0A\u0000s\u0000p\u0000o\u0000s\0e\0D\u0000i\u0000g\u0000i\u0000t\0a\u0000l\u0000S\u0000i\u0000g\u0000n\0a\u0000t\u0000u\u0000r\0e)/AP <</N 8 0 R>>>>",
                getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf");

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf");

        Assert.assertFalse(pdfDocument.getForm().getSignaturesExist());

        SignatureField signatureField = (SignatureField) pdfDocument.getForm().get(1);

        Assert.assertEquals("AsposeDigitalSignature", signatureField.getFullName());
        Assert.assertEquals("AsposeDigitalSignature", signatureField.getPartialName());
        Assert.assertEquals(com.aspose.pdf.PKCS7.class.getName(), signatureField.getSignature().getClass().getName());
        Assert.assertEquals("þÿ\u0000M\u0000o\u0000r\u0000z\0a\u0000l\u0000.\u0000M\0e", signatureField.getSignature().getAuthority());
        Assert.assertEquals("þÿ\0A\u0000s\u0000p\u0000o\u0000s\0e\u0000 \u0000O\0f\0f\u0000i\0c\0e", signatureField.getSignature().getLocation());
        Assert.assertEquals("þÿ\u0000T\0e\u0000s\u0000t\u0000 \u0000S\u0000i\u0000g\u0000n\u0000i\u0000n\u0000g", signatureField.getSignature().getReason());
        Assert.assertTrue(signatureField.getSignature().getTimestampSettings() == null);

        pdfDocument.close();
    }

    @Test(dataProvider = "renderMetafileDataProvider")
    public void renderMetafile(int renderingMode) throws Exception {
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

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.RenderMetafile.pdf");

        switch (renderingMode) {
            case EmfPlusDualRenderingMode.EMF:
            case EmfPlusDualRenderingMode.EMF_PLUS_WITH_FALLBACK:
                Assert.assertEquals(0, pdfDocument.getPages().get_Item(1).getResources().getImages().size());
                TestUtil.fileContainsString("4 0 obj\r\n" +
                                "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 595.29998779 841.90002441]/Resources<</Font<</FAAAAH 7 0 R/FAAABA 10 0 R/FAAABD 13 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                        getArtifactsDir() + "PdfSaveOptions.RenderMetafile.pdf");
                break;
            case EmfPlusDualRenderingMode.EMF_PLUS:
                Assert.assertEquals(1, pdfDocument.getPages().get_Item(1).getResources().getImages().size());
                TestUtil.fileContainsString("4 0 obj\r\n" +
                                "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 595.29998779 841.90002441]/Resources<</Font<</FAAAAH 7 0 R/FAAABB 11 0 R/FAAABE 14 0 R>>/XObject<</X1 9 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                        getArtifactsDir() + "PdfSaveOptions.RenderMetafile.pdf");
                break;
        }

        pdfDocument.close();
    }

    @DataProvider(name = "renderMetafileDataProvider")
    public static Object[][] renderMetafileDataProvider() {
        return new Object[][]
                {
                        {EmfPlusDualRenderingMode.EMF},
                        {EmfPlusDualRenderingMode.EMF_PLUS},
                        {EmfPlusDualRenderingMode.EMF_PLUS_WITH_FALLBACK},
                };
    }

    @Test
    public void encryptionPermissions() throws Exception {
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

        Assert.assertThrows(InvalidPasswordException.class, () -> new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.EncryptionPermissions.pdf"));

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.EncryptionPermissions.pdf", "password");
        TextFragmentAbsorber textAbsorber = new TextFragmentAbsorber();

        pdfDocument.getPages().get_Item(1).accept(textAbsorber);

        Assert.assertEquals("Hello world!", textAbsorber.getText());

        pdfDocument.close();
    }

    @Test(dataProvider = "setNumeralFormatDataProvider")
    public void setNumeralFormat(/*NumeralFormat*/int numeralFormat) throws Exception {
        //ExStart
        //ExFor:FixedPageSaveOptions.NumeralFormat
        //ExFor:NumeralFormat
        //ExSummary:Shows how to set the numeral format used when saving to PDF.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setLocaleId(1025);
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

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.SetNumeralFormat.pdf");
        TextFragmentAbsorber textAbsorber = new TextFragmentAbsorber();

        pdfDocument.getPages().get_Item(1).accept(textAbsorber);

        switch (numeralFormat) {
            case NumeralFormat.EUROPEAN:
                Assert.assertEquals("1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 50, 100", textAbsorber.getText());
                break;
            case NumeralFormat.ARABIC_INDIC:
                Assert.assertEquals(", ٢, ٣, ٤, ٥, ٦, ٧, ٨, ٩, ١٠, ٥٠, ١١٠٠", textAbsorber.getText());
                break;
            case NumeralFormat.EASTERN_ARABIC_INDIC:
                Assert.assertEquals("۱۰۰ ,۵۰ ,۱۰ ,۹ ,۸ ,۷ ,۶ ,۵ ,۴ ,۳ ,۲ ,۱", textAbsorber.getText());
                break;
        }

        pdfDocument.close();
    }

    @DataProvider(name = "setNumeralFormatDataProvider")
    public static Object[][] setNumeralFormatDataProvider() {
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
    public void exportPageSet() throws Exception {
        //ExStart
        //ExFor:FixedPageSaveOptions.PageSet
        //ExSummary:Shows how to export Odd pages from the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < 5; i++) {
            builder.writeln(MessageFormat.format("Page {0} ({1})", i + 1, (i % 2 == 0 ? "odd" : "even")));
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

        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.ExportPageSet.Even.pdf");
        TextAbsorber textAbsorber = new TextAbsorber();
        pdfDocument.getPages().accept(textAbsorber);

        Assert.assertEquals("Page 2 (even)\r\n" +
                "Page 4 (even)", textAbsorber.getText());

        pdfDocument.close();

        pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.ExportPageSet.Odd.pdf");
        textAbsorber = new TextAbsorber();
        pdfDocument.getPages().accept(textAbsorber);

        Assert.assertEquals("Page 1 (odd)\r\n" +
                "Page 3 (odd)\r\n" +
                "Page 5 (odd)", textAbsorber.getText());

        pdfDocument.close();

        pdfDocument = new com.aspose.pdf.Document(getArtifactsDir() + "PdfSaveOptions.ExportPageSet.All.pdf");
        textAbsorber = new TextAbsorber();
        pdfDocument.getPages().accept(textAbsorber);

        Assert.assertEquals("Page 1 (odd)\r\n" +
                "Page 2 (even)\r\n" +
                "Page 3 (odd)\r\n" +
                "Page 4 (even)\r\n" +
                "Page 5 (odd)", textAbsorber.getText());

        pdfDocument.close();
    }
}
