package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.Font;
import com.aspose.words.Shape;
import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.text.MessageFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.UUID;

public class ExDocumentBuilder extends ApiExampleBase {
    @Test
    public void writeAndFont() throws Exception {
        //ExStart
        //ExFor:Font.Size
        //ExFor:Font.Bold
        //ExFor:Font.Name
        //ExFor:Font.Color
        //ExFor:Font.Underline
        //ExFor:DocumentBuilder.#ctor
        //ExSummary:Inserts formatted text using DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Specify font formatting before adding text
        Font font = builder.getFont();
        font.setSize(16);
        font.setBold(true);
        font.setColor(Color.BLUE);
        font.setName("Courier New");
        font.setUnderline(Underline.DASH);

        builder.write("Hello world!");
        //ExEnd

        doc = DocumentHelper.saveOpen(builder.getDocument());
        Run firstRun = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("Hello world!", firstRun.getText().trim());
        Assert.assertEquals(16.0, firstRun.getFont().getSize());
        Assert.assertTrue(firstRun.getFont().getBold());
        Assert.assertEquals("Courier New", firstRun.getFont().getName());
        Assert.assertEquals(Color.BLUE.getRGB(), firstRun.getFont().getColor().getRGB());
        Assert.assertEquals(Underline.DASH, firstRun.getFont().getUnderline());
    }

    @Test
    public void headersAndFooters() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder
        //ExFor:DocumentBuilder.#ctor(Document)
        //ExFor:DocumentBuilder.MoveToHeaderFooter
        //ExFor:DocumentBuilder.MoveToSection
        //ExFor:DocumentBuilder.InsertBreak
        //ExFor:DocumentBuilder.Writeln
        //ExFor:HeaderFooterType
        //ExFor:PageSetup.DifferentFirstPageHeaderFooter
        //ExFor:PageSetup.OddAndEvenPagesHeaderFooter
        //ExFor:BreakType
        //ExSummary:Shows how to create headers and footers in a document using DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Specify that we want headers and footers different for first, even and odd pages
        builder.getPageSetup().setDifferentFirstPageHeaderFooter(true);
        builder.getPageSetup().setOddAndEvenPagesHeaderFooter(true);

        // Create the headers
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
        builder.write("Header for the first page");
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_EVEN);
        builder.write("Header for even pages");
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.write("Header for all other pages");

        // Create three pages in the document
        builder.moveToSection(0);
        builder.writeln("Page1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page2");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page3");

        doc.save(getArtifactsDir() + "DocumentBuilder.HeadersAndFooters.docx");
        //ExEnd

        HeaderFooterCollection headersFooters =
                new Document(getArtifactsDir() + "DocumentBuilder.HeadersAndFooters.docx").getFirstSection().getHeadersFooters();

        Assert.assertEquals(3, headersFooters.getCount());
        Assert.assertEquals("Header for the first page", headersFooters.getByHeaderFooterType(HeaderFooterType.HEADER_FIRST).getText().trim());
        Assert.assertEquals("Header for even pages", headersFooters.getByHeaderFooterType(HeaderFooterType.HEADER_EVEN).getText().trim());
        Assert.assertEquals("Header for all other pages", headersFooters.getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY).getText().trim());

    }

    @Test
    public void mergeFields() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertField(String)
        //ExFor:DocumentBuilder.MoveToMergeField(String, Boolean, Boolean)
        //ExSummary:Shows how to insert merge fields and move between them.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField("MERGEFIELD MyMergeField1 \\* MERGEFORMAT");
        builder.insertField("MERGEFIELD MyMergeField2 \\* MERGEFORMAT");

        // The second merge field starts immediately after the end of the first
        // We'll move the builder's cursor to the end of the first so we can split them by text
        builder.moveToMergeField("MyMergeField1", true, false);
        Assert.assertEquals(doc.getRange().getFields().get(1).getStart(), builder.getCurrentNode());

        builder.write(" Text between our two merge fields. ");

        doc.save(getArtifactsDir() + "DocumentBuilder.MergeFields.docx");
        //ExEnd		

        doc = new Document(getArtifactsDir() + "DocumentBuilder.MergeFields.docx");

        Assert.assertEquals(2, doc.getRange().getFields().getCount());

        TestUtil.verifyField(FieldType.FIELD_MERGE_FIELD, "MERGEFIELD MyMergeField1 \\* MERGEFORMAT", "«MyMergeField1»", doc.getRange().getFields().get(0));
        TestUtil.verifyField(FieldType.FIELD_MERGE_FIELD, "MERGEFIELD MyMergeField2 \\* MERGEFORMAT", "«MyMergeField2»", doc.getRange().getFields().get(1));
    }

    @Test
    public void insertHorizontalRule() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertHorizontalRule
        //ExFor:ShapeBase.IsHorizontalRule
        //ExFor:Shape.HorizontalRuleFormat
        //ExFor:HorizontalRuleFormat
        //ExFor:HorizontalRuleFormat.Alignment
        //ExFor:HorizontalRuleFormat.WidthPercent
        //ExFor:HorizontalRuleFormat.Height
        //ExFor:HorizontalRuleFormat.Color
        //ExFor:HorizontalRuleFormat.NoShade
        //ExSummary:Shows how to insert horizontal rule shape in a document and customize the formatting.
        // Use a document builder to insert a horizontal rule
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.insertHorizontalRule();

        HorizontalRuleFormat horizontalRuleFormat = shape.getHorizontalRuleFormat();
        horizontalRuleFormat.setAlignment(HorizontalRuleAlignment.CENTER);
        horizontalRuleFormat.setWidthPercent(70.0);
        horizontalRuleFormat.setHeight(3.0);
        horizontalRuleFormat.setColor(Color.BLUE);
        horizontalRuleFormat.setNoShade(true);

        Assert.assertTrue(shape.isHorizontalRule());
        Assert.assertTrue(shape.getHorizontalRuleFormat().getNoShade());
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(HorizontalRuleAlignment.CENTER, shape.getHorizontalRuleFormat().getAlignment());
        Assert.assertEquals(70.0, shape.getHorizontalRuleFormat().getWidthPercent());
        Assert.assertEquals(3.0, shape.getHorizontalRuleFormat().getHeight());
        Assert.assertEquals(Color.BLUE.getRGB(), shape.getHorizontalRuleFormat().getColor().getRGB());
    }

    @Test(description = "Checking the boundary conditions of WidthPercent and Height properties")
    public void horizontalRuleFormatExceptions() throws Exception {
        DocumentBuilder builder = new DocumentBuilder();
        Shape shape = builder.insertHorizontalRule();

        HorizontalRuleFormat horizontalRuleFormat = shape.getHorizontalRuleFormat();
        horizontalRuleFormat.setWidthPercent(1.0);
        horizontalRuleFormat.setWidthPercent(100.0);
        Assert.assertThrows(IllegalArgumentException.class, () -> horizontalRuleFormat.setWidthPercent(0.0));
        Assert.assertThrows(IllegalArgumentException.class, () -> horizontalRuleFormat.setWidthPercent(101.0));

        horizontalRuleFormat.setHeight(0.0);
        horizontalRuleFormat.setHeight(1584.0);
        Assert.assertThrows(IllegalArgumentException.class, () -> horizontalRuleFormat.setHeight(-1));
        Assert.assertThrows(IllegalArgumentException.class, () -> horizontalRuleFormat.setHeight(1585.0));
    }

    @Test
    public void insertHyperlink() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertHyperlink
        //ExFor:Font.ClearFormatting
        //ExFor:Font.Color
        //ExFor:Font.Underline
        //ExFor:Underline
        //ExSummary:Shows how to insert a hyperlink into a document using DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Please make sure to visit ");

        // Specify font formatting for the hyperlink
        builder.getFont().setColor(Color.BLUE);
        builder.getFont().setUnderline(Underline.SINGLE);

        // Insert the link
        builder.insertHyperlink("Aspose Website", "http://www.aspose.com", false);

        // Revert to default formatting
        builder.getFont().clearFormatting();
        builder.write(" for more information.");

        // Holding Ctrl and left clicking on the field in Microsoft Word will take you to the link's address in a web browser
        doc.save(getArtifactsDir() + "DocumentBuilder.InsertHyperlink.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertHyperlink.docx");

        FieldHyperlink hyperlink = (FieldHyperlink) doc.getRange().getFields().get(0);

        Run fieldContents = (Run) hyperlink.getStart().getNextSibling();

        Assert.assertEquals(Color.BLUE.getRGB(), fieldContents.getFont().getColor().getRGB());
        Assert.assertEquals(Underline.SINGLE, fieldContents.getFont().getUnderline());
        Assert.assertEquals("HYPERLINK \"http://www.aspose.com\"", fieldContents.getText().trim());
    }

    @Test
    public void pushPopFont() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.PushFont
        //ExFor:DocumentBuilder.PopFont
        //ExFor:DocumentBuilder.InsertHyperlink
        //ExSummary:Shows how to use temporarily save and restore character formatting when building a document with DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set up font formatting and write text that goes before the hyperlink
        builder.getFont().setName("Arial");
        builder.getFont().setSize(24.0);
        builder.getFont().setBold(true);
        builder.write("To visit Google, hold Ctrl and click ");

        // Save the font formatting so we use different formatting for hyperlink and restore old formatting later
        builder.pushFont();

        // Set new font formatting for the hyperlink and insert the hyperlink
        // The "Hyperlink" style is a Microsoft Word built-in style so we don't have to worry to 
        // create it, it will be created automatically if it does not yet exist in the document
        builder.getFont().setStyleIdentifier(StyleIdentifier.HYPERLINK);
        builder.insertHyperlink("here", "http://www.google.com", false);

        // Restore the formatting that was before the hyperlink
        builder.popFont();

        builder.write(". We hope you enjoyed the example.");

        doc.save(getArtifactsDir() + "DocumentBuilder.PushPopFont.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.PushPopFont.docx");
        RunCollection runs = doc.getFirstSection().getBody().getFirstParagraph().getRuns();

        Assert.assertEquals(4, runs.getCount());

        Assert.assertEquals("To visit Google, hold Ctrl and click", runs.get(0).getText().trim());
        Assert.assertEquals(". We hope you enjoyed the example.", runs.get(3).getText().trim());
        Assert.assertEquals(runs.get(0).getFont().getColor(), runs.get(3).getFont().getColor());
        Assert.assertEquals(runs.get(0).getFont().getUnderline(), runs.get(3).getFont().getUnderline());

        Assert.assertEquals("here", runs.get(2).getText().trim());
        Assert.assertEquals(Color.BLUE.getRGB(), runs.get(2).getFont().getColor().getRGB());
        Assert.assertEquals(Underline.SINGLE, runs.get(2).getFont().getUnderline());
        Assert.assertNotEquals(runs.get(0).getFont().getColor(), runs.get(2).getFont().getColor());
        Assert.assertNotEquals(runs.get(0).getFont().getUnderline(), runs.get(2).getFont().getUnderline());
    }

    @Test
    public void insertWatermark() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.MoveToHeaderFooter
        //ExFor:PageSetup.PageWidth
        //ExFor:PageSetup.PageHeight
        //ExFor:DocumentBuilder.InsertImage(Image)
        //ExFor:WrapType
        //ExFor:RelativeHorizontalPosition
        //ExFor:RelativeVerticalPosition
        //ExSummary:Shows how to a watermark image into a document using DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The best place for the watermark image is in the header or footer so it is shown on every page
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);

        BufferedImage image = ImageIO.read(new File(getImageDir() + "Transparent background logo.png"));

        // Insert a floating picture
        Shape shape = builder.insertImage(image);
        shape.setWrapType(WrapType.NONE);
        shape.setBehindText(true);

        shape.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        shape.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);

        // Calculate image left and top position so it appears in the center of the page
        shape.setLeft((builder.getPageSetup().getPageWidth() - shape.getWidth()) / 2.0);
        shape.setTop((builder.getPageSetup().getPageHeight() - shape.getHeight()) / 2.0);

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertWatermark.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertWatermark.docx");
        shape = (Shape) doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY).getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.PNG, shape);
        Assert.assertEquals(WrapType.NONE, shape.getWrapType());
        Assert.assertTrue(shape.getBehindText());
        Assert.assertEquals(RelativeHorizontalPosition.PAGE, shape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.PAGE, shape.getRelativeVerticalPosition());
        Assert.assertEquals((doc.getFirstSection().getPageSetup().getPageWidth() - shape.getWidth()) / 2.0, shape.getLeft());
        Assert.assertEquals((doc.getFirstSection().getPageSetup().getPageHeight() - shape.getHeight()) / 2.0, shape.getTop());
    }

    @Test
    public void insertOleObject() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertOleObject(String, Boolean, Boolean, Image)
        //ExFor:DocumentBuilder.InsertOleObject(String, String, Boolean, Boolean, Image)
        //ExFor:DocumentBuilder.InsertOleObjectAsIcon(String, Boolean, String, String)
        //ExSummary:Shows how to insert an OLE object into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert ole object
        BufferedImage representingImage = ImageIO.read(new File(getImageDir() + "Logo.jpg"));
        builder.insertOleObject(getMyDir() + "Spreadsheet.xlsx", false, false, representingImage);

        // Insert ole object with ProgId
        builder.insertOleObject(getMyDir() + "Spreadsheet.xlsx", "Excel.Sheet", false, true, null);

        // Insert ole object as Icon
        // There is one limitation for now: the maximum size of the icon must be 32x32 for the correct display
        builder.insertOleObjectAsIcon(getMyDir() + "Presentation.pptx", false, getImageDir() + "Logo icon.ico",
                "Caption (can not be null)");

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertOleObject.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertOleObject.docx");
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(ShapeType.OLE_OBJECT, shape.getShapeType());
        Assert.assertEquals("Excel.Sheet.12", shape.getOleFormat().getProgId());
        Assert.assertEquals(".xlsx", shape.getOleFormat().getSuggestedExtension());

        shape = (Shape) doc.getChild(NodeType.SHAPE, 1, true);

        Assert.assertEquals(ShapeType.OLE_OBJECT, shape.getShapeType());
        Assert.assertEquals("Package", shape.getOleFormat().getProgId());
        Assert.assertEquals(".xlsx", shape.getOleFormat().getSuggestedExtension());

        shape = (Shape) doc.getChild(NodeType.SHAPE, 2, true);

        Assert.assertEquals(ShapeType.OLE_OBJECT, shape.getShapeType());
        Assert.assertEquals("PowerPoint.Show.12", shape.getOleFormat().getProgId());
        Assert.assertEquals(".pptx", shape.getOleFormat().getSuggestedExtension());
    }

    @Test
    public void insertHtml() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertHtml(String)
        //ExSummary:Shows how to insert Html content into a document using a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        final String HTML = "<P align='right'>Paragraph right</P>" + "<b>Implicit paragraph left</b>" +
                "<div align='center'>Div center</div>" + "<h1 align='left'>Heading 1 left.</h1>";

        builder.insertHtml(HTML);

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertHtml.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertHtml.docx");
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        Assert.assertEquals("Paragraph right", paragraphs.get(0).getText().trim());
        Assert.assertEquals(ParagraphAlignment.RIGHT, paragraphs.get(0).getParagraphFormat().getAlignment());

        Assert.assertEquals("Implicit paragraph left", paragraphs.get(1).getText().trim());
        Assert.assertEquals(ParagraphAlignment.LEFT, paragraphs.get(1).getParagraphFormat().getAlignment());
        Assert.assertTrue(paragraphs.get(1).getRuns().get(0).getFont().getBold());

        Assert.assertEquals("Div center", paragraphs.get(2).getText().trim());
        Assert.assertEquals(ParagraphAlignment.CENTER, paragraphs.get(2).getParagraphFormat().getAlignment());

        Assert.assertEquals("Heading 1 left.", paragraphs.get(3).getText().trim());
        Assert.assertEquals("Heading 1", paragraphs.get(3).getParagraphFormat().getStyle().getName());
    }

    @Test
    public void insertHtmlWithFormatting() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertHtml(String, Boolean)
        //ExSummary:Shows how to insert Html content into a document using a builder while applying the builder's formatting. 
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the builder's text alignment
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.DISTRIBUTED);

        // If we insert text while setting useBuilderFormatting to true, any formatting applied to the builder will be applied to inserted .html content
        // However, if the html text has formatting coded into it, that formatting takes precedence over the builder's formatting
        // In this case, elements with "align" attributes do not get affected by the ParagraphAlignment we specified above
        builder.insertHtml(
                "<P align='right'>Paragraph right</P>" + "<b>Implicit paragraph left</b>" +
                        "<div align='center'>Div center</div>" + "<h1 align='left'>Heading 1 left.</h1>", true);

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertHtmlWithFormatting.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertHtmlWithFormatting.docx");
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        Assert.assertEquals("Paragraph right", paragraphs.get(0).getText().trim());
        Assert.assertEquals(ParagraphAlignment.RIGHT, paragraphs.get(0).getParagraphFormat().getAlignment());

        Assert.assertEquals("Implicit paragraph left", paragraphs.get(1).getText().trim());
        Assert.assertEquals(ParagraphAlignment.DISTRIBUTED, paragraphs.get(1).getParagraphFormat().getAlignment());
        Assert.assertTrue(paragraphs.get(1).getRuns().get(0).getFont().getBold());

        Assert.assertEquals("Div center", paragraphs.get(2).getText().trim());
        Assert.assertEquals(ParagraphAlignment.CENTER, paragraphs.get(2).getParagraphFormat().getAlignment());

        Assert.assertEquals("Heading 1 left.", paragraphs.get(3).getText().trim());
        Assert.assertEquals("Heading 1", paragraphs.get(3).getParagraphFormat().getStyle().getName());
    }

    @Test
    public void mathML() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        final String mathML =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><msub><mi>a</mi><mrow><mn>1</mn></mrow></msub><mo>+</mo><msub><mi>b</mi><mrow><mn>1</mn></mrow></msub></mrow></math>";

        builder.insertHtml(mathML);

        doc.save(getArtifactsDir() + "DocumentBuilder.MathML.docx");
        doc.save(getArtifactsDir() + "DocumentBuilder.MathML.pdf");

        Assert.assertTrue(DocumentHelper.compareDocs(getGoldsDir() + "DocumentBuilder.MathML Gold.docx", getArtifactsDir() + "DocumentBuilder.MathML.docx"));
    }

    @Test
    public void insertTextAndBookmark() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.StartBookmark
        //ExFor:DocumentBuilder.EndBookmark
        //ExSummary:Shows how to add some text into the document and encloses the text in a bookmark using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder();

        builder.startBookmark("MyBookmark");
        builder.writeln("Text inside a bookmark.");
        builder.endBookmark("MyBookmark");
        //ExEnd

        Document doc = DocumentHelper.saveOpen(builder.getDocument());

        Assert.assertEquals(1, doc.getRange().getBookmarks().getCount());
        Assert.assertEquals("MyBookmark", doc.getRange().getBookmarks().get(0).getName());
        Assert.assertEquals("Text inside a bookmark.", doc.getRange().getBookmarks().get(0).getText().trim());
    }

    @Test
    public void createForm() throws Exception {
        //ExStart
        //ExFor:TextFormFieldType
        //ExFor:DocumentBuilder.InsertTextInput
        //ExFor:DocumentBuilder.InsertComboBox
        //ExSummary:Shows how to build a form field.
        DocumentBuilder builder = new DocumentBuilder();

        // Insert a text form field for input a name
        builder.insertTextInput("", TextFormFieldType.REGULAR, "", "Enter your name here", 30);

        // Insert two blank lines
        builder.writeln("");
        builder.writeln("");

        String[] items = new String[]{"-- Select your favorite footwear --", "Sneakers", "Oxfords", "Flip-flops", "Other", "I prefer to be barefoot"};

        // Insert a combo box to select a footwear type
        builder.insertComboBox("", items, 0);

        // Insert 2 blank lines
        builder.writeln("");
        builder.writeln("");

        builder.getDocument().save(getArtifactsDir() + "DocumentBuilder.CreateForm.docx");
        //ExEnd

        Document doc = new Document(getArtifactsDir() + "DocumentBuilder.CreateForm.docx");
        FormField formField = doc.getRange().getFormFields().get(0);

        Assert.assertEquals(TextFormFieldType.REGULAR, formField.getTextInputType());
        Assert.assertEquals("Enter your name here", formField.getResult());

        formField = doc.getRange().getFormFields().get(1);

        Assert.assertEquals(TextFormFieldType.REGULAR, formField.getTextInputType());
        Assert.assertEquals("-- Select your favorite footwear --", formField.getResult());
        Assert.assertEquals(0, formField.getDropDownSelectedIndex());
        Assert.assertEquals(Arrays.asList(new String[]{"-- Select your favorite footwear --", "Sneakers", "Oxfords", "Flip-flops", "Other",
                "I prefer to be barefoot"}), formField.getDropDownItems());
    }

    @Test
    public void insertCheckBox() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertCheckBox(string, bool, bool, int)
        //ExFor:DocumentBuilder.InsertCheckBox(String, bool, int)
        //ExSummary:Shows how to insert checkboxes to the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertCheckBox("", false, false, 0);
        builder.insertCheckBox("CheckBox_Default", true, true, 50);
        builder.insertCheckBox("CheckBox_OnlyCheckedValue", true, 100);
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        // Get checkboxes from the document
        FormFieldCollection formFields = doc.getRange().getFormFields();

        // Check that is the right checkbox
        Assert.assertEquals(formFields.get(0).getName(), "");

        // Assert that parameters sets correctly
        Assert.assertEquals(formFields.get(0).getChecked(), false);
        Assert.assertEquals(formFields.get(0).getDefault(), false);
        Assert.assertEquals(formFields.get(0).getCheckBoxSize(), 10.0);

        // Check that is the right checkbox
        // Please pay attention that MS Word allows strings with at most 20 characters
        Assert.assertEquals(formFields.get(1).getName(), "CheckBox_Default");

        // Assert that parameters sets correctly
        Assert.assertEquals(true, formFields.get(1).getChecked());
        Assert.assertEquals(true, formFields.get(1).getDefault());
        Assert.assertEquals(50.0, formFields.get(1).getCheckBoxSize());

        // Check that is the right checkbox
        // Please pay attention that MS Word allows strings with at most 20 characters
        Assert.assertEquals(formFields.get(2).getName(), "CheckBox_OnlyChecked");

        // Assert that parameters sets correctly
        Assert.assertEquals(formFields.get(2).getChecked(), true);
        Assert.assertEquals(formFields.get(2).getDefault(), true);
        Assert.assertEquals(formFields.get(2).getCheckBoxSize(), 100.0);
    }

    @Test
    public void insertCheckBoxEmptyName() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Checking that the checkbox insertion with an empty name working correctly
        builder.insertCheckBox("", true, false, 1);
        builder.insertCheckBox("", false, 1);
    }

    @Test
    public void workingWithNodes() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.MoveTo(Node)
        //ExFor:DocumentBuilder.MoveToBookmark(String)
        //ExFor:DocumentBuilder.CurrentParagraph
        //ExFor:DocumentBuilder.CurrentNode
        //ExFor:DocumentBuilder.MoveToDocumentStart
        //ExFor:DocumentBuilder.MoveToDocumentEnd
        //ExFor:DocumentBuilder.IsAtEndOfParagraph
        //ExFor:DocumentBuilder.IsAtStartOfParagraph
        //ExSummary:Shows how to move a DocumentBuilder to different nodes in a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a bookmark and add content to it using a DocumentBuilder
        builder.startBookmark("MyBookmark");
        builder.writeln("Bookmark contents.");
        builder.endBookmark("MyBookmark");

        // The node that the DocumentBuilder is currently at is past the boundaries of the bookmark  
        Assert.assertEquals(doc.getRange().getBookmarks().get(0).getBookmarkEnd(), builder.getCurrentParagraph().getFirstChild());

        // If we wish to revise the content of our bookmark with the DocumentBuilder, we can move back to it like this
        builder.moveToBookmark("MyBookmark");

        // Now we're located between the bookmark's BookmarkStart and BookmarkEnd nodes, so any text the builder adds will be within it
        Assert.assertEquals(doc.getRange().getBookmarks().get(0).getBookmarkStart(), builder.getCurrentParagraph().getFirstChild());

        // We can move the builder to an individual node,
        // which in this case will be the first node of the first paragraph, like this
        builder.moveTo(doc.getFirstSection().getBody().getFirstParagraph().getChildNodes(NodeType.ANY, false).get(0));

        Assert.assertEquals(NodeType.BOOKMARK_START, builder.getCurrentNode().getNodeType());
        Assert.assertTrue(builder.isAtStartOfParagraph());

        // A shorter way of moving the very start/end of a document is with these methods
        builder.moveToDocumentEnd();

        Assert.assertTrue(builder.isAtEndOfParagraph());

        builder.moveToDocumentStart();

        Assert.assertTrue(builder.isAtStartOfParagraph());
        //ExEnd
    }

    @Test
    public void fillMergeFields() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.MoveToMergeField(String)
        //ExFor:DocumentBuilder.Bold
        //ExFor:DocumentBuilder.Italic
        //ExSummary:Shows how to fill MERGEFIELDs with data with a DocumentBuilder and without a mail merge.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some MERGEFIELDS, which accept data from columns of the same name in a data source during a mail merge
        builder.insertField(" MERGEFIELD Chairman ");
        builder.insertField(" MERGEFIELD ChiefFinancialOfficer ");
        builder.insertField(" MERGEFIELD ChiefTechnologyOfficer ");

        // They can also be filled in manually like this
        builder.moveToMergeField("Chairman");
        builder.setBold(true);
        builder.writeln("John Doe");

        builder.moveToMergeField("ChiefFinancialOfficer");
        builder.setItalic(true);
        builder.writeln("Jane Doe");

        builder.moveToMergeField("ChiefTechnologyOfficer");
        builder.setItalic(true);
        builder.writeln("John Bloggs");

        doc.save(getArtifactsDir() + "DocumentBuilder.FillMergeFields.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.FillMergeFields.docx");
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        Assert.assertTrue(paragraphs.get(0).getRuns().get(0).getFont().getBold());
        Assert.assertEquals("John Doe", paragraphs.get(0).getRuns().get(0).getText().trim());

        Assert.assertTrue(paragraphs.get(1).getRuns().get(0).getFont().getItalic());
        Assert.assertEquals("Jane Doe", paragraphs.get(1).getRuns().get(0).getText().trim());

        Assert.assertTrue(paragraphs.get(2).getRuns().get(0).getFont().getItalic());
        Assert.assertEquals("John Bloggs", paragraphs.get(2).getRuns().get(0).getText().trim());

    }

    @Test
    public void insertToc() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertTableOfContents
        //ExFor:Document.UpdateFields
        //ExFor:DocumentBuilder.#ctor(Document)
        //ExFor:ParagraphFormat.StyleIdentifier
        //ExFor:DocumentBuilder.InsertBreak
        //ExFor:BreakType
        //ExSummary:Shows how to insert a Table of contents (TOC) into a document using heading styles as entries.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table of contents at the beginning of the document,
        // and set it to pick up paragraphs with headings of levels 1 to 3 and entries to act like hyperlinks
        builder.insertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Start the actual document content on the second page
        builder.insertBreak(BreakType.PAGE_BREAK);

        // Build a document with complex structure by applying different heading styles thus creating TOC entries
        // The heading levels we use below will affect the list levels in which these items will appear in the TOC,
        // and only levels 1-3 will be picked up by our TOC due to its settings
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.writeln("Heading 1");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);
        builder.writeln("Heading 1.1");
        builder.writeln("Heading 1.2");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.writeln("Heading 2");
        builder.writeln("Heading 3");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);
        builder.writeln("Heading 3.1");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_3);
        builder.writeln("Heading 3.1.1");
        builder.writeln("Heading 3.1.2");
        builder.writeln("Heading 3.1.3");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);
        builder.writeln("Heading 3.2");
        builder.writeln("Heading 3.3");

        // Call the method below to update the TOC and save
        doc.updateFields();
        doc.save(getArtifactsDir() + "DocumentBuilder.InsertToc.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertToc.docx");
        FieldToc tableOfContents = (FieldToc) doc.getRange().getFields().get(0);

        Assert.assertEquals("1-3", tableOfContents.getHeadingLevelRange());
        Assert.assertTrue(tableOfContents.getInsertHyperlinks());
        Assert.assertTrue(tableOfContents.getHideInWebLayout());
        Assert.assertTrue(tableOfContents.getUseParagraphOutlineLevel());
    }

    @Test
    public void insertTable() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder
        //ExFor:DocumentBuilder.Write
        //ExFor:DocumentBuilder.StartTable
        //ExFor:DocumentBuilder.InsertCell
        //ExFor:DocumentBuilder.EndRow
        //ExFor:DocumentBuilder.EndTable
        //ExFor:DocumentBuilder.CellFormat
        //ExFor:DocumentBuilder.RowFormat
        //ExFor:CellFormat
        //ExFor:CellFormat.FitText
        //ExFor:CellFormat.Width
        //ExFor:CellFormat.VerticalAlignment
        //ExFor:CellFormat.Shading
        //ExFor:CellFormat.Orientation
        //ExFor:CellFormat.WrapText
        //ExFor:RowFormat
        //ExFor:RowFormat.Borders
        //ExFor:RowFormat.ClearFormatting
        //ExFor:Shading.ClearFormatting
        //ExSummary:Shows how to build a nice bordered table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table
        builder.startTable();

        // Set the appropriate paragraph, cell, and row formatting. The formatting properties are preserved
        // until they are explicitly modified so there's no need to set them for each row or cell
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);

        builder.getCellFormat().clearFormatting();
        builder.getCellFormat().setWidth(150.0);
        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.GREEN);
        builder.getCellFormat().setWrapText(false);
        builder.getCellFormat().setFitText(true);

        builder.getRowFormat().clearFormatting();
        builder.getRowFormat().setHeightRule(HeightRule.EXACTLY);
        builder.getRowFormat().setHeight(50.0);
        builder.getRowFormat().getBorders().setLineStyle(LineStyle.ENGRAVE_3_D);
        builder.getRowFormat().getBorders().setColor(Color.ORANGE);

        builder.insertCell();
        builder.write("Row 1, Col 1");

        builder.insertCell();
        builder.write("Row 1, Col 2");

        builder.endRow();

        // Remove the shading (clear background)
        builder.getCellFormat().getShading().clearFormatting();

        builder.insertCell();
        builder.write("Row 2, Col 1");

        builder.insertCell();
        builder.write("Row 2, Col 2");

        builder.endRow();

        builder.insertCell();

        // Make the row height bigger so that a vertically oriented text could fit into cells
        builder.getRowFormat().setHeight(150.0);
        builder.getCellFormat().setOrientation(TextOrientation.UPWARD);
        builder.write("Row 3, Col 1");

        builder.insertCell();
        builder.getCellFormat().setOrientation(TextOrientation.DOWNWARD);
        builder.write("Row 3, Col 2");

        builder.endRow();

        builder.endTable();

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertTable.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertTable.docx");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals("Row 1, Col 1", table.getRows().get(0).getCells().get(0).getText().trim());
        Assert.assertEquals("Row 1, Col 2", table.getRows().get(0).getCells().get(1).getText().trim());
        Assert.assertEquals(HeightRule.EXACTLY, table.getRows().get(0).getRowFormat().getHeightRule());
        Assert.assertEquals(50.0d, table.getRows().get(0).getRowFormat().getHeight());
        Assert.assertEquals(LineStyle.ENGRAVE_3_D, table.getRows().get(0).getRowFormat().getBorders().getLineStyle());
        Assert.assertEquals(Color.ORANGE.getRGB(), table.getRows().get(0).getRowFormat().getBorders().getColor().getRGB());

        for (Cell c : (Iterable<Cell>) table.getRows().get(0).getCells()) {
            Assert.assertEquals(150.0, c.getCellFormat().getWidth());
            Assert.assertEquals(CellVerticalAlignment.CENTER, c.getCellFormat().getVerticalAlignment());
            Assert.assertEquals(Color.GREEN.getRGB(), c.getCellFormat().getShading().getBackgroundPatternColor().getRGB());
            Assert.assertFalse(c.getCellFormat().getWrapText());
            Assert.assertTrue(c.getCellFormat().getFitText());

            Assert.assertEquals(ParagraphAlignment.CENTER, c.getFirstParagraph().getParagraphFormat().getAlignment());
        }

        Assert.assertEquals("Row 2, Col 1", table.getRows().get(1).getCells().get(0).getText().trim());
        Assert.assertEquals("Row 2, Col 2", table.getRows().get(1).getCells().get(1).getText().trim());


        for (Cell c : (Iterable<Cell>) table.getRows().get(1).getCells()) {
            Assert.assertEquals(150.0, c.getCellFormat().getWidth());
            Assert.assertEquals(CellVerticalAlignment.CENTER, c.getCellFormat().getVerticalAlignment());
            Assert.assertEquals(0, c.getCellFormat().getShading().getBackgroundPatternColor().getRGB());
            Assert.assertFalse(c.getCellFormat().getWrapText());
            Assert.assertTrue(c.getCellFormat().getFitText());

            Assert.assertEquals(ParagraphAlignment.CENTER, c.getFirstParagraph().getParagraphFormat().getAlignment());
        }

        Assert.assertEquals(150.0, table.getRows().get(2).getRowFormat().getHeight());

        Assert.assertEquals("Row 3, Col 1", table.getRows().get(2).getCells().get(0).getText().trim());
        Assert.assertEquals(TextOrientation.UPWARD, table.getRows().get(2).getCells().get(0).getCellFormat().getOrientation());
        Assert.assertEquals(ParagraphAlignment.CENTER, table.getRows().get(2).getCells().get(0).getFirstParagraph().getParagraphFormat().getAlignment());

        Assert.assertEquals("Row 3, Col 2", table.getRows().get(2).getCells().get(1).getText().trim());
        Assert.assertEquals(TextOrientation.DOWNWARD, table.getRows().get(2).getCells().get(1).getCellFormat().getOrientation());
        Assert.assertEquals(ParagraphAlignment.CENTER, table.getRows().get(2).getCells().get(1).getFirstParagraph().getParagraphFormat().getAlignment());
    }

    @Test
    public void insertTableWithStyle() throws Exception {
        //ExStart
        //ExFor:Table.StyleIdentifier
        //ExFor:Table.StyleOptions
        //ExFor:TableStyleOptions
        //ExFor:Table.AutoFit
        //ExFor:AutoFitBehavior
        //ExSummary:Shows how to build a new table with a table style applied.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.startTable();

        // We must insert at least one row first before setting any table formatting
        builder.insertCell();

        // Set the table style used based of the unique style identifier
        // Note that not all table styles are available when saving as .doc format
        table.setStyleIdentifier(StyleIdentifier.MEDIUM_SHADING_1_ACCENT_1);

        // Apply which features should be formatted by the style
        table.setStyleOptions(TableStyleOptions.FIRST_COLUMN | TableStyleOptions.ROW_BANDS | TableStyleOptions.FIRST_ROW);
        table.autoFit(AutoFitBehavior.AUTO_FIT_TO_CONTENTS);

        // Continue with building the table as normal
        builder.writeln("Item");
        builder.getCellFormat().setRightPadding(40.0);
        builder.insertCell();
        builder.writeln("Quantity (kg)");
        builder.endRow();

        builder.insertCell();
        builder.writeln("Apples");
        builder.insertCell();
        builder.writeln("20");
        builder.endRow();

        builder.insertCell();
        builder.writeln("Bananas");
        builder.insertCell();
        builder.writeln("40");
        builder.endRow();

        builder.insertCell();
        builder.writeln("Carrots");
        builder.insertCell();
        builder.writeln("50");
        builder.endRow();

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertTableWithStyle.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertTableWithStyle.docx");

        // Verify that the style was set by expanding to direct formatting
        doc.expandTableStylesToDirectFormatting();

        Assert.assertEquals("Medium Shading 1 Accent 1", table.getStyle().getName());
        Assert.assertEquals(TableStyleOptions.FIRST_COLUMN | TableStyleOptions.ROW_BANDS | TableStyleOptions.FIRST_ROW,
                table.getStyleOptions());
        Assert.assertEquals(189, (table.getFirstRow().getFirstCell().getCellFormat().getShading().getBackgroundPatternColor().getBlue() & 0xFF));
        Assert.assertEquals(Color.WHITE.getRGB(), table.getFirstRow().getFirstCell().getFirstParagraph().getRuns().get(0).getFont().getColor().getRGB());
        Assert.assertNotEquals(Color.BLUE.getRGB(),
                (table.getLastRow().getFirstCell().getCellFormat().getShading().getBackgroundPatternColor().getBlue() & 0xFF));
        Assert.assertEquals(0, table.getLastRow().getFirstCell().getFirstParagraph().getRuns().get(0).getFont().getColor().getRGB());
    }

    @Test
    public void insertTableSetHeadingRow() throws Exception {
        //ExStart
        //ExFor:RowFormat.HeadingFormat
        //ExSummary:Shows how to build a table which include heading rows that repeat on subsequent pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startTable();
        builder.getRowFormat().setHeadingFormat(true);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.getCellFormat().setWidth(100.0);
        builder.insertCell();
        builder.writeln("Heading row 1");
        builder.endRow();
        builder.insertCell();
        builder.writeln("Heading row 2");
        builder.endRow();

        builder.getCellFormat().setWidth(50.0);
        builder.getParagraphFormat().clearFormatting();

        // Insert some content so the table is long enough to continue onto the next page
        for (int i = 0; i < 50; i++) {
            builder.insertCell();
            builder.getRowFormat().setHeadingFormat(false);
            builder.write("Column 1 Text");
            builder.insertCell();
            builder.write("Column 2 Text");
            builder.endRow();
        }

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertTableSetHeadingRow.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertTableSetHeadingRow.docx");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertTrue(table.getFirstRow().getRowFormat().getHeadingFormat());
        Assert.assertTrue(table.getRows().get(1).getRowFormat().getHeadingFormat());
        Assert.assertFalse(table.getRows().get(2).getRowFormat().getHeadingFormat());
    }

    @Test
    public void insertTableWithPreferredWidth() throws Exception {
        //ExStart
        //ExFor:Table.PreferredWidth
        //ExFor:PreferredWidth.FromPercent
        //ExFor:PreferredWidth
        //ExSummary:Shows how to set a table to auto fit to 50% of the page width.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with a width that takes up half the page width
        Table table = builder.startTable();

        // Insert a few cells
        builder.insertCell();
        table.setPreferredWidth(PreferredWidth.fromPercent(50.0));
        builder.writeln("Cell #1");

        builder.insertCell();
        builder.writeln("Cell #2");

        builder.insertCell();
        builder.writeln("Cell #3");

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertTableWithPreferredWidth.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertTableWithPreferredWidth.docx");
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(PreferredWidthType.PERCENT, table.getPreferredWidth().getType());
        Assert.assertEquals(50.0, table.getPreferredWidth().getValue());
    }

    @Test
    public void insertCellsWithPreferredWidths() throws Exception {
        //ExStart
        //ExFor:CellFormat.PreferredWidth
        //ExFor:PreferredWidth
        //ExFor:PreferredWidth.Auto
        //ExFor:PreferredWidth.Equals(PreferredWidth)
        //ExFor:PreferredWidth.Equals(System.Object)
        //ExFor:PreferredWidth.FromPoints
        //ExFor:PreferredWidth.FromPercent
        //ExFor:PreferredWidth.GetHashCode
        //ExFor:PreferredWidth.ToString
        //ExSummary:Shows how to set the different preferred width settings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table row made up of three cells which have different preferred widths
        Table table = builder.startTable();

        // Insert an absolute sized cell
        builder.insertCell();
        builder.getCellFormat().setPreferredWidth(PreferredWidth.fromPoints(40));
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.RED);
        builder.writeln("Cell at 40 points width");

        PreferredWidth width = builder.getCellFormat().getPreferredWidth();
        System.out.println(MessageFormat.format("Width \"{0}\": {1}", width.hashCode(), width.toString()));

        // Insert a relative (percent) sized cell
        builder.insertCell();
        builder.getCellFormat().setPreferredWidth(PreferredWidth.fromPercent(20));
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.BLUE);
        builder.writeln("Cell at 20% width");

        // Each cell had its own PreferredWidth
        Assert.assertFalse(builder.getCellFormat().getPreferredWidth().equals(width));

        width = builder.getCellFormat().getPreferredWidth();
        System.out.println(MessageFormat.format("Width \"{0}\": {1}", width.hashCode(), width.toString()));

        // Insert a auto sized cell
        builder.insertCell();
        builder.getCellFormat().setPreferredWidth(PreferredWidth.AUTO);
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.GREEN);
        builder.writeln("Cell automatically sized. The size of this cell is calculated from the table preferred width.");
        builder.writeln("In this case the cell will fill up the rest of the available space.");

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertCellsWithPreferredWidths.docx");
        //ExEnd

        Assert.assertEquals(100.0d, PreferredWidth.fromPercent(100.0).getValue());
        Assert.assertEquals(100.0d, PreferredWidth.fromPoints(100.0).getValue());

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertCellsWithPreferredWidths.docx");
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(PreferredWidthType.POINTS, table.getFirstRow().getCells().get(0).getCellFormat().getPreferredWidth().getType());
        Assert.assertEquals(40.0d, table.getFirstRow().getCells().get(0).getCellFormat().getPreferredWidth().getValue());

        Assert.assertEquals(PreferredWidthType.PERCENT, table.getFirstRow().getCells().get(1).getCellFormat().getPreferredWidth().getType());
        Assert.assertEquals(20.0d, table.getFirstRow().getCells().get(1).getCellFormat().getPreferredWidth().getValue());

        Assert.assertEquals(PreferredWidthType.AUTO, table.getFirstRow().getCells().get(2).getCellFormat().getPreferredWidth().getType());
        Assert.assertEquals(0.0d, table.getFirstRow().getCells().get(2).getCellFormat().getPreferredWidth().getValue());
    }

    @Test
    public void insertTableFromHtml() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the table from HTML. Note that AutoFitSettings does not apply to tables
        // inserted from HTML.
        builder.insertHtml("<table>" + "<tr>" + "<td>Row 1, Cell 1</td>" + "<td>Row 1, Cell 2</td>" + "</tr>" +
                "<tr>" + "<td>Row 2, Cell 2</td>" + "<td>Row 2, Cell 2</td>" + "</tr>" + "</table>");

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertTableFromHtml.docx");

        // Verify the table was constructed properly
        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertTableFromHtml.docx");

        Assert.assertEquals(1, doc.getChildNodes(NodeType.TABLE, true).getCount());
        Assert.assertEquals(2, doc.getChildNodes(NodeType.ROW, true).getCount());
        Assert.assertEquals(4, doc.getChildNodes(NodeType.CELL, true).getCount());
    }

    @Test
    public void insertNestedTable() throws Exception {
        //ExStart
        //ExFor:Cell.FirstParagraph
        //ExSummary:Shows how to insert a nested table using DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the outer table
        Cell cell = builder.insertCell();
        builder.writeln("Outer Table Cell 1");

        builder.insertCell();
        builder.writeln("Outer Table Cell 2");

        // This call is important in order to create a nested table within the first table
        // Without this call the cells inserted below will be appended to the outer table
        builder.endTable();

        // Move to the first cell of the outer table
        builder.moveTo(cell.getFirstParagraph());

        // Build the inner table
        builder.insertCell();
        builder.writeln("Inner Table Cell 1");
        builder.insertCell();
        builder.writeln("Inner Table Cell 2");

        builder.endTable();

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertNestedTable.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertNestedTable.docx");

        Assert.assertEquals(2, doc.getChildNodes(NodeType.TABLE, true).getCount());
        Assert.assertEquals(4, doc.getChildNodes(NodeType.CELL, true).getCount());
        Assert.assertEquals(1, cell.getTables().get(0).getCount());
        Assert.assertEquals(2, cell.getTables().get(0).getFirstRow().getCells().getCount());
    }

    @Test
    public void createSimpleTable() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder
        //ExFor:DocumentBuilder.Write
        //ExFor:DocumentBuilder.InsertCell
        //ExSummary:Shows how to create a simple table using DocumentBuilder with default formatting.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // We call this method to start building the table
        builder.startTable();
        builder.insertCell();
        builder.write("Row 1, Cell 1 Content.");

        // Build the second cell
        builder.insertCell();
        builder.write("Row 1, Cell 2 Content.");
        // Call the following method to end the row and start a new row
        builder.endRow();

        // Build the first cell of the second row
        builder.insertCell();
        builder.write("Row 2, Cell 1 Content.");

        // Build the second cell.
        builder.insertCell();
        builder.write("Row 2, Cell 2 Content.");
        builder.endRow();

        // Signal that we have finished building the table
        builder.endTable();

        // Save the document to disk
        doc.save(getArtifactsDir() + "DocumentBuilder.CreateSimpleTable.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.CreateSimpleTable.docx");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(4, table.getChildNodes(NodeType.CELL, true).getCount());

        Assert.assertEquals("Row 1, Cell 1 Content.", table.getRows().get(0).getCells().get(0).getText().trim());
        Assert.assertEquals("Row 1, Cell 2 Content.", table.getRows().get(0).getCells().get(1).getText().trim());
        Assert.assertEquals("Row 2, Cell 1 Content.", table.getRows().get(1).getCells().get(0).getText().trim());
        Assert.assertEquals("Row 2, Cell 2 Content.", table.getRows().get(1).getCells().get(1).getText().trim());
    }

    @Test
    public void buildFormattedTable() throws Exception {
        //ExStart
        //ExFor:RowFormat.Height
        //ExFor:RowFormat.HeightRule
        //ExFor:Table.LeftIndent
        //ExFor:DocumentBuilder.ParagraphFormat
        //ExFor:DocumentBuilder.Font
        //ExSummary:Shows how to create a formatted table using DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();

        // Make the header row
        builder.insertCell();

        // Set the left indent for the table. Table wide formatting must be applied after 
        // at least one row is present in the table
        table.setLeftIndent(20.0);

        // Set height and define the height rule for the header row
        builder.getRowFormat().setHeight(40.0);
        builder.getRowFormat().setHeightRule(HeightRule.AT_LEAST);

        // Some special features for the header row
        builder.getCellFormat().getShading().setBackgroundPatternColor(new Color(198, 217, 241));
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.getFont().setSize(16);
        builder.getFont().setName("Arial");
        builder.getFont().setBold(true);

        builder.getCellFormat().setWidth(100.0);
        builder.write("Header Row,\n Cell 1");

        // We don't need to specify the width of this cell because it's inherited from the previous cell
        builder.insertCell();
        builder.write("Header Row,\n Cell 2");

        builder.insertCell();
        builder.getCellFormat().setWidth(200.0);
        builder.write("Header Row,\n Cell 3");
        builder.endRow();

        // Set features for the other rows and cells
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.WHITE);
        builder.getCellFormat().setWidth(100.0);
        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);

        // Reset height and define a different height rule for table body
        builder.getRowFormat().setHeight(30.0);
        builder.getRowFormat().setHeightRule(HeightRule.AUTO);
        builder.insertCell();
        // Reset font formatting
        builder.getFont().setSize(12);
        builder.getFont().setBold(false);

        // Build the other cells
        builder.write("Row 1, Cell 1 Content");
        builder.insertCell();
        builder.write("Row 1, Cell 2 Content");

        builder.insertCell();
        builder.getCellFormat().setWidth(200.0);
        builder.write("Row 1, Cell 3 Content");
        builder.endRow();

        builder.insertCell();
        builder.getCellFormat().setWidth(100.0);
        builder.write("Row 2, Cell 1 Content");

        builder.insertCell();
        builder.write("Row 2, Cell 2 Content");

        builder.insertCell();
        builder.getCellFormat().setWidth(200.0);
        builder.write("Row 2, Cell 3 Content.");
        builder.endRow();
        builder.endTable();

        doc.save(getArtifactsDir() + "DocumentBuilder.CreateFormattedTable.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.CreateFormattedTable.docx");
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(20.0d, table.getLeftIndent());

        Assert.assertEquals(HeightRule.AT_LEAST, table.getRows().get(0).getRowFormat().getHeightRule());
        Assert.assertEquals(40.0d, table.getRows().get(0).getRowFormat().getHeight());

        for (Cell c : (Iterable<Cell>) doc.getChildNodes(NodeType.CELL, true)) {
            Assert.assertEquals(ParagraphAlignment.CENTER, c.getFirstParagraph().getParagraphFormat().getAlignment());

            for (Run r : (Iterable<Run>) c.getFirstParagraph().getRuns()) {
                Assert.assertEquals("Arial", r.getFont().getName());

                if (c.getParentRow() == table.getFirstRow()) {
                    Assert.assertEquals(16.0, r.getFont().getSize());
                    Assert.assertTrue(r.getFont().getBold());
                } else {
                    Assert.assertEquals(12.0, r.getFont().getSize());
                    Assert.assertFalse(r.getFont().getBold());
                }
            }
        }
    }

    @Test
    public void tableBordersAndShading() throws Exception {
        //ExStart
        //ExFor:Shading
        //ExFor:Table.SetBorders
        //ExFor:BorderCollection.Left
        //ExFor:BorderCollection.Right
        //ExFor:BorderCollection.Top
        //ExFor:BorderCollection.Bottom
        //ExSummary:Shows how to format table and cell with different borders and shadings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and set a default color/thickness for its borders
        Table table = builder.startTable();
        table.setBorders(LineStyle.SINGLE, 2.0, Color.BLACK);

        // Set the cell shading for this cell
        builder.insertCell();
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.RED);
        builder.writeln("Cell #1");

        // Specify a different cell shading for the second cell
        builder.insertCell();
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.GREEN);
        builder.writeln("Cell #2");

        // End this row
        builder.endRow();

        // Clear the cell formatting from previous operations
        builder.getCellFormat().clearFormatting();

        // Create the second row
        builder.insertCell();
        builder.writeln("Cell #3");

        // Clear the cell formatting from the previous cell
        builder.getCellFormat().clearFormatting();

        builder.getCellFormat().getBorders().getLeft().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getRight().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getTop().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getBottom().setLineWidth(4.0);

        builder.insertCell();
        builder.writeln("Cell #4");

        doc.save(getArtifactsDir() + "DocumentBuilder.TableBordersAndShading.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.TableBordersAndShading.docx");
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        for (Cell c : table.getFirstRow()) {
            Assert.assertEquals(0.5d, c.getCellFormat().getBorders().getTop().getLineWidth());
            Assert.assertEquals(0.5d, c.getCellFormat().getBorders().getBottom().getLineWidth());
            Assert.assertEquals(0.5d, c.getCellFormat().getBorders().getLeft().getLineWidth());
            Assert.assertEquals(0.5d, c.getCellFormat().getBorders().getRight().getLineWidth());

            Assert.assertEquals(0, c.getCellFormat().getBorders().getLeft().getColor().getRGB());
            Assert.assertEquals(LineStyle.SINGLE, c.getCellFormat().getBorders().getLeft().getLineStyle());
        }

        Assert.assertEquals(Color.RED.getRGB(),
                table.getFirstRow().getFirstCell().getCellFormat().getShading().getBackgroundPatternColor().getRGB());
        Assert.assertEquals(Color.GREEN.getRGB(),
                table.getFirstRow().getCells().get(1).getCellFormat().getShading().getBackgroundPatternColor().getRGB());

        for (Cell c : table.getLastRow()) {
            Assert.assertEquals(4.0d, c.getCellFormat().getBorders().getTop().getLineWidth());
            Assert.assertEquals(4.0d, c.getCellFormat().getBorders().getBottom().getLineWidth());
            Assert.assertEquals(4.0d, c.getCellFormat().getBorders().getLeft().getLineWidth());
            Assert.assertEquals(4.0d, c.getCellFormat().getBorders().getRight().getLineWidth());

            Assert.assertEquals(0, c.getCellFormat().getBorders().getLeft().getColor().getRGB());
            Assert.assertEquals(LineStyle.SINGLE, c.getCellFormat().getBorders().getLeft().getLineStyle());
            Assert.assertEquals(0, c.getCellFormat().getShading().getBackgroundPatternColor().getRGB());
        }
    }

    @Test
    public void setPreferredTypeConvertUtil() throws Exception {
        //ExStart
        //ExFor:PreferredWidth.FromPoints
        //ExSummary:Shows how to specify a cell preferred width by converting inches to points.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.getCellFormat().setPreferredWidth(PreferredWidth.fromPoints(ConvertUtil.inchToPoint(3)));
        builder.insertCell();
        //ExEnd

        Assert.assertEquals(table.getFirstRow().getFirstCell().getCellFormat().getPreferredWidth().getValue(), 216.0);
    }

    @Test
    public void insertHyperlinkToLocalBookmark() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.StartBookmark
        //ExFor:DocumentBuilder.EndBookmark
        //ExFor:DocumentBuilder.InsertHyperlink
        //ExSummary:Shows how to insert a hyperlink referencing a local bookmark.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("Bookmark1");
        builder.write("Bookmarked text.");
        builder.endBookmark("Bookmark1");

        builder.writeln("Some other text");

        // Specify font formatting for the hyperlink
        builder.getFont().setColor(Color.BLUE);
        builder.getFont().setUnderline(Underline.SINGLE);

        // Insert hyperlink
        // Switch \o is used to provide hyperlink tip text
        builder.insertHyperlink("Hyperlink Text", "Bookmark1\" \\o \"Hyperlink Tip", true);

        // Clear hyperlink formatting
        builder.getFont().clearFormatting();

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertHyperlinkToLocalBookmark.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertHyperlinkToLocalBookmark.docx");
        FieldHyperlink hyperlink = (FieldHyperlink) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_HYPERLINK, " HYPERLINK \\l \"Bookmark1\" \\o \"Hyperlink Tip\" ", "Hyperlink Text", hyperlink);
        Assert.assertEquals("Bookmark1", hyperlink.getSubAddress());
        //Assert.IsTrue(doc.getRange().getBookmarks().Any(b => b.Name == "Bookmark1")); //TODO: Check how ot works on java
    }

    @Test
    public void documentBuilderCursorPosition() throws Exception {
        // Write some text in a blank Document using a DocumentBuilder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Hello world!");

        // If the builder's cursor is at the end of the document, there will be no nodes in front of it so the current node will be null
        Assert.assertNull(builder.getCurrentNode());

        // However, the current paragraph the cursor is in will be valid
        Assert.assertEquals("Hello world!", builder.getCurrentParagraph().getText().trim());

        // Move to the beginning of the document and place the cursor at an existing node
        builder.moveToDocumentStart();
        Assert.assertEquals(NodeType.RUN, builder.getCurrentNode().getNodeType());
    }

    @Test
    public void documentBuilderMoveToNode() throws Exception {
        //ExStart
        //ExFor:Story.LastParagraph
        //ExFor:DocumentBuilder.MoveTo(Node)
        //ExSummary:Shows how to move a DocumentBuilder's cursor position to a specified node.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph with the DocumentBuilder
        builder.writeln("Text 1. ");

        // Move the DocumentBuilder to the first paragraph of the document and add another paragraph
        Assert.assertEquals(doc.getFirstSection().getBody().getLastParagraph(), builder.getCurrentParagraph()); //ExSkip
        builder.moveTo(doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0));
        Assert.assertEquals(doc.getFirstSection().getBody().getFirstParagraph(), builder.getCurrentParagraph()); //ExSkip
        builder.writeln("Text 2. ");

        // Since we moved to a node before the first paragraph before we added a second paragraph,
        // the second paragraph will appear in front of the first paragraph
        Assert.assertEquals("Text 2. \rText 1.", doc.getText().trim());

        // We can move the DocumentBuilder back to the end of the document like this
        // and carry on adding text to the end of the document
        builder.moveTo(doc.getFirstSection().getBody().getLastParagraph());
        builder.writeln("Text 3. ");

        Assert.assertEquals("Text 2. \rText 1. \rText 3.", doc.getText().trim());
        Assert.assertEquals(doc.getFirstSection().getBody().getLastParagraph(), builder.getCurrentParagraph()); //ExSkip
        //ExEnd
    }

    @Test
    public void documentBuilderMoveToDocumentStartEnd() throws Exception {
        Document doc = new Document(getMyDir() + "Document.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToDocumentEnd();
        builder.writeln("This is the end of the document.");

        builder.moveToDocumentStart();
        builder.writeln("This is the beginning of the document.");
    }

    @Test
    public void documentBuilderMoveToSection() throws Exception {
        // Create a blank document and append a section to it, giving it two sections
        Document doc = new Document();
        doc.appendChild(new Section(doc));

        // Move a DocumentBuilder to the second section and add text
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToSection(1);
        builder.writeln("Text added to the 2nd section.");
    }

    @Test
    public void documentBuilderMoveToParagraph() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.MoveToParagraph
        //ExSummary:Shows how to move a cursor position to the specified paragraph.
        // Open a document with a lot of paragraphs
        Document doc = new Document(getMyDir() + "Paragraphs.docx");
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        Assert.assertEquals(22, paragraphs.getCount());

        // When we create a DocumentBuilder for a document, its cursor is at the very beginning of the document by default,
        // and any content added by the DocumentBuilder will just be prepended to the document
        DocumentBuilder builder = new DocumentBuilder(doc);

        Assert.assertEquals(0, paragraphs.indexOf(builder.getCurrentParagraph()));

        // We can manually move the DocumentBuilder to any paragraph in the document via a 0-based index like this
        builder.moveToParagraph(2, 0);
        Assert.assertEquals(2, paragraphs.indexOf(builder.getCurrentParagraph())); //ExSkip
        builder.writeln("This is a new third paragraph. ");
        //ExEnd

        Assert.assertEquals(3, paragraphs.indexOf(builder.getCurrentParagraph()));

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals("This is a new third paragraph.", doc.getFirstSection().getBody().getParagraphs().get(2).getText().trim());
    }

    @Test
    public void documentBuilderMoveToTableCell() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.MoveToCell
        //ExSummary:Shows how to move a cursor position to the specified table cell.
        Document doc = new Document(getMyDir() + "Tables.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder to row 3, cell 4 of the first table
        builder.moveToCell(0, 2, 3, 0);
        builder.write("\nCell contents added by DocumentBuilder");
        //ExEnd

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(table.getRows().get(2).getCells().get(3), builder.getCurrentNode().getParentNode().getParentNode());
        Assert.assertEquals("Cell contents added by DocumentBuilderCell 3 contents", table.getRows().get(2).getCells().get(3).getText().trim());

    }

    @Test
    public void documentBuilderMoveToBookmarkEnd() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.MoveToBookmark(String, Boolean, Boolean)
        //ExSummary:Shows how to move a cursor position to just after the bookmark end.
        Document doc = new Document(getMyDir() + "Bookmarks.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move to after the end of the first bookmark
        Assert.assertTrue(builder.moveToBookmark("MyBookmark1", false, true));
        builder.write(" Text appended via DocumentBuilder.");
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertFalse(doc.getRange().getBookmarks().get("MyBookmark1").getText().contains(" Text appended via DocumentBuilder."));
    }

    @Test
    public void documentBuilderBuildTable() throws Exception {
        //ExStart
        //ExFor:Table
        //ExFor:DocumentBuilder.StartTable
        //ExFor:DocumentBuilder.EndRow
        //ExFor:DocumentBuilder.EndTable
        //ExFor:DocumentBuilder.CellFormat
        //ExFor:DocumentBuilder.RowFormat
        //ExFor:DocumentBuilder.Write(String)
        //ExFor:DocumentBuilder.Writeln(String)
        //ExFor:CellVerticalAlignment
        //ExFor:CellFormat.Orientation
        //ExFor:TextOrientation
        //ExFor:Table.AutoFit
        //ExFor:AutoFitBehavior
        //ExSummary:Shows how to build a formatted table that contains 2 rows and 2 columns.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();

        // Insert a cell
        builder.insertCell();
        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
        builder.write("This is row 1 cell 1");

        // Use fixed column widths
        table.autoFit(AutoFitBehavior.FIXED_COLUMN_WIDTHS);

        // Insert a cell
        builder.insertCell();
        builder.write("This is row 1 cell 2");
        builder.endRow();

        // Insert a cell
        builder.insertCell();

        // Apply new row formatting
        builder.getRowFormat().setHeight(100.0);
        builder.getRowFormat().setHeightRule(HeightRule.EXACTLY);

        builder.getCellFormat().setOrientation(TextOrientation.UPWARD);
        builder.write("This is row 2 cell 1");

        // Insert a cell
        builder.insertCell();
        builder.getCellFormat().setOrientation(TextOrientation.DOWNWARD);
        builder.write("This is row 2 cell 2");

        builder.endRow();
        builder.endTable();
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(2, table.getRows().getCount());
        Assert.assertEquals(2, table.getRows().get(0).getCells().getCount());
        Assert.assertEquals(2, table.getRows().get(1).getCells().getCount());
        Assert.assertFalse(table.getAllowAutoFit());

        Assert.assertEquals(0.0, table.getRows().get(0).getRowFormat().getHeight());
        Assert.assertEquals(HeightRule.AUTO, table.getRows().get(0).getRowFormat().getHeightRule());
        Assert.assertEquals(100.0, table.getRows().get(1).getRowFormat().getHeight());
        Assert.assertEquals(HeightRule.EXACTLY, table.getRows().get(1).getRowFormat().getHeightRule());

        Assert.assertEquals("This is row 1 cell 1", table.getRows().get(0).getCells().get(0).getText().trim());
        Assert.assertEquals(CellVerticalAlignment.CENTER, table.getRows().get(0).getCells().get(0).getCellFormat().getVerticalAlignment());

        Assert.assertEquals("This is row 1 cell 2", table.getRows().get(0).getCells().get(1).getText().trim());

        Assert.assertEquals("This is row 2 cell 1", table.getRows().get(1).getCells().get(0).getText().trim());
        Assert.assertEquals(TextOrientation.UPWARD, table.getRows().get(1).getCells().get(0).getCellFormat().getOrientation());

        Assert.assertEquals("This is row 2 cell 2", table.getRows().get(1).getCells().get(1).getText().trim());
        Assert.assertEquals(TextOrientation.DOWNWARD, table.getRows().get(1).getCells().get(1).getCellFormat().getOrientation());
    }

    @Test
    public void tableCellVerticalRotatedFarEastTextOrientation() throws Exception {
        Document doc = new Document(getMyDir() + "Rotated cell text.docx");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        Cell cell = table.getFirstRow().getFirstCell();

        Assert.assertEquals(cell.getCellFormat().getOrientation(), TextOrientation.VERTICAL_ROTATED_FAR_EAST);

        doc = DocumentHelper.saveOpen(doc);

        table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        cell = table.getFirstRow().getFirstCell();

        Assert.assertEquals(cell.getCellFormat().getOrientation(), TextOrientation.VERTICAL_ROTATED_FAR_EAST);
    }

    @Test
    public void documentBuilderInsertBreak() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("This is page 1.");
        builder.insertBreak(BreakType.PAGE_BREAK);

        builder.writeln("This is page 2.");
        builder.insertBreak(BreakType.PAGE_BREAK);

        builder.writeln("This is page 3.");
    }

    @Test
    public void documentBuilderInsertInlineImage() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertImage(getImageDir() + "Transparent background logo.png");
    }

    @Test
    public void documentBuilderInsertFloatingImage() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert a floating image from a file or URL.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertImage(getImageDir() + "Transparent background logo.png", RelativeHorizontalPosition.MARGIN, 100.0,
                RelativeVerticalPosition.MARGIN, 100.0, 200.0, 100.0, WrapType.SQUARE);
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        Shape image = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.PNG, image);
        Assert.assertEquals(100.0d, image.getLeft());
        Assert.assertEquals(100.0d, image.getTop());
        Assert.assertEquals(200.0d, image.getWidth());
        Assert.assertEquals(100.0d, image.getHeight());
        Assert.assertEquals(WrapType.SQUARE, image.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.MARGIN, image.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.MARGIN, image.getRelativeVerticalPosition());
    }

    @Test
    public void insertImageFromUrl() throws Exception {
        // Insert an image from a URL
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertImage(getAsposelogoUri().toURL().openStream());

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertImageFromUrl.doc");

        // Verify that the image was inserted into the document
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        Assert.assertNotNull(shape);
        Assert.assertTrue(shape.hasImage());
    }

    @Test
    public void insertImageOriginalSize() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert a floating image from a file or URL and retain the original image size in the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Pass a negative value to the width and height values to specify using the size of the source image
        builder.insertImage(getImageDir() + "Logo.jpg", RelativeHorizontalPosition.MARGIN, 200.0,
                RelativeVerticalPosition.MARGIN, 100.0, -1, -1, WrapType.SQUARE);
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        Shape image = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, image);
        Assert.assertEquals(200.0d, image.getLeft());
        Assert.assertEquals(100.0d, image.getTop());
        Assert.assertEquals(268.0d, image.getWidth());
        Assert.assertEquals(268.0d, image.getHeight());
        Assert.assertEquals(WrapType.SQUARE, image.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.MARGIN, image.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.MARGIN, image.getRelativeVerticalPosition());
    }

    @Test
    public void documentBuilderInsertTextInputFormField() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertTextInput
        //ExSummary:Shows how to insert a text input form field into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertTextInput("TextInput", TextFormFieldType.REGULAR, "", "Hello", 0);
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        FormField formField = doc.getRange().getFormFields().get(0);

        Assert.assertTrue(formField.getEnabled());
        Assert.assertEquals("TextInput", formField.getName());
        Assert.assertEquals(0, formField.getMaxLength());
        Assert.assertEquals("Hello", formField.getResult());
        Assert.assertEquals(FieldType.FIELD_FORM_TEXT_INPUT, formField.getType());
        Assert.assertEquals("", formField.getTextInputFormat());
        Assert.assertEquals(TextFormFieldType.REGULAR, formField.getTextInputType());
    }

    @Test
    public void documentBuilderInsertComboBoxFormField() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertComboBox
        //ExSummary:Shows how to insert a combobox form field into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        String[] items = {"One", "Two", "Three"};
        builder.insertComboBox("DropDown", items, 0);
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        FormField formField = doc.getRange().getFormFields().get(0);

        Assert.assertTrue(formField.getEnabled());
        Assert.assertEquals("DropDown", formField.getName());
        Assert.assertEquals(0, formField.getDropDownSelectedIndex());
        Assert.assertEquals(Arrays.asList(new String[]{"One", "Two", "Three"}), formField.getDropDownItems());
        Assert.assertEquals(FieldType.FIELD_FORM_DROP_DOWN, formField.getType());
    }

    @Test
    public void documentBuilderInsertToc() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table of contents at the beginning of the document
        builder.insertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // The newly inserted table of contents will be initially empty
        // It needs to be populated by updating the fields in the document
        doc.updateFields();
    }

    @Test(enabled = false, description = "WORDSNET-16868, WORDSJAVA-2406")
    public void signatureLineProviderId() throws Exception {
        //ExStart
        //ExFor:SignatureLine.ProviderId
        //ExFor:SignatureLineOptions.ShowDate
        //ExFor:SignatureLineOptions.Email
        //ExFor:SignatureLineOptions.DefaultInstructions
        //ExFor:SignatureLineOptions.Instructions
        //ExFor:SignatureLineOptions.AllowComments
        //ExFor:DocumentBuilder.InsertSignatureLine(SignatureLineOptions)
        //ExFor:SignOptions.ProviderId
        //ExSummary:Shows how to sign document with personal certificate and specific signature line.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Date currentDate = new Date();

        SignatureLineOptions signatureLineOptions = new SignatureLineOptions();

        signatureLineOptions.setSigner("vderyushev");
        signatureLineOptions.setSignerTitle("QA");
        signatureLineOptions.setEmail("vderyushev@aspose.com");
        signatureLineOptions.setShowDate(true);
        signatureLineOptions.setDefaultInstructions(false);
        signatureLineOptions.setInstructions("You need more info about signature line");
        signatureLineOptions.setAllowComments(true);

        SignatureLine signatureLine = builder.insertSignatureLine(signatureLineOptions).getSignatureLine();
        signatureLine.setProviderId(UUID.fromString("CF5A7BB4-8F3C-4756-9DF6-BEF7F13259A2"));

        doc.save(getArtifactsDir() + "DocumentBuilder.SignatureLineProviderId.docx");

        SignOptions signOptions = new SignOptions();
        signOptions.setSignatureLineId(signatureLine.getId());
        signOptions.setProviderId(signatureLine.getProviderId());
        signOptions.setComments("Document was signed by vderyushev");
        signOptions.setSignTime(currentDate);

        CertificateHolder certHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        DigitalSignatureUtil.sign(getArtifactsDir() + "DocumentBuilder.SignatureLineProviderId.docx",
                getArtifactsDir() + "DocumentBuilder.SignatureLineProviderId.Signed.docx", certHolder, signOptions);
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.SignatureLineProviderId.Signed.docx");
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        signatureLine = shape.getSignatureLine();

        Assert.assertEquals("vderyushev", signatureLine.getSigner());
        Assert.assertEquals("QA", signatureLine.getSignerTitle());
        Assert.assertEquals("vderyushev@aspose.com", signatureLine.getEmail());
        Assert.assertTrue(signatureLine.getShowDate());
        Assert.assertFalse(signatureLine.getDefaultInstructions());
        Assert.assertEquals("You need more info about signature line", signatureLine.getInstructions());
        Assert.assertTrue(signatureLine.getAllowComments());
        Assert.assertTrue(signatureLine.isSigned());
        Assert.assertTrue(signatureLine.isValid());

        DigitalSignatureCollection signatures = DigitalSignatureUtil.loadSignatures(
                getArtifactsDir() + "DocumentBuilder.SignatureLineProviderId.Signed.docx");

        Assert.assertEquals(1, signatures.getCount());
        Assert.assertTrue(signatures.get(0).isValid());
        Assert.assertEquals("Document was signed by vderyushev", signatures.get(0).getComments());
        Assert.assertEquals(currentDate, signatures.get(0).getSignTime());
        Assert.assertEquals("CN=Morzal.Me", signatures.get(0).getIssuerName());
        Assert.assertEquals(DigitalSignatureType.XML_DSIG, signatures.get(0).getSignatureType());
    }

    @Test
    public void insertSignatureLineCurrentPosition() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertSignatureLine(SignatureLineOptions, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, WrapType)
        //ExSummary:Shows how to insert signature line at the specified position.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        SignatureLineOptions options = new SignatureLineOptions();
        options.setSigner("John Doe");
        options.setSignerTitle("Manager");
        options.setEmail("johndoe@aspose.com");
        options.setShowDate(true);
        options.setDefaultInstructions(false);
        options.setInstructions("You need more info about signature line");
        options.setAllowComments(true);

        builder.insertSignatureLine(options, RelativeHorizontalPosition.RIGHT_MARGIN, 2.0, RelativeVerticalPosition.PAGE, 3.0, WrapType.INLINE);
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        SignatureLine signatureLine = shape.getSignatureLine();

        Assert.assertEquals(signatureLine.getSigner(), "John Doe");
        Assert.assertEquals(signatureLine.getSignerTitle(), "Manager");
        Assert.assertEquals(signatureLine.getEmail(), "johndoe@aspose.com");
        Assert.assertEquals(signatureLine.getShowDate(), true);
        Assert.assertEquals(signatureLine.getDefaultInstructions(), false);
        Assert.assertEquals(signatureLine.getInstructions(), "You need more info about signature line");
        Assert.assertEquals(signatureLine.getAllowComments(), true);
        Assert.assertEquals(signatureLine.isSigned(), false);
        Assert.assertEquals(signatureLine.isValid(), false);
    }

    @Test
    public void documentBuilderSetFontFormatting() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set font formatting properties
        Font font = builder.getFont();
        font.setBold(true);
        font.setColor(Color.BLUE);
        font.setItalic(true);
        font.setName("Arial");
        font.setSize(24.0);
        font.setSpacing(5.0);
        font.setUnderline(Underline.DOUBLE);

        // Output formatted text
        builder.writeln("I'm a very nice formatted String.");
    }

    @Test
    public void documentBuilderSetParagraphFormatting() throws Exception {
        //ExStart
        //ExFor:ParagraphFormat.RightIndent
        //ExFor:ParagraphFormat.LeftIndent
        //ExSummary:Shows how to set paragraph formatting.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set paragraph formatting properties
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();
        paragraphFormat.setAlignment(ParagraphAlignment.CENTER);
        paragraphFormat.setLeftIndent(50.0);
        paragraphFormat.setRightIndent(50.0);
        paragraphFormat.setSpaceAfter(25.0);

        // Output text
        builder.writeln(
                "This paragraph demonstrates how the left and right indents affect word wrapping.");
        builder.writeln(
                "The space between the above paragraph and this one depends on the DocumentBuilder's paragraph format.");

        doc.save(getArtifactsDir() + "DocumentBuilder.DocumentBuilderSetParagraphFormatting.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.DocumentBuilderSetParagraphFormatting.docx");

        for (Paragraph paragraph : doc.getFirstSection().getBody().getParagraphs()) {
            Assert.assertEquals(ParagraphAlignment.CENTER, paragraph.getParagraphFormat().getAlignment());
            Assert.assertEquals(50.0d, paragraph.getParagraphFormat().getLeftIndent());
            Assert.assertEquals(50.0d, paragraph.getParagraphFormat().getRightIndent());
            Assert.assertEquals(25.0d, paragraph.getParagraphFormat().getSpaceAfter());

        }
    }

    @Test
    public void documentBuilderSetCellFormatting() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.CellFormat
        //ExFor:CellFormat.Width
        //ExFor:CellFormat.LeftPadding
        //ExFor:CellFormat.RightPadding
        //ExFor:CellFormat.TopPadding
        //ExFor:CellFormat.BottomPadding
        //ExFor:DocumentBuilder.StartTable
        //ExFor:DocumentBuilder.EndTable
        //ExSummary:Shows how to create a table that contains a single formatted cell.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startTable();
        builder.insertCell();

        // Set the cell formatting
        CellFormat cellFormat = builder.getCellFormat();
        cellFormat.setWidth(250.0);
        cellFormat.setLeftPadding(30.0);
        cellFormat.setRightPadding(30.0);
        cellFormat.setTopPadding(30.0);
        cellFormat.setBottomPadding(30.0);

        builder.write("Formatted cell");
        builder.endRow();
        builder.endTable();

        doc.save(getArtifactsDir() + "DocumentBuilder.DocumentBuilderSetCellFormatting.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.DocumentBuilderSetCellFormatting.docx");
        Cell firstCell = ((Table) doc.getChild(NodeType.TABLE, 0, true)).getFirstRow().getFirstCell();

        Assert.assertEquals("Formatted cell", firstCell.getText().trim());

        Assert.assertEquals(250.0d, firstCell.getCellFormat().getWidth());
        Assert.assertEquals(30.0d, firstCell.getCellFormat().getLeftPadding());
        Assert.assertEquals(30.0d, firstCell.getCellFormat().getRightPadding());
        Assert.assertEquals(30.0d, firstCell.getCellFormat().getTopPadding());
        Assert.assertEquals(30.0d, firstCell.getCellFormat().getBottomPadding());

    }

    @Test
    public void documentBuilderSetRowFormatting() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.RowFormat
        //ExFor:HeightRule
        //ExFor:RowFormat.Height
        //ExFor:RowFormat.HeightRule
        //ExFor:Table.LeftPadding
        //ExFor:Table.RightPadding
        //ExFor:Table.TopPadding
        //ExFor:Table.BottomPadding
        //ExSummary:Shows how to create a table that contains a single cell and apply row formatting.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();

        // Set the row formatting
        RowFormat rowFormat = builder.getRowFormat();
        rowFormat.setHeight(100.0);
        rowFormat.setHeightRule(HeightRule.EXACTLY);
        // These formatting properties are set on the table and are applied to all rows in the table
        table.setLeftPadding(30.0);
        table.setRightPadding(30.0);
        table.setTopPadding(30.0);
        table.setBottomPadding(30.0);

        builder.writeln("Contents of formatted row.");

        builder.endRow();
        builder.endTable();

        doc.save(getArtifactsDir() + "DocumentBuilder.DocumentBuilderSetRowFormatting.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.DocumentBuilderSetRowFormatting.docx");
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(30.0d, table.getLeftPadding());
        Assert.assertEquals(30.0d, table.getRightPadding());
        Assert.assertEquals(30.0d, table.getTopPadding());
        Assert.assertEquals(30.0d, table.getBottomPadding());

        Assert.assertEquals(100.0d, table.getFirstRow().getRowFormat().getHeight());
        Assert.assertEquals(HeightRule.EXACTLY, table.getFirstRow().getRowFormat().getHeightRule());
    }

    @Test
    public void documentBuilderSetListFormatting() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getListFormat().applyNumberDefault();

        builder.writeln("Item 1");
        builder.writeln("Item 2");

        builder.getListFormat().listIndent();

        builder.writeln("Item 2.1");
        builder.writeln("Item 2.2");

        builder.getListFormat().listIndent();

        builder.writeln("Item 2.2.1");
        builder.writeln("Item 2.2.2");

        builder.getListFormat().listOutdent();

        builder.writeln("Item 2.3");

        builder.getListFormat().listOutdent();

        builder.writeln("Item 3");

        builder.getListFormat().removeNumbers();
    }

    @Test
    public void documentBuilderSetSectionFormatting() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set page properties
        builder.getPageSetup().setOrientation(Orientation.LANDSCAPE);
        builder.getPageSetup().setLeftMargin(50.0);
        builder.getPageSetup().setPaperSize(PaperSize.PAPER_10_X_14);
    }

    @Test
    public void insertFootnote() throws Exception {
        //ExStart
        //ExFor:FootnoteType
        //ExFor:DocumentBuilder.InsertFootnote(FootnoteType,String)
        //ExFor:DocumentBuilder.InsertFootnote(FootnoteType,String,String)
        //ExSummary:Shows how to reference text with a footnote and an endnote.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some text and mark it with a footnote with the IsAuto attribute set to "true" by default,
        // so the marker seen in the body text will be auto-numbered at "1", and the footnote will appear at the bottom of the page
        builder.write("This text will be referenced by a footnote.");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote comment regarding referenced text.");

        // Insert more text and mark it with an endnote with a custom reference mark,
        // which will be used in place of the number "2" and will set "IsAuto" to false
        builder.write("This text will be referenced by an endnote.");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Endnote comment regarding referenced text.", "CustomMark");

        // Footnotes always appear at the bottom of the page of their referenced text, so this page break will not affect the footnote
        // On the other hand, endnotes are always at the end of the document, so this page break will push the endnote down to the next page
        builder.insertBreak(BreakType.PAGE_BREAK);

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertFootnote.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertFootnote.docx");

        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "",
                "Footnote comment regarding referenced text.", (Footnote) doc.getChild(NodeType.FOOTNOTE, 0, true));
        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, false, "CustomMark",
                "CustomMark Endnote comment regarding referenced text.", (Footnote) doc.getChild(NodeType.FOOTNOTE, 1, true));
    }

    @Test
    public void documentBuilderApplyParagraphStyle() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set paragraph style
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.TITLE);

        builder.write("Hello");
    }

    @Test
    public void documentBuilderApplyBordersAndShading() throws Exception {
        //ExStart
        //ExFor:BorderCollection.Item(BorderType)
        //ExFor:Shading
        //ExFor:TextureIndex
        //ExFor:ParagraphFormat.Shading
        //ExFor:Shading.Texture
        //ExFor:Shading.BackgroundPatternColor
        //ExFor:Shading.ForegroundPatternColor
        //ExSummary:Shows how to apply borders and shading to a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set paragraph borders
        BorderCollection borders = builder.getParagraphFormat().getBorders();
        borders.setDistanceFromText(20.0);
        borders.getByBorderType(BorderType.LEFT).setLineStyle(LineStyle.DOUBLE);
        borders.getByBorderType(BorderType.RIGHT).setLineStyle(LineStyle.DOUBLE);
        borders.getByBorderType(BorderType.TOP).setLineStyle(LineStyle.DOUBLE);
        borders.getByBorderType(BorderType.BOTTOM).setLineStyle(LineStyle.DOUBLE);

        // Set paragraph shading
        Shading shading = builder.getParagraphFormat().getShading();
        shading.setTexture(TextureIndex.TEXTURE_DIAGONAL_CROSS);
        shading.setBackgroundPatternColor(new Color(240, 128, 128));  // Light Coral
        shading.setForegroundPatternColor(new Color(255, 160, 122));  // Light Salmon

        builder.write("This paragraph is formatted with a double border and shading.");
        doc.save(getArtifactsDir() + "DocumentBuilder.DocumentBuilderApplyBordersAndShading.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.DocumentBuilderApplyBordersAndShading.docx");
        borders = doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getBorders();

        Assert.assertEquals(20.0d, borders.getDistanceFromText());
        Assert.assertEquals(LineStyle.DOUBLE, borders.getByBorderType(BorderType.LEFT).getLineStyle());
        Assert.assertEquals(LineStyle.DOUBLE, borders.getByBorderType(BorderType.RIGHT).getLineStyle());
        Assert.assertEquals(LineStyle.DOUBLE, borders.getByBorderType(BorderType.TOP).getLineStyle());
        Assert.assertEquals(LineStyle.DOUBLE, borders.getByBorderType(BorderType.BOTTOM).getLineStyle());

        Assert.assertEquals(TextureIndex.TEXTURE_DIAGONAL_CROSS, shading.getTexture());
        Assert.assertEquals(Color.decode("#f08080").getRGB(), shading.getBackgroundPatternColor().getRGB());
        Assert.assertEquals(Color.decode("#ffa07a").getRGB(), shading.getForegroundPatternColor().getRGB());

    }

    @Test
    public void deleteRow() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.DeleteRow
        //ExSummary:Shows how to delete a row from a table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with 2 rows
        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Cell 1");
        builder.insertCell();
        builder.write("Cell 2");
        builder.endRow();
        builder.insertCell();
        builder.write("Cell 3");
        builder.insertCell();
        builder.write("Cell 4");
        builder.endTable();

        Assert.assertEquals(2, table.getRows().getCount());

        // Delete the first row of the first table in the document
        builder.deleteRow(0, 0);

        Assert.assertEquals(1, table.getRows().getCount());
        //ExEnd

        Assert.assertEquals("Cell 3\u0007Cell 4", table.getText().trim());
    }

    @Test(enabled = false, description = "Bug: does not insert headers and footers, all lists (bullets, numbering, multilevel) breaks")
    public void insertDocument() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertDocument(Document, ImportFormatMode)
        //ExFor:ImportFormatMode
        //ExSummary:Shows how to insert a document content into another document keep formatting of inserted document.
        Document doc = new Document(getMyDir() + "Document.docx");

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveToDocumentEnd();
        builder.insertBreak(BreakType.PAGE_BREAK);

        Document docToInsert = new Document(getMyDir() + "Formatted elements.docx");

        builder.insertDocument(docToInsert, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        builder.getDocument().save(getArtifactsDir() + "DocumentBuilder.InsertDocument.docx");
        //ExEnd

        Assert.assertEquals(29, doc.getStyles().getCount());
        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "DocumentBuilder.InsertDocument.docx",
                getGoldsDir() + "DocumentBuilder.InsertDocument Gold.docx"));
    }

    @Test
    public void keepSourceNumbering() throws Exception {
        //ExStart
        //ExFor:ImportFormatOptions.KeepSourceNumbering
        //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how the numbering will be imported when it clashes in source and destination documents.
        // Open a document with a custom list numbering scheme and clone it
        // Since both have the same numbering format, the formats will clash if we import one document into the other
        Document srcDoc = new Document(getMyDir() + "Custom list numbering.docx");
        Document dstDoc = srcDoc.deepClone();

        // Both documents have the same numbering in their lists, but if we set this flag to false and then import one document into the other
        // the numbering of the imported source document will continue from where it ends in the destination document
        ImportFormatOptions importFormatOptions = new ImportFormatOptions();
        importFormatOptions.setKeepSourceNumbering(false);

        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING, importFormatOptions);

        ParagraphCollection srcParas = srcDoc.getFirstSection().getBody().getParagraphs();
        for (int i = 0; i < srcParas.getCount(); i++) {
            Paragraph srcPara = srcParas.get(i);
            Node importedNode = importer.importNode(srcPara, true);
            dstDoc.getFirstSection().getBody().appendChild(importedNode);
        }

        dstDoc.updateListLabels();
        dstDoc.save(getArtifactsDir() + "DocumentBuilder.KeepSourceNumbering.docx");
        //ExEnd
    }

    @Test
    public void resolveStyleBehaviorWhileAppendDocument() throws Exception {
        //ExStart
        //ExFor:Document.AppendDocument(Document, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how to resolve styles behavior while append document.
        // Open a document with text in a custom style and clone it
        Document srcDoc = new Document(getMyDir() + "Custom list numbering.docx");
        Document dstDoc = srcDoc.deepClone();

        // We now have two documents, each with an identical style named "CustomStyle" 
        // We can change the text color of one of the styles
        dstDoc.getStyles().get("CustomStyle").getFont().setColor(Color.RED);

        ImportFormatOptions options = new ImportFormatOptions();
        // Specify that if numbering clashes in source and destination documents
        // then a numbering from the source document will be used
        options.setKeepSourceNumbering(true);

        // If we join two documents which have different styles that share the same name,
        // we can resolve the style clash with an ImportFormatMode
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_DIFFERENT_STYLES, options);
        dstDoc.updateListLabels();

        dstDoc.save(getArtifactsDir() + "DocumentBuilder.ResolveStyleBehaviorWhileAppendDocument.docx");
        //ExEnd
    }

    @Test(enabled = false, dataProvider = "ignoreTextBoxesDataProvider")
    public void ignoreTextBoxes(boolean isIgnoreTextBoxes) throws Exception {
        //ExStart
        //ExFor:ImportFormatOptions.IgnoreTextBoxes
        //ExSummary:Shows how to manage formatting in the text boxes of the source destination during the import.
        // Create a document and add text
        Document dstDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        builder.writeln("Hello world! Text box to follow.");

        // Create another document with a textbox, and insert some formatted text into it
        Document srcDoc = new Document();
        builder = new DocumentBuilder(srcDoc);

        Shape textBox = builder.insertShape(ShapeType.TEXT_BOX, 300.0, 100.0);
        builder.moveTo(textBox.getFirstParagraph());
        builder.getParagraphFormat().getStyle().getFont().setName("Courier New");
        builder.getParagraphFormat().getStyle().getFont().setSize(24.0d);
        builder.write("Textbox contents");

        // When we import the document with the textbox as a node into the first document, by default the text inside the text box will keep its formatting
        // Setting the IgnoreTextBoxes flag will clear the formatting during importing of the node
        ImportFormatOptions importFormatOptions = new ImportFormatOptions();
        importFormatOptions.setIgnoreTextBoxes(true);

        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING, importFormatOptions);

        ParagraphCollection srcParas = srcDoc.getFirstSection().getBody().getParagraphs();
        for (int i = 0; i < srcParas.getCount(); i++) {
            Paragraph srcPara = srcParas.get(i);
            Node importedNode = importer.importNode(srcPara, true);
            dstDoc.getFirstSection().getBody().appendChild(importedNode);
        }

        dstDoc.save(getArtifactsDir() + "DocumentBuilder.IgnoreTextBoxes.docx");
        //ExEnd

        dstDoc = new Document(getArtifactsDir() + "DocumentBuilder.IgnoreTextBoxes.docx");
        textBox = (Shape) dstDoc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals("Textbox contents", textBox.getText().trim());

        if (isIgnoreTextBoxes) {
            Assert.assertEquals(12.0d, textBox.getFirstParagraph().getRuns().get(0).getFont().getSize());
            Assert.assertEquals("Times New Roman", textBox.getFirstParagraph().getRuns().get(0).getFont().getName());
        } else {
            Assert.assertEquals(24.0d, textBox.getFirstParagraph().getRuns().get(0).getFont().getSize());
            Assert.assertEquals("Courier New", textBox.getFirstParagraph().getRuns().get(0).getFont().getName());
        }
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "ignoreTextBoxesDataProvider")
    public static Object[][] ignoreTextBoxesDataProvider() throws Exception {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }

    @Test
    public void moveToField() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.MoveToField
        //ExSummary:Shows how to move document builder's cursor to a specific field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a field using the DocumentBuilder and add a run of text after it
        Field field = builder.insertField("MERGEFIELD field");
        builder.write(" Text after the field.");

        // The builder's cursor is currently at end of the document
        Assert.assertNull(builder.getCurrentNode());

        // We can move the builder to a field like this, placing the cursor at immediately after the field
        builder.moveToField(field, true);

        // Note that the cursor is at a place past the FieldEnd node of the field, meaning that we are not actually inside the field
        // If we wish to move the DocumentBuilder to inside a field,
        // we will need to move it to a field's FieldStart or FieldSeparator node using the DocumentBuilder.MoveTo() method
        Assert.assertEquals(field.getEnd(), builder.getCurrentNode().getPreviousSibling());

        builder.write(" Text immediately after the field.");
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals("MERGEFIELD field\u0014«field»\u0015 Text immediately after the field. Text after the field.", doc.getText().trim());
    }

    @Test
    public void insertOleObjectException() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Assert.assertThrows(IllegalArgumentException.class, () -> builder.insertOleObject("", "checkbox", false, true, null));
    }

    @Test
    public void insertChartDouble() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertChart(ChartType, Double, Double)
        //ExSummary:Shows how to insert a chart into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertChart(ChartType.PIE, ConvertUtil.pixelToPoint(300.0), ConvertUtil.pixelToPoint(300.0));
        Assert.assertEquals(225.0d, ConvertUtil.pixelToPoint(300.0)); //ExSkip

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertedChartDouble.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertedChartDouble.docx");
        Shape chartShape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals("Chart Title", chartShape.getChart().getTitle().getText());
        Assert.assertEquals(225.0d, chartShape.getWidth());
        Assert.assertEquals(225.0d, chartShape.getHeight());
    }

    @Test
    public void insertChartRelativePosition() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertChart(ChartType, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert a chart into a document and specify position and size.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertChart(ChartType.PIE, RelativeHorizontalPosition.MARGIN, 100.0, RelativeVerticalPosition.MARGIN, 100.0, 200.0, 100.0, WrapType.SQUARE);

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertedChartRelativePosition.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertedChartRelativePosition.docx");
        Shape chartShape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(100.0d, chartShape.getTop());
        Assert.assertEquals(100.0d, chartShape.getLeft());
        Assert.assertEquals(200.0d, chartShape.getWidth());
        Assert.assertEquals(100.0d, chartShape.getHeight());
        Assert.assertEquals(WrapType.SQUARE, chartShape.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.MARGIN, chartShape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.MARGIN, chartShape.getRelativeVerticalPosition());
    }

    @Test
    public void insertField() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertField(String)
        //ExFor:Field
        //ExFor:Field.Update
        //ExFor:Field.Result
        //ExFor:Field.GetFieldCode
        //ExFor:Field.Type
        //ExFor:Field.Remove
        //ExFor:FieldType
        //ExSummary:Shows how to insert a field into a document by FieldCode.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple Date field into the document
        // When we insert a field through the DocumentBuilder class we can get the
        // special Field object which contains information about the field
        Field dateField = builder.insertField("DATE \\* MERGEFORMAT");

        // Update this particular field in the document so we can get the FieldResult
        dateField.update();

        // Display some information from this field
        // The field result is where the last evaluated value is stored. This is what is displayed in the document
        // When field codes are not showing
        Assert.assertEquals(LocalDate.now().format(DateTimeFormatter.ofPattern("M/d/YYYY")), dateField.getResult());

        // Display the field code which defines the behavior of the field. This can been seen in Microsoft Word by pressing ALT+F9
        Assert.assertEquals("DATE \\* MERGEFORMAT", dateField.getFieldCode());

        // The field type defines what type of field in the Document this is. In this case the type is "FieldDate" 
        Assert.assertEquals(FieldType.FIELD_DATE, dateField.getType());

        // Finally let's completely remove the field from the document. This can easily be done by invoking the Remove method on the object
        dateField.remove();
        //ExEnd			

        Assert.assertEquals(0, doc.getRange().getFields().getCount());
    }

    @Test
    public void insertFieldByType() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertField(FieldType, Boolean)
        //ExSummary:Shows how to insert a field into a document using FieldType.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an AUTHOR field using a DocumentBuilder
        doc.getBuiltInDocumentProperties().setAuthor("John Doe");
        builder.write("This document was written by ");
        builder.insertField(FieldType.FIELD_AUTHOR, true);
        Assert.assertEquals(" AUTHOR ", doc.getRange().getFields().get(0).getFieldCode()); //ExSkip
        Assert.assertEquals("John Doe", doc.getRange().getFields().get(0).getResult()); //ExSkip

        // Insert a PAGE field using a DocumentBuilder, but do not immediately update it
        builder.write("\nThis is page ");
        builder.insertField(FieldType.FIELD_PAGE, false);
        Assert.assertEquals(" PAGE ", doc.getRange().getFields().get(1).getFieldCode()); //ExSkip
        Assert.assertEquals("", doc.getRange().getFields().get(1).getResult()); //ExSkip

        // Some fields types, such as ones that display document word/page counts may not keep track of their results in real time,
        // and will only display an accurate result during a field update
        // We can defer the updating of those fields until right before we need to see an accurate result
        // This method will manually update all the fields in a document
        doc.updateFields();

        Assert.assertEquals("1", doc.getRange().getFields().get(1).getResult());
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals("This document was written by \u0013 AUTHOR \u0014John Doe\u0015" +
                "\rThis is page \u0013 PAGE \u00141", doc.getText().trim());

        TestUtil.verifyField(FieldType.FIELD_AUTHOR, " AUTHOR ", "John Doe", doc.getRange().getFields().get(0));
        TestUtil.verifyField(FieldType.FIELD_PAGE, " PAGE ", "1", doc.getRange().getFields().get(1));
    }

    //ExStart
    //ExFor:IFieldResultFormatter
    //ExFor:IFieldResultFormatter.Format(Double, GeneralFormat)
    //ExFor:IFieldResultFormatter.Format(String, GeneralFormat)
    //ExFor:IFieldResultFormatter.FormatDateTime(DateTime, String, CalendarType)
    //ExFor:IFieldResultFormatter.FormatNumeric(Double, String)
    //ExFor:FieldOptions.ResultFormatter
    //ExFor:CalendarType
    //ExSummary:Shows how to control how the field result is formatted.
    @Test //ExSkip
    public void fieldResultFormatting() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        doc.getFieldOptions().setResultFormatter(new FieldResultFormatter("${0}", "Date: {0}", "Item # {0}:"));

        // Insert a field with a numeric format
        builder.insertField(" = 2 + 3 \\# $###", null);

        // Insert a field with a date/time format
        builder.insertField("DATE \\@ \"d MMMM yyyy\"", null);

        // Insert a field with a general format
        builder.insertField("QUOTE \"2\" \\* Ordinal", null);

        // Formats will be applied and recorded by the formatter during the field update
        doc.updateFields();
        ((FieldResultFormatter) doc.getFieldOptions().getResultFormatter()).printInvocations();

        // Our formatter has also overridden the formats that were originally applied in the fields
        Assert.assertEquals(doc.getRange().getFields().get(0).getResult(), "$5");
        Assert.assertTrue(doc.getRange().getFields().get(1).getResult().startsWith("Date: "));
        Assert.assertEquals(doc.getRange().getFields().get(2).getResult(), "Item # 2:");
    }

    /// <summary>
    /// Custom IFieldResult implementation that applies formats and tracks format invocations
    /// </summary>
    private static class FieldResultFormatter implements IFieldResultFormatter {
        public FieldResultFormatter(final String numberFormat, final String dateFormat, final String generalFormat) {
            mNumberFormat = numberFormat;
            mDateFormat = dateFormat;
            mGeneralFormat = generalFormat;
        }

        public String formatNumeric(final double value, final String format) {
            mNumberFormatInvocations.add(new Object[]{value, format});

            return (mNumberFormat == null || "".equals(mNumberFormat)) ? null : MessageFormat.format(mNumberFormat, value);
        }

        public String formatDateTime(final Date value, final String format, final int calendarType) {
            mDateFormatInvocations.add(new Object[]{value, format, calendarType});

            return (mDateFormat == null || "".equals(mDateFormat)) ? null : MessageFormat.format(mDateFormat, value);
        }

        public String format(final String value, final int format) {
            return format((Object) value, format);
        }

        public String format(final double value, final int format) {
            return format((Object) value, format);
        }

        private String format(final Object value, final int format) {
            mGeneralFormatInvocations.add(new Object[]{value, format});

            return (mGeneralFormat == null || "".equals(mGeneralFormat)) ? null : MessageFormat.format(mGeneralFormat, value);
        }

        public void printInvocations() {
            System.out.println(MessageFormat.format("Number format invocations ({0}):", mNumberFormatInvocations.size()));
            for (Object[] s : (Iterable<Object[]>) mNumberFormatInvocations) {
                System.out.println("\tValue: " + s[0] + ", original format: " + s[1]);
            }

            System.out.println(MessageFormat.format("Date format invocations ({0}):", mDateFormatInvocations.size()));
            for (Object[] s : (Iterable<Object[]>) mDateFormatInvocations) {
                System.out.println("\tValue: " + s[0] + ", original format: " + s[1] + ", calendar type: " + s[2]);
            }

            System.out.println(MessageFormat.format("General format invocations ({0}):", mGeneralFormatInvocations.size()));
            for (Object[] s : (Iterable<Object[]>) mGeneralFormatInvocations) {
                System.out.println("\tValue: " + s[0] + ", original format: " + s[1]);
            }
        }

        private String mNumberFormat;
        private String mDateFormat;
        private String mGeneralFormat;

        private ArrayList mNumberFormatInvocations = new ArrayList();
        private ArrayList mDateFormatInvocations = new ArrayList();
        private ArrayList mGeneralFormatInvocations = new ArrayList();

    }
    //ExEnd

    @Test
    public void insertVideoWithUrl() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertOnlineVideo(String, Double, Double)
        //ExSummary:Shows how to insert online video into a document using video url
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a video from Youtube
        builder.insertOnlineVideo("https://youtu.be/t_1LYZ102RA", 360.0, 270.0);

        // Click on the shape in the output document to watch the video from Microsoft Word
        doc.save(getArtifactsDir() + "DocumentBuilder.InsertVideoWithUrl.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertVideoWithUrl.docx");
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(480, 360, ImageType.JPEG, shape);

        Assert.assertEquals(360.0d, shape.getWidth());
        Assert.assertEquals(270.0d, shape.getHeight());
    }

    @Test
    public void insertUnderline() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.Underline
        //ExSummary:Shows how to set and edit a document builder's underline.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a new style for our underline
        builder.setUnderline(Underline.DASH);

        // Same object as DocumentBuilder.Font.Underline
        Assert.assertEquals(builder.getFont().getUnderline(), builder.getUnderline());
        Assert.assertEquals(builder.getFont().getUnderline(), Underline.DASH);

        // These properties will be applied to the underline as well
        builder.getFont().setColor(Color.BLUE);
        builder.getFont().setSize(32.0);

        builder.writeln("Underlined text.");

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertUnderline.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertUnderline.docx");
        Run firstRun = doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0);

        Assert.assertEquals("Underlined text.", firstRun.getText().trim());
        Assert.assertEquals(Underline.DASH, firstRun.getFont().getUnderline());
        Assert.assertEquals(Color.BLUE.getRGB(), firstRun.getFont().getColor().getRGB());
        Assert.assertEquals(32.0d, firstRun.getFont().getSize());
    }

    @Test
    public void currentStory() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.CurrentStory
        //ExSummary:Shows how to work with a document builder's current story.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // A Story is a type of node that have child Paragraph nodes, such as a Body,
        // which would usually be a parent node to a DocumentBuilder's current paragraph
        Assert.assertEquals(builder.getCurrentStory(), doc.getFirstSection().getBody());
        Assert.assertEquals(builder.getCurrentStory(), builder.getCurrentParagraph().getParentNode());
        Assert.assertEquals(StoryType.MAIN_TEXT, builder.getCurrentStory().getStoryType());

        builder.getCurrentStory().appendParagraph("Text added to current Story.");

        // A Story can contain tables too
        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Row 1 cell 1");
        builder.insertCell();
        builder.write("Row 1 cell 2");
        builder.endTable();

        // The table we just made is automatically placed in the story
        Assert.assertTrue(builder.getCurrentStory().getTables().contains(table));
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        Assert.assertEquals(1, doc.getFirstSection().getBody().getTables().getCount());
        Assert.assertEquals("Row 1 cell 1\u0007Row 1 cell 2\u0007\u0007\rText added to current Story.", doc.getFirstSection().getBody().getText().trim());
    }

    @Test
    public void insertOlePowerpoint() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertOleObject(Stream, String, Boolean, Image)
        //ExSummary:Shows how to use document builder to embed Ole objects in a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Let's take a spreadsheet from our system and insert it into the document
        InputStream spreadsheetStream = new FileInputStream(getMyDir() + "Spreadsheet.xlsx");

        // The spreadsheet can be activated by double clicking the panel that you'll see in the document immediately under the text we will add
        // We did not set the area to double click as an icon nor did we change its appearance so it looks like a simple panel
        builder.writeln("Spreadsheet Ole object:");
        builder.insertOleObject(spreadsheetStream, "OleObject.xlsx", false, null);

        // A powerpoint presentation is another type of object we can embed in our document
        // This time we'll also exercise some control over how it looks
        InputStream powerpointStream = new FileInputStream(getMyDir() + "Presentation.pptx");
        byte[] imageBytes = DocumentHelper.getBytesFromStream(getAsposelogoUri().toURL().openStream());
        BufferedImage image = ImageIO.read(new ByteArrayInputStream(imageBytes));

        // If we double click the image, the powerpoint presentation will open
        builder.insertParagraph();
        builder.writeln("Powerpoint Ole object:");
        builder.insertOleObject(powerpointStream, "OleObject.pptx", true, image);

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertOlePowerpoint.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertOlePowerpoint.docx");

        Assert.assertEquals(2, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        Assert.assertEquals("", shape.getOleFormat().getIconCaption());
        Assert.assertFalse(shape.getOleFormat().getOleIcon());

        shape = (Shape) doc.getChild(NodeType.SHAPE, 1, true);
        Assert.assertEquals("Unknown", shape.getOleFormat().getIconCaption());
        Assert.assertTrue(shape.getOleFormat().getOleIcon());
    }

    @Test
    public void insertStyleSeparator() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertStyleSeparator
        //ExSummary:Shows how to separate styles from two different paragraphs used in one logical printed paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Append text in the "Heading 1" style
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.write("This text is in a Heading style. ");

        // Insert a style separator
        builder.insertStyleSeparator();

        // The style separator appears in the form of a paragraph break that doesn't start a new line
        // So, while this looks like one continuous paragraph with two styles in the output document, 
        // it is actually two paragraphs with different styles, but no line break between the first and second paragraph
        Assert.assertEquals(2, doc.getFirstSection().getBody().getParagraphs().getCount());

        // Append text with another style
        Style paraStyle = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "MyParaStyle");
        paraStyle.getFont().setBold(false);
        paraStyle.getFont().setSize(8.0);
        paraStyle.getFont().setName("Arial");

        // Set the style of the current paragraph to our custom style
        // This will apply to only the text after the style separator
        builder.getParagraphFormat().setStyleName(paraStyle.getName());
        builder.write("This text is in a custom style. ");

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertStyleSeparator.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertStyleSeparator.docx");

        Assert.assertEquals(2, doc.getFirstSection().getBody().getParagraphs().getCount());
        Assert.assertEquals("This text is in a Heading style. \r This text is in a custom style.",
                doc.getText().trim());
    }

    @Test
    public void withoutStyleSeparator() throws Exception {
        DocumentBuilder builder = new DocumentBuilder(new Document());

        Style paraStyle = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "MyParaStyle");
        paraStyle.getFont().setBold(false);
        paraStyle.getFont().setSize(8.0);
        paraStyle.getFont().setName("Arial");

        // Append text with "Heading 1" style
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.write("This text is in a Heading style. ");

        // Append text with another style
        builder.getParagraphFormat().setStyleName(paraStyle.getName());
        builder.write("This text is in a custom style. ");

        builder.getDocument().save(getArtifactsDir() + "DocumentBuilder.WithoutStyleSeparator.docx");
    }

    @Test
    public void smartStyleBehavior() throws Exception {
        //ExStart
        //ExFor:ImportFormatOptions
        //ExFor:ImportFormatOptions.SmartStyleBehavior
        //ExFor:DocumentBuilder.InsertDocument(Document, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how to resolve styles behavior while inserting documents.
        Document dstDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        Style myStyle = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "MyStyle");
        myStyle.getFont().setSize(14.0);
        myStyle.getFont().setName("Courier New");
        myStyle.getFont().setColor(Color.BLUE);

        // Append text with custom style
        builder.getParagraphFormat().setStyleName(myStyle.getName());
        builder.writeln("Hello world!");

        // Clone the document, and edit the clone's "MyStyle" style so it is a different color than that of the original
        // If we append this document to the original, the different styles will clash since they are the same name, and we will need to resolve it
        Document srcDoc = dstDoc.deepClone();
        srcDoc.getStyles().get("MyStyle").getFont().setColor(Color.RED);

        // When SmartStyleBehavior is enabled,
        // a source style will be expanded into a direct attributes inside a destination document,
        // if KeepSourceFormatting importing mode is used
        ImportFormatOptions options = new ImportFormatOptions();
        options.setSmartStyleBehavior(true);

        builder.insertDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING, options);

        dstDoc.save(getArtifactsDir() + "DocumentBuilder.SmartStyleBehavior.docx");
        //ExEnd

        dstDoc = new Document(getArtifactsDir() + "DocumentBuilder.SmartStyleBehavior.docx");

        Assert.assertEquals(Color.BLUE.getRGB(), dstDoc.getStyles().get("MyStyle").getFont().getColor().getRGB());
        Assert.assertEquals("MyStyle", dstDoc.getFirstSection().getBody().getParagraphs().get(0).getParagraphFormat().getStyle().getName());

        Assert.assertEquals("Normal", dstDoc.getFirstSection().getBody().getParagraphs().get(1).getParagraphFormat().getStyle().getName());
        Assert.assertEquals(14.0, dstDoc.getFirstSection().getBody().getParagraphs().get(1).getRuns().get(0).getFont().getSize());
        Assert.assertEquals("Courier New", dstDoc.getFirstSection().getBody().getParagraphs().get(1).getRuns().get(0).getFont().getName());
        Assert.assertEquals(Color.RED.getRGB(), dstDoc.getFirstSection().getBody().getParagraphs().get(1).getRuns().get(0).getFont().getColor().getRGB());
    }

    /// <summary>
    /// All markdown tests work with the same file
    /// That's why we need order for them 
    /// </summary>
    @Test(priority = 1)
    public void markdownDocumentEmphases() throws Exception {
        DocumentBuilder builder = new DocumentBuilder();

        // Bold and Italic are represented as Font.Bold and Font.Italic
        builder.getFont().setItalic(true);
        builder.writeln("This text will be italic");

        // Use clear formatting if don't want to combine styles between paragraphs
        builder.getFont().clearFormatting();

        builder.getFont().setBold(true);
        builder.writeln("This text will be bold");

        builder.getFont().clearFormatting();

        // You can also write create BoldItalic text
        builder.getFont().setItalic(true);
        builder.write("You ");
        builder.getFont().setBold(true);
        builder.write("can");
        builder.getFont().setBold(false);
        builder.writeln(" combine them");

        builder.getFont().clearFormatting();

        builder.getFont().setStrikeThrough(true);
        builder.writeln("This text will be strikethrough");

        // Markdown treats asterisks (*), underscores (_) and tilde (~) as indicators of emphasis
        builder.getDocument().save(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file
    /// That's why we need order for them 
    /// </summary>
    @Test(priority = 2)
    public void markdownDocumentInlineCode() throws Exception {
        Document doc = new Document(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare our created document for further work
        // And clear paragraph formatting not to use the previous styles
        builder.moveToDocumentEnd();
        builder.getParagraphFormat().clearFormatting();
        builder.writeln("\n");

        // Style with name that starts from word InlineCode, followed by optional dot (.) and number of backticks (`)
        // If number of backticks is missed, then one backtick will be used by default
        Style inlineCode1BackTicks = doc.getStyles().add(StyleType.CHARACTER, "InlineCode");
        builder.getFont().setStyle(inlineCode1BackTicks);
        builder.writeln("Text with InlineCode style with one backtick");

        // Use optional dot (.) and number of backticks (`)
        // There will be 3 backticks
        Style inlineCode3BackTicks = doc.getStyles().add(StyleType.CHARACTER, "InlineCode.3");
        builder.getFont().setStyle(inlineCode3BackTicks);
        builder.writeln("Text with InlineCode style with 3 backticks");

        builder.getDocument().save(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file
    /// That's why we need order for them 
    /// </summary>
    @Test(description = "WORDSNET-19850", priority = 3)
    public void markdownDocumentHeadings() throws Exception {
        Document doc = new Document(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare our created document for further work
        // And clear paragraph formatting not to use the previous styles
        builder.moveToDocumentEnd();
        builder.getParagraphFormat().clearFormatting();
        builder.writeln("\n");

        // By default Heading styles in Word may have bold and italic formatting
        // If we do not want text to be emphasized, set these properties explicitly to false
        // Thus we can't use 'builder.Font.ClearFormatting()' because Bold/Italic will be set to true
        builder.getFont().setBold(false);
        builder.getFont().setItalic(false);

        // Create for one heading for each level
        builder.getParagraphFormat().setStyleName("Heading 1");
        builder.getFont().setItalic(true);
        builder.writeln("This is an italic H1 tag");

        // Reset our styles from the previous paragraph to not combine styles between paragraphs
        builder.getFont().setBold(false);
        builder.getFont().setItalic(false);

        // Structure-enhanced text heading can be added through style inheritance
        Style setextHeading1 = doc.getStyles().add(StyleType.PARAGRAPH, "SetextHeading1");
        builder.getParagraphFormat().setStyle(setextHeading1);
        doc.getStyles().get("SetextHeading1").setBaseStyleName("Heading 1");
        builder.writeln("SetextHeading 1");

        builder.getParagraphFormat().setStyleName("Heading 2");
        builder.writeln("This is an H2 tag");

        builder.getFont().setBold(false);
        builder.getFont().setItalic(false);

        Style setextHeading2 = doc.getStyles().add(StyleType.PARAGRAPH, "SetextHeading2");
        builder.getParagraphFormat().setStyle(setextHeading2);
        doc.getStyles().get("SetextHeading2").setBaseStyleName("Heading 2");
        builder.writeln("SetextHeading 2");

        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 3"));
        builder.writeln("This is an H3 tag");

        builder.getFont().setBold(false);
        builder.getFont().setItalic(false);

        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 4"));
        builder.getFont().setBold(true);
        builder.writeln("This is an bold H4 tag");

        builder.getFont().setBold(false);
        builder.getFont().setItalic(false);

        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 5"));
        builder.getFont().setItalic(true);
        builder.getFont().setBold(true);
        builder.writeln("This is an italic and bold H5 tag");

        builder.getFont().setBold(false);
        builder.getFont().setItalic(false);

        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 6"));
        builder.writeln("This is an H6 tag");

        doc.save(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file
    /// That's why we need order for them 
    /// </summary>
    @Test(priority = 4)
    public void markdownDocumentBlockquotes() throws Exception {
        Document doc = new Document(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare our created document for further work
        // And clear paragraph formatting not to use the previous styles
        builder.moveToDocumentEnd();
        builder.getParagraphFormat().clearFormatting();
        builder.writeln("\n");

        // By default document stores blockquote style for the first level
        builder.getParagraphFormat().setStyleName("Quote");
        builder.writeln("Blockquote");

        // Create styles for nested levels through style inheritance
        Style quoteLevel2 = doc.getStyles().add(StyleType.PARAGRAPH, "Quote1");
        builder.getParagraphFormat().setStyle(quoteLevel2);
        doc.getStyles().get("Quote1").setBaseStyleName("Quote");
        builder.writeln("1. Nested blockquote");

        Style quoteLevel3 = doc.getStyles().add(StyleType.PARAGRAPH, "Quote2");
        builder.getParagraphFormat().setStyle(quoteLevel3);
        doc.getStyles().get("Quote2").setBaseStyleName("Quote1");
        builder.getFont().setItalic(true);
        builder.writeln("2. Nested italic blockquote");

        Style quoteLevel4 = doc.getStyles().add(StyleType.PARAGRAPH, "Quote3");
        builder.getParagraphFormat().setStyle(quoteLevel4);
        doc.getStyles().get("Quote3").setBaseStyleName("Quote2");
        builder.getFont().setItalic(false);
        builder.getFont().setBold(true);
        builder.writeln("3. Nested bold blockquote");

        Style quoteLevel5 = doc.getStyles().add(StyleType.PARAGRAPH, "Quote4");
        builder.getParagraphFormat().setStyle(quoteLevel5);
        doc.getStyles().get("Quote4").setBaseStyleName("Quote3");
        builder.getFont().setBold(false);
        builder.writeln("4. Nested blockquote");

        Style quoteLevel6 = doc.getStyles().add(StyleType.PARAGRAPH, "Quote5");
        builder.getParagraphFormat().setStyle(quoteLevel6);
        doc.getStyles().get("Quote5").setBaseStyleName("Quote4");
        builder.writeln("5. Nested blockquote");

        Style quoteLevel7 = doc.getStyles().add(StyleType.PARAGRAPH, "Quote6");
        builder.getParagraphFormat().setStyle(quoteLevel7);
        doc.getStyles().get("Quote6").setBaseStyleName("Quote5");
        builder.getFont().setItalic(true);
        builder.getFont().setBold(true);
        builder.writeln("6. Nested italic bold blockquote");

        doc.save(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file
    /// That's why we need order for them 
    /// </summary>
    @Test(priority = 5)
    public void markdownDocumentIndentedCode() throws Exception {
        Document doc = new Document(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare our created document for further work
        // And clear paragraph formatting not to use the previous styles
        builder.moveToDocumentEnd();
        builder.writeln("\n");
        builder.getParagraphFormat().clearFormatting();
        builder.writeln("\n");

        Style indentedCode = doc.getStyles().add(StyleType.PARAGRAPH, "IndentedCode");
        builder.getParagraphFormat().setStyle(indentedCode);
        builder.writeln("This is an indented code");

        doc.save(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file
    /// That's why we need order for them 
    /// </summary>
    @Test(priority = 6)
    public void markdownDocumentFencedCode() throws Exception {
        Document doc = new Document(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare our created document for further work
        // And clear paragraph formatting not to use the previous styles
        builder.moveToDocumentEnd();
        builder.writeln("\n");
        builder.getParagraphFormat().clearFormatting();
        builder.writeln("\n");

        Style fencedCode = doc.getStyles().add(StyleType.PARAGRAPH, "FencedCode");
        builder.getParagraphFormat().setStyle(fencedCode);
        builder.writeln("This is a fenced code");

        Style fencedCodeWithInfo = doc.getStyles().add(StyleType.PARAGRAPH, "FencedCode.C#");
        builder.getParagraphFormat().setStyle(fencedCodeWithInfo);
        builder.writeln("This is a fenced code with info string");

        doc.save(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file
    /// That's why we need order for them 
    /// </summary>
    @Test(priority = 7)
    public void markdownDocumentHorizontalRule() throws Exception {
        Document doc = new Document(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare our created document for further work
        // And clear paragraph formatting not to use the previous styles
        builder.moveToDocumentEnd();
        builder.getParagraphFormat().clearFormatting();
        builder.writeln("\n");

        // Insert HorizontalRule that will be present in .md file as '-----'
        builder.insertHorizontalRule();

        builder.getDocument().save(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file
    /// That's why we need order for them 
    /// </summary>
    @Test(priority = 8)
    public void markdownDocumentBulletedList() throws Exception {
        Document doc = new Document(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare our created document for further work
        // And clear paragraph formatting not to use the previous styles
        builder.moveToDocumentEnd();
        builder.getParagraphFormat().clearFormatting();
        builder.writeln("\n");

        // Bulleted lists are represented using paragraph numbering
        builder.getListFormat().applyBulletDefault();
        // There can be 3 types of bulleted lists
        // The only diff in a numbering format of the very first level are: �?-’, �?+’ or �?*’ respectively
        builder.getListFormat().getList().getListLevels().get(0).setNumberFormat("-");

        builder.writeln("Item 1");
        builder.writeln("Item 2");
        builder.getListFormat().listIndent();
        builder.writeln("Item 2a");
        builder.writeln("Item 2b");

        builder.getDocument().save(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file.
    /// That's why we need order for them.
    /// </summary>
    @Test(dataProvider = "loadMarkdownDocumentAndAssertContentDataProvider", priority = 9)
    public void loadMarkdownDocumentAndAssertContent(String text, String styleName, boolean isItalic, boolean isBold) throws Exception {
        // Load created document from previous tests
        Document doc = new Document(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        for (Paragraph paragraph : (Iterable<Paragraph>) paragraphs) {
            if (paragraph.getRuns().getCount() != 0) {
                // Check that all document text has the necessary styles
                if (paragraph.getRuns().get(0).getText().equals(text) && !text.contains("InlineCode")) {
                    Assert.assertEquals(styleName, paragraph.getParagraphFormat().getStyle().getName());
                    Assert.assertEquals(isItalic, paragraph.getRuns().get(0).getFont().getItalic());
                    Assert.assertEquals(isBold, paragraph.getRuns().get(0).getFont().getBold());
                } else if (paragraph.getRuns().get(0).getText().equals(text) && text.contains("InlineCode")) {
                    Assert.assertEquals(styleName, paragraph.getRuns().get(0).getFont().getStyleName());
                }
            }

            // Check that document also has a HorizontalRule present as a shape
            NodeCollection shapesCollection = doc.getFirstSection().getBody().getChildNodes(NodeType.SHAPE, true);
            Shape horizontalRuleShape = (Shape) shapesCollection.get(0);

            Assert.assertTrue(shapesCollection.getCount() == 1);
            Assert.assertTrue(horizontalRuleShape.isHorizontalRule());
        }
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "loadMarkdownDocumentAndAssertContentDataProvider")
    public static Object[][] loadMarkdownDocumentAndAssertContentDataProvider() throws Exception {
        return new Object[][]
                {
                        {"Italic", "Normal", true, false},
                        {"Bold", "Normal", false, true},
                        {"ItalicBold", "Normal", true, true},
                        {"Text with InlineCode style with one backtick", "InlineCode", false, false},
                        {"Text with InlineCode style with 3 backticks", "InlineCode.3", false, false},
                        {"This is an italic H1 tag", "Heading 1", true, false},
                        {"SetextHeading 1", "SetextHeading1", false, false},
                        {"This is an H2 tag", "Heading 2", false, false},
                        {"SetextHeading 2", "SetextHeading2", false, false},
                        {"This is an H3 tag", "Heading 3", false, false},
                        {"This is an bold H4 tag", "Heading 4", false, true},
                        {"This is an italic and bold H5 tag", "Heading 5", true, true},
                        {"This is an H6 tag", "Heading 6", false, false},
                        {"Blockquote", "Quote", false, false},
                        {"1. Nested blockquote", "Quote1", false, false},
                        {"2. Nested italic blockquote", "Quote2", true, false},
                        {"3. Nested bold blockquote", "Quote3", false, true},
                        {"4. Nested blockquote", "Quote4", false, false},
                        {"5. Nested blockquote", "Quote5", false, false},
                        {"6. Nested italic bold blockquote", "Quote6", true, true},
                        {"This is an indented code", "IndentedCode", false, false},
                        {"This is a fenced code", "FencedCode", false, false},
                        {"This is a fenced code with info string", "FencedCode.C#", false, false},
                        {"Item 1", "Normal", false, false},
                };
    }

    @Test
    public void insertOnlineVideo() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertOnlineVideo(String, String, Byte[], Double, Double)
        //ExFor:DocumentBuilder.InsertOnlineVideo(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExFor:DocumentBuilder.InsertOnlineVideo(String, String, Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert online video into a document using html code.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Visible url
        String vimeoVideoUrl = "https://vimeo.com/52477838";

        // Embed Html code
        String vimeoEmbedCode =
                "<iframe src=\"https://player.vimeo.com/video/52477838\" width=\"640\" height=\"360\" frameborder=\"0\" title=\"Aspose\" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe>";

        // This video will have an automatically generated thumbnail, and we are setting the size according to its 16:9 aspect ratio
        builder.writeln("Video with an automatically generated thumbnail at the top left corner of the page:");
        builder.insertOnlineVideo(vimeoVideoUrl, RelativeHorizontalPosition.LEFT_MARGIN, 0.0,
                RelativeVerticalPosition.TOP_MARGIN, 0.0, 320.0, 180.0, WrapType.SQUARE);
        builder.insertBreak(BreakType.PAGE_BREAK);

        // We can get an image to use as a custom thumbnail
        byte[] imageBytes = DocumentHelper.getBytesFromStream(getAsposelogoUri().toURL().openStream());
        BufferedImage image = ImageIO.read(new ByteArrayInputStream(imageBytes));
        // This puts the video where we are with our document builder, with a custom thumbnail and size depending on the size of the image
        builder.writeln("Custom thumbnail at document builder's cursor:");
        builder.insertOnlineVideo(vimeoVideoUrl, vimeoEmbedCode, imageBytes, image.getWidth(), image.getHeight());
        builder.insertBreak(BreakType.PAGE_BREAK);

        // We can put the video at the bottom right edge of the page too, but we'll have to take the page margins into account
        double left = builder.getPageSetup().getRightMargin() - image.getWidth();
        double top = builder.getPageSetup().getBottomMargin() - image.getHeight();

        // Here we use a custom thumbnail and relative positioning to put it and the bottom right of tha page
        builder.writeln("Bottom right of page with custom thumbnail:");

        builder.insertOnlineVideo(vimeoVideoUrl, vimeoEmbedCode, imageBytes,
                RelativeHorizontalPosition.RIGHT_MARGIN, left, RelativeVerticalPosition.BOTTOM_MARGIN, top,
                image.getWidth(), image.getHeight(), WrapType.SQUARE);


        doc.save(getArtifactsDir() + "DocumentBuilder.InsertOnlineVideo.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertOnlineVideo.docx");
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(640, 360, ImageType.JPEG, shape);

        Assert.assertEquals(320.0d, shape.getWidth());
        Assert.assertEquals(180.0d, shape.getHeight());
        Assert.assertEquals(0.0d, shape.getLeft());
        Assert.assertEquals(0.0d, shape.getTop());
        Assert.assertEquals(WrapType.SQUARE, shape.getWrapType());
        Assert.assertEquals(RelativeVerticalPosition.TOP_MARGIN, shape.getRelativeVerticalPosition());
        Assert.assertEquals(RelativeHorizontalPosition.LEFT_MARGIN, shape.getRelativeHorizontalPosition());

        Assert.assertEquals("https://vimeo.com/52477838", shape.getHRef());

        shape = (Shape) doc.getChild(NodeType.SHAPE, 1, true);

        TestUtil.verifyImageInShape(320, 320, ImageType.PNG, shape);
        Assert.assertEquals(320.0d, shape.getWidth());
        Assert.assertEquals(320.0d, shape.getHeight());
        Assert.assertEquals(0.0d, shape.getLeft());
        Assert.assertEquals(0.0d, shape.getTop());
        Assert.assertEquals(WrapType.INLINE, shape.getWrapType());
        Assert.assertEquals(RelativeVerticalPosition.PARAGRAPH, shape.getRelativeVerticalPosition());
        Assert.assertEquals(RelativeHorizontalPosition.COLUMN, shape.getRelativeHorizontalPosition());

        Assert.assertEquals("https://vimeo.com/52477838", shape.getHRef());
    }
}
