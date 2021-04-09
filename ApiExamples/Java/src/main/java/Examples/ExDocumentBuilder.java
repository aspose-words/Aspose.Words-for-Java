package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.Font;
import com.aspose.words.List;
import com.aspose.words.Shape;
import com.aspose.words.*;
import org.apache.commons.collections4.IterableUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.net.URL;
import java.text.DecimalFormat;
import java.text.MessageFormat;
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
        //ExSummary:Shows how to insert formatted text using DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Specify font formatting, then add text.
        Font font = builder.getFont();
        font.setSize(16.0);
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

        // Specify that we want different headers and footers for first, even and odd pages.
        builder.getPageSetup().setDifferentFirstPageHeaderFooter(true);
        builder.getPageSetup().setOddAndEvenPagesHeaderFooter(true);

        // Create the headers, then add three pages to the document to display each header type.
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
        builder.write("Header for the first page");
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_EVEN);
        builder.write("Header for even pages");
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.write("Header for all other pages");

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
        //ExSummary:Shows how to insert fields, and move the document builder's cursor to them.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField("MERGEFIELD MyMergeField1 \\* MERGEFORMAT");
        builder.insertField("MERGEFIELD MyMergeField2 \\* MERGEFORMAT");

        // Move the cursor to the first MERGEFIELD.
        builder.moveToMergeField("MyMergeField1", true, false);

        // Note that the cursor is placed immediately after the first MERGEFIELD, and before the second.
        Assert.assertEquals(doc.getRange().getFields().get(1).getStart(), builder.getCurrentNode());
        Assert.assertEquals(doc.getRange().getFields().get(0).getEnd(), builder.getCurrentNode().getPreviousSibling());

        // If we wish to edit the field's field code or contents using the builder,
        // its cursor would need to be inside a field.
        // To place it inside a field, we would need to call the document builder's MoveTo method
        // and pass the field's start or separator node as an argument.
        builder.write(" Text between our merge fields. ");

        doc.save(getArtifactsDir() + "DocumentBuilder.MergeFields.docx");
        //ExEnd		

        doc = new Document(getArtifactsDir() + "DocumentBuilder.MergeFields.docx");

        Assert.assertEquals("MERGEFIELD MyMergeField1 \\* MERGEFORMAT\u0014«MyMergeField1»\u0015" +
                " Text between our merge fields. " +
                "\u0013MERGEFIELD MyMergeField2 \\* MERGEFORMAT\u0014«MyMergeField2»", doc.getText().trim());
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
        //ExSummary:Shows how to insert a horizontal rule shape, and customize its formatting.
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
        //ExSummary:Shows how to insert a hyperlink field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("For more information, please visit the ");

        // Insert a hyperlink and emphasize it with custom formatting.
        // The hyperlink will be a clickable piece of text which will take us to the location specified in the URL.
        builder.getFont().setColor(Color.BLUE);
        builder.getFont().setUnderline(Underline.SINGLE);
        builder.insertHyperlink("Google website", "https://www.google.com", false);
        builder.getFont().clearFormatting();
        builder.writeln(".");

        // Ctrl + left clicking the link in the text in Microsoft Word will take us to the URL via a new web browser window.
        doc.save(getArtifactsDir() + "DocumentBuilder.InsertHyperlink.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertHyperlink.docx");

        FieldHyperlink hyperlink = (FieldHyperlink)doc.getRange().getFields().get(0);
        TestUtil.verifyWebResponseStatusCode(200, new URL(hyperlink.getAddress()));

        Run fieldContents = (Run) hyperlink.getStart().getNextSibling();

        Assert.assertEquals(Color.BLUE.getRGB(), fieldContents.getFont().getColor().getRGB());
        Assert.assertEquals(Underline.SINGLE, fieldContents.getFont().getUnderline());
        Assert.assertEquals("HYPERLINK \"https://www.google.com\"", fieldContents.getText().trim());
    }

    @Test
    public void pushPopFont() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.PushFont
        //ExFor:DocumentBuilder.PopFont
        //ExFor:DocumentBuilder.InsertHyperlink
        //ExSummary:Shows how to use a document builder's formatting stack.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set up font formatting, then write the text that goes before the hyperlink.
        builder.getFont().setName("Arial");
        builder.getFont().setSize(24.0);
        builder.write("To visit Google, hold Ctrl and click ");

        // Preserve our current formatting configuration on the stack.
        builder.pushFont();

        // Alter the builder's current formatting by applying a new style.
        builder.getFont().setStyleIdentifier(StyleIdentifier.HYPERLINK);
        builder.insertHyperlink("here", "http://www.google.com", false);

        Assert.assertEquals(Color.BLUE.getRGB(), builder.getFont().getColor().getRGB());
        Assert.assertEquals(Underline.SINGLE, builder.getFont().getUnderline());

        // Restore the font formatting that we saved earlier and remove the element from the stack.
        builder.popFont();

        Assert.assertEquals(0, builder.getFont().getColor().getRGB());
        Assert.assertEquals(Underline.NONE, builder.getFont().getUnderline());

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
        //ExFor:WrapType
        //ExFor:RelativeHorizontalPosition
        //ExFor:RelativeVerticalPosition
        //ExSummary:Shows how to insert an image, and use it as a watermark.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image into the header so that it will be visible on every page.
        BufferedImage image = ImageIO.read(new File(getImageDir() + "Transparent background logo.png"));
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        Shape shape = builder.insertImage(image);
        shape.setWrapType(WrapType.NONE);
        shape.setBehindText(true);

        // Place the image at the center of the page.
        shape.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        shape.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
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
        //ExFor:DocumentBuilder.InsertOleObject(String, Boolean, Boolean, Stream)
        //ExFor:DocumentBuilder.InsertOleObject(String, String, Boolean, Boolean, Stream)
        //ExFor:DocumentBuilder.InsertOleObjectAsIcon(String, Boolean, String, String)
        //ExSummary:Shows how to insert an OLE object into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // OLE objects are links to files in our local file system that can be opened by other installed applications.
        // Double clicking these shapes will launch the application, and then use it to open the linked object.
        // There are three ways of using the InsertOleObject method to insert these shapes and configure their appearance.
        // If 'presentation' is omitted and 'asIcon' is set, this overloaded method selects
        // the icon according to the file extension and uses the filename for the icon caption.
        // 1 -  Image taken from the local file system:
        builder.insertOleObject(getMyDir() + "Spreadsheet.xlsx", false, false, new FileInputStream(getImageDir() + "Logo.jpg"));

        // If 'presentation' is omitted and 'asIcon' is set, this overloaded method selects
        // the icon according to 'progId' and uses the filename for the icon caption.
        // 2 -  Icon based on the application that will open the object:
        builder.insertOleObject(getMyDir() + "Spreadsheet.xlsx", "Excel.Sheet", false, true, new FileInputStream(getImageDir() + "Logo.jpg"));

        // If 'iconFile' and 'iconCaption' are omitted, this overloaded method selects
        // the icon according to 'progId' and uses the predefined icon caption.
        // 3 -  Image icon that's 32 x 32 pixels or smaller from the local file system, with a custom caption:
        builder.insertOleObjectAsIcon(getMyDir() + "Presentation.pptx", false, getImageDir() + "Logo icon.ico",
                "Double click to view presentation!");

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
        //ExSummary:Shows how to use a document builder to insert html content into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        final String HTML = "<p align='right'>Paragraph right</p>" +
                "<b>Implicit paragraph left</b>" +
                "<div align='center'>Div center</div>" +
                "<h1 align='left'>Heading 1 left.</h1>";

        builder.insertHtml(HTML);

        // Inserting HTML code parses the formatting of each element into equivalent document text formatting.
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

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertHtml.docx");
        //ExEnd
    }

    @Test(dataProvider = "insertHtmlWithFormattingDataProvider")
    public void insertHtmlWithFormatting(boolean useBuilderFormatting) throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertHtml(String, Boolean)
        //ExSummary:Shows how to apply a document builder's formatting while inserting HTML content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a text alignment for the builder, insert an HTML paragraph with a specified alignment, and one without.
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.DISTRIBUTED);
        builder.insertHtml(
                "<p align='right'>Paragraph 1.</p>" +
                        "<p>Paragraph 2.</p>", useBuilderFormatting);

        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        // The first paragraph has an alignment specified. When InsertHtml parses the HTML code,
        // the paragraph alignment value found in the HTML code always supersedes the document builder's value.
        Assert.assertEquals("Paragraph 1.", paragraphs.get(0).getText().trim());
        Assert.assertEquals(ParagraphAlignment.RIGHT, paragraphs.get(0).getParagraphFormat().getAlignment());

        // The second paragraph has no alignment specified. It can have its alignment value filled in
        // by the builder's value depending on the flag we passed to the InsertHtml method.
        Assert.assertEquals("Paragraph 2.", paragraphs.get(1).getText().trim());
        Assert.assertEquals(useBuilderFormatting ? ParagraphAlignment.DISTRIBUTED : ParagraphAlignment.LEFT,
                paragraphs.get(1).getParagraphFormat().getAlignment());

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertHtmlWithFormatting.docx");
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "insertHtmlWithFormattingDataProvider")
    public static Object[][] insertHtmlWithFormattingDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void mathMl() throws Exception {
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
        //ExSummary:Shows how create a bookmark.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // A valid bookmark needs to have document body text enclosed by
        // BookmarkStart and BookmarkEnd nodes created with a matching bookmark name.
        builder.startBookmark("MyBookmark");
        builder.writeln("Hello world!");
        builder.endBookmark("MyBookmark");

        Assert.assertEquals(1, doc.getRange().getBookmarks().getCount());
        Assert.assertEquals("MyBookmark", doc.getRange().getBookmarks().get(0).getName());
        Assert.assertEquals("Hello world!", doc.getRange().getBookmarks().get(0).getText().trim());
        //ExEnd
    }

    @Test
    public void createColumnBookmark() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.StartColumnBookmark
        //ExFor:DocumentBuilder.EndColumnBookmark
        //ExSummary:Shows how to create a column bookmark.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startTable();

        builder.insertCell();
        // Cells 1,2,4,5 will be bookmarked.
        builder.startColumnBookmark("MyBookmark_1");
        // Badly formed bookmarks or bookmarks with duplicate names will be ignored when the document is saved.
        builder.startColumnBookmark("MyBookmark_1");
        builder.startColumnBookmark("BadStartBookmark");
        builder.write("Cell 1");

        builder.insertCell();
        builder.write("Cell 2");

        builder.insertCell();
        builder.write("Cell 3");

        builder.endRow();

        builder.insertCell();
        builder.write("Cell 4");

        builder.insertCell();
        builder.write("Cell 5");
        builder.endColumnBookmark("MyBookmark_1");
        builder.endColumnBookmark("MyBookmark_1");

        Assert.assertThrows(IllegalStateException.class, () -> builder.endColumnBookmark("BadEndBookmark")); //ExSkip

        builder.insertCell();
        builder.write("Cell 6");

        builder.endRow();
        builder.endTable();

        doc.save(getArtifactsDir() + "Bookmarks.CreateColumnBookmark.docx");
        //ExEnd
    }

    @Test
    public void createForm() throws Exception
    {
        //ExStart
        //ExFor:TextFormFieldType
        //ExFor:DocumentBuilder.InsertTextInput
        //ExFor:DocumentBuilder.InsertComboBox
        //ExSummary:Shows how to create form fields.
        DocumentBuilder builder = new DocumentBuilder();

        // Form fields are objects in the document that the user can interact with by being prompted to enter values.
        // We can create them using a document builder, and below are two ways of doing so.
        // 1 -  Basic text input:
        builder.insertTextInput("My text input", TextFormFieldType.REGULAR,
                "", "Enter your name here", 30);

        // 2 -  Combo box with prompt text, and a range of possible values:
        String[] items =
                {
                        "-- Select your favorite footwear --", "Sneakers", "Oxfords", "Flip-flops", "Other"
                };

        builder.insertParagraph();
        builder.insertComboBox("My combo box", items, 0);

        builder.getDocument().save(getArtifactsDir() + "DocumentBuilder.CreateForm.docx");
        //ExEnd

        Document doc = new Document(getArtifactsDir() + "DocumentBuilder.CreateForm.docx");
        FormField formField = doc.getRange().getFormFields().get(0);

        Assert.assertEquals("My text input", formField.getName());
        Assert.assertEquals(TextFormFieldType.REGULAR, formField.getTextInputType());
        Assert.assertEquals("Enter your name here", formField.getResult());

        formField = doc.getRange().getFormFields().get(1);

        Assert.assertEquals("My combo box", formField.getName());
        Assert.assertEquals(TextFormFieldType.REGULAR, formField.getTextInputType());
        Assert.assertEquals("-- Select your favorite footwear --", formField.getResult());
        Assert.assertEquals(0, formField.getDropDownSelectedIndex());
        Assert.assertEquals(Arrays.asList("-- Select your favorite footwear --", "Sneakers", "Oxfords", "Flip-flops", "Other"),
                formField.getDropDownItems());
    }

    @Test
    public void insertCheckBox() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertCheckBox(string, bool, bool, int)
        //ExFor:DocumentBuilder.InsertCheckBox(String, bool, int)
        //ExSummary:Shows how to insert checkboxes into the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert checkboxes of varying sizes and default checked statuses.
        builder.write("Unchecked check box of a default size: ");
        builder.insertCheckBox("", false, false, 0);
        builder.insertParagraph();

        builder.write("Large checked check box: ");
        builder.insertCheckBox("CheckBox_Default", true, true, 50);
        builder.insertParagraph();

        // Form fields have a name length limit of 20 characters.
        builder.write("Very large checked check box: ");
        builder.insertCheckBox("CheckBox_OnlyCheckedValue", true, 100);

        Assert.assertEquals("CheckBox_OnlyChecked", doc.getRange().getFormFields().get(2).getName());

        // We can interact with these check boxes in Microsoft Word by double clicking them.
        doc.save(getArtifactsDir() + "DocumentBuilder.InsertCheckBox.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertCheckBox.docx");

        FormFieldCollection formFields = doc.getRange().getFormFields();

        Assert.assertEquals("", formFields.get(0).getName());
        Assert.assertEquals(false, formFields.get(0).getChecked());
        Assert.assertEquals(false, formFields.get(0).getDefault());
        Assert.assertEquals(10.0, formFields.get(0).getCheckBoxSize());

        Assert.assertEquals("CheckBox_Default", formFields.get(1).getName());
        Assert.assertEquals(true, formFields.get(1).getChecked());
        Assert.assertEquals(true, formFields.get(1).getDefault());
        Assert.assertEquals(50.0, formFields.get(1).getCheckBoxSize());

        Assert.assertEquals("CheckBox_OnlyChecked", formFields.get(2).getName());
        Assert.assertEquals(true, formFields.get(2).getChecked());
        Assert.assertEquals(true, formFields.get(2).getDefault());
        Assert.assertEquals(100.0, formFields.get(2).getCheckBoxSize());
    }

    @Test
    public void insertCheckBoxEmptyName() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Checking that the checkbox insertion with an empty name working correctly.
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
        //ExSummary:Shows how to move a document builder's cursor to different nodes in a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a valid bookmark, an entity that consists of nodes enclosed by a bookmark start node,
        // and a bookmark end node. 
        builder.startBookmark("MyBookmark");
        builder.write("Bookmark contents.");
        builder.endBookmark("MyBookmark");

        NodeCollection firstParagraphNodes = doc.getFirstSection().getBody().getFirstParagraph().getChildNodes();

        Assert.assertEquals(NodeType.BOOKMARK_START, firstParagraphNodes.get(0).getNodeType());
        Assert.assertEquals(NodeType.RUN, firstParagraphNodes.get(1).getNodeType());
        Assert.assertEquals("Bookmark contents.", firstParagraphNodes.get(1).getText().trim());
        Assert.assertEquals(NodeType.BOOKMARK_END, firstParagraphNodes.get(2).getNodeType());

        // The document builder's cursor is always ahead of the node that we last added with it.
        // If the builder's cursor is at the end of the document, its current node will be null.
        // The previous node is the bookmark end node that we last added.
        // Adding new nodes with the builder will append them to the last node.
        Assert.assertNull(builder.getCurrentNode());

        // If we wish to edit a different part of the document with the builder,
        // we will need to bring its cursor to the node we wish to edit.
        builder.moveToBookmark("MyBookmark");

        // Moving it to a bookmark will move it to the first node within the bookmark start and end nodes, the enclosed run.
        Assert.assertEquals(firstParagraphNodes.get(1), builder.getCurrentNode());

        // We can also move the cursor to an individual node like this.
        builder.moveTo(doc.getFirstSection().getBody().getFirstParagraph().getChildNodes(NodeType.ANY, false).get(0));

        Assert.assertEquals(NodeType.BOOKMARK_START, builder.getCurrentNode().getNodeType());
        Assert.assertEquals(doc.getFirstSection().getBody().getFirstParagraph(), builder.getCurrentParagraph());
        Assert.assertTrue(builder.isAtStartOfParagraph());

        // We can use specific methods to move to the start/end of a document.
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
        //ExSummary:Shows how to fill MERGEFIELDs with data with a document builder instead of a mail merge.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some MERGEFIELDS, which accept data from columns of the same name in a data source during a mail merge,
        // and then fill them manually.
        builder.insertField(" MERGEFIELD Chairman ");
        builder.insertField(" MERGEFIELD ChiefFinancialOfficer ");
        builder.insertField(" MERGEFIELD ChiefTechnologyOfficer ");

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

        // Insert a table of contents for the first page of the document.
        // Configure the table to pick up paragraphs with headings of levels 1 to 3.
        // Also, set its entries to be hyperlinks that will take us
        // to the location of the heading when left-clicked in Microsoft Word.
        builder.insertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.insertBreak(BreakType.PAGE_BREAK);

        // Populate the table of contents by adding paragraphs with heading styles.
        // Each such heading with a level between 1 and 3 will create an entry in the table.
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

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_4);
        builder.writeln("Heading 3.1.3.1");
        builder.writeln("Heading 3.1.3.2");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);
        builder.writeln("Heading 3.2");
        builder.writeln("Heading 3.3");

        // A table of contents is a field of a type that needs to be updated to show an up-to-date result.
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
        //ExSummary:Shows how to build a table with custom borders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startTable();

        // Setting table formatting options for a document builder
        // will apply them to every row and cell that we add with it.
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

        // Changing the formatting will apply it to the current cell,
        // and any new cells that we create with the builder afterward.
        // This will not affect the cells that we have added previously.
        builder.getCellFormat().getShading().clearFormatting();

        builder.insertCell();
        builder.write("Row 2, Col 1");

        builder.insertCell();
        builder.write("Row 2, Col 2");

        builder.endRow();

        // Increase row height to fit the vertical text.
        builder.insertCell();
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

        for (Cell c : table.getRows().get(0).getCells()) {
            Assert.assertEquals(150.0, c.getCellFormat().getWidth());
            Assert.assertEquals(CellVerticalAlignment.CENTER, c.getCellFormat().getVerticalAlignment());
            Assert.assertEquals(Color.GREEN.getRGB(), c.getCellFormat().getShading().getBackgroundPatternColor().getRGB());
            Assert.assertFalse(c.getCellFormat().getWrapText());
            Assert.assertTrue(c.getCellFormat().getFitText());

            Assert.assertEquals(ParagraphAlignment.CENTER, c.getFirstParagraph().getParagraphFormat().getAlignment());
        }

        Assert.assertEquals("Row 2, Col 1", table.getRows().get(1).getCells().get(0).getText().trim());
        Assert.assertEquals("Row 2, Col 2", table.getRows().get(1).getCells().get(1).getText().trim());


        for (Cell c : table.getRows().get(1).getCells()) {
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
        //ExSummary:Shows how to build a new table while applying a style.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.startTable();

        // We must insert at least one row before setting any table formatting.
        builder.insertCell();

        // Set the table style used based on the style identifier.
        // Note that not all table styles are available when saving to .doc format.
        table.setStyleIdentifier(StyleIdentifier.MEDIUM_SHADING_1_ACCENT_1);

        // Partially apply the style to features of the table based on predicates, then build the table.
        table.setStyleOptions(TableStyleOptions.FIRST_COLUMN | TableStyleOptions.ROW_BANDS | TableStyleOptions.FIRST_ROW);
        table.autoFit(AutoFitBehavior.AUTO_FIT_TO_CONTENTS);

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
        //ExSummary:Shows how to build a table with rows that repeat on every page. 
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();

        // Any rows inserted while the "HeadingFormat" flag is set to "true"
        // will show up at the top of the table on every page that it spans.
        builder.getRowFormat().setHeadingFormat(true);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.getCellFormat().setWidth(100.0);
        builder.insertCell();
        builder.write("Heading row 1");
        builder.endRow();
        builder.insertCell();
        builder.write("Heading row 2");
        builder.endRow();

        builder.getCellFormat().setWidth(50.0);
        builder.getParagraphFormat().clearFormatting();
        builder.getRowFormat().setHeadingFormat(false);

        // Add enough rows for the table to span two pages.
        for (int i = 0; i < 50; i++) {
            builder.insertCell();
            builder.write(MessageFormat.format("Row {0}, column 1.", table.getRows().toArray().length));
            builder.insertCell();
            builder.write(MessageFormat.format("Row {0}, column 2.", table.getRows().toArray().length));
            builder.endRow();
        }

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertTableSetHeadingRow.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertTableSetHeadingRow.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

        for (int i = 0; i < table.getRows().getCount(); i++)
            Assert.assertEquals(i < 2, table.getRows().get(i).getRowFormat().getHeadingFormat());
    }

    @Test
    public void insertTableWithPreferredWidth() throws Exception {
        //ExStart
        //ExFor:Table.PreferredWidth
        //ExFor:PreferredWidth.FromPercent
        //ExFor:PreferredWidth
        //ExSummary:Shows how to set a table to auto fit to 50% of the width of the page.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Cell #1");
        builder.insertCell();
        builder.write("Cell #2");
        builder.insertCell();
        builder.write("Cell #3");

        table.setPreferredWidth(PreferredWidth.fromPercent(50.0));

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertTableWithPreferredWidth.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertTableWithPreferredWidth.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

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
        //ExSummary:Shows how to set a preferred width for table cells.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.startTable();

        // There are two ways of applying the "PreferredWidth" class to table cells.
        // 1 -  Set an absolute preferred width based on points:
        builder.insertCell();
        builder.getCellFormat().setPreferredWidth(PreferredWidth.fromPoints(40.0));
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.YELLOW);
        builder.writeln(MessageFormat.format("Cell with a width of {0}.", builder.getCellFormat().getPreferredWidth()));

        // 2 -  Set a relative preferred width based on percent of the table's width:
        builder.insertCell();
        builder.getCellFormat().setPreferredWidth(PreferredWidth.fromPercent(20.0));
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.BLUE);
        builder.writeln(MessageFormat.format("Cell with a width of {0}.", builder.getCellFormat().getPreferredWidth()));

        builder.insertCell();

        // A cell with no preferred width specified will take up the rest of the available space.
        builder.getCellFormat().setPreferredWidth(PreferredWidth.AUTO);

        // Each configuration of the "PreferredWidth" property creates a new object.
        Assert.assertNotEquals(table.getFirstRow().getCells().get(1).getCellFormat().getPreferredWidth().hashCode(),
                builder.getCellFormat().getPreferredWidth().hashCode());

        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.GREEN);
        builder.writeln("Automatically sized cell.");

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertCellsWithPreferredWidths.docx");
        //ExEnd

        Assert.assertEquals(100.0d, PreferredWidth.fromPercent(100.0).getValue());
        Assert.assertEquals(100.0d, PreferredWidth.fromPoints(100.0).getValue());

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertCellsWithPreferredWidths.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertEquals(PreferredWidthType.POINTS, table.getFirstRow().getCells().get(0).getCellFormat().getPreferredWidth().getType());
        Assert.assertEquals(40.0d, table.getFirstRow().getCells().get(0).getCellFormat().getPreferredWidth().getValue());
        Assert.assertEquals("Cell with a width of 800.", table.getFirstRow().getCells().get(0).getText().trim());

        Assert.assertEquals(PreferredWidthType.PERCENT, table.getFirstRow().getCells().get(1).getCellFormat().getPreferredWidth().getType());
        Assert.assertEquals(20.0d, table.getFirstRow().getCells().get(1).getCellFormat().getPreferredWidth().getValue());
        Assert.assertEquals("Cell with a width of 20%.", table.getFirstRow().getCells().get(1).getText().trim());

        Assert.assertEquals(PreferredWidthType.AUTO, table.getFirstRow().getCells().get(2).getCellFormat().getPreferredWidth().getType());
        Assert.assertEquals(0.0d, table.getFirstRow().getCells().get(2).getCellFormat().getPreferredWidth().getValue());
        Assert.assertEquals("Automatically sized cell.", table.getFirstRow().getCells().get(2).getText().trim());
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

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertTableFromHtml.docx");

        Assert.assertEquals(1, doc.getChildNodes(NodeType.TABLE, true).getCount());
        Assert.assertEquals(2, doc.getChildNodes(NodeType.ROW, true).getCount());
        Assert.assertEquals(4, doc.getChildNodes(NodeType.CELL, true).getCount());
    }

    @Test
    public void insertNestedTable() throws Exception {
        //ExStart
        //ExFor:Cell.FirstParagraph
        //ExSummary:Shows how to create a nested table using a document builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the outer table.
        Cell cell = builder.insertCell();
        builder.writeln("Outer Table Cell 1");
        builder.insertCell();
        builder.writeln("Outer Table Cell 2");
        builder.endTable();

        // Move to the first cell of the outer table, the build another table inside the cell.
        builder.moveTo(cell.getFirstParagraph());
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
    public void createTable() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder
        //ExFor:DocumentBuilder.Write
        //ExFor:DocumentBuilder.InsertCell
        //ExSummary:Shows how to use a document builder to create a table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table, then populate the first row with two cells.
        builder.startTable();
        builder.insertCell();
        builder.write("Row 1, Cell 1.");
        builder.insertCell();
        builder.write("Row 1, Cell 2.");

        // Call the builder's "EndRow" method to start a new row.
        builder.endRow();
        builder.insertCell();
        builder.write("Row 2, Cell 1.");
        builder.insertCell();
        builder.write("Row 2, Cell 2.");
        builder.endTable();

        doc.save(getArtifactsDir() + "DocumentBuilder.CreateTable.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.CreateTable.docx");
        Table table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertEquals(4, table.getChildNodes(NodeType.CELL, true).getCount());

        Assert.assertEquals("Row 1, Cell 1.", table.getRows().get(0).getCells().get(0).getText().trim());
        Assert.assertEquals("Row 1, Cell 2.", table.getRows().get(0).getCells().get(1).getText().trim());
        Assert.assertEquals("Row 2, Cell 1.", table.getRows().get(1).getCells().get(0).getText().trim());
        Assert.assertEquals("Row 2, Cell 2.", table.getRows().get(1).getCells().get(1).getText().trim());
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
        builder.insertCell();
        table.setLeftIndent(20.0);

        // Set some formatting options for text and table appearance.
        builder.getRowFormat().setHeight(40.0);
        builder.getRowFormat().setHeightRule(HeightRule.AT_LEAST);
        builder.getCellFormat().getShading().setBackgroundPatternColor(new Color((198), (217), (241)));

        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.getFont().setSize(16.0);
        builder.getFont().setName("Arial");
        builder.getFont().setBold(true);

        // Configuring the formatting options in a document builder will apply them
        // to the current cell/row its cursor is in,
        // as well as any new cells and rows created using that builder.
        builder.write("Header Row,\n Cell 1");
        builder.insertCell();
        builder.write("Header Row,\n Cell 2");
        builder.insertCell();
        builder.write("Header Row,\n Cell 3");
        builder.endRow();

        // Reconfigure the builder's formatting objects for new rows and cells that we are about to make.
        // The builder will not apply these to the first row already created so that it will stand out as a header row.
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.WHITE);
        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
        builder.getRowFormat().setHeight(30.0);
        builder.getRowFormat().setHeightRule(HeightRule.AUTO);
        builder.insertCell();
        builder.getFont().setSize(12.0);
        builder.getFont().setBold(false);

        builder.write("Row 1, Cell 1.");
        builder.insertCell();
        builder.write("Row 1, Cell 2.");
        builder.insertCell();
        builder.write("Row 1, Cell 3.");
        builder.endRow();
        builder.insertCell();
        builder.write("Row 2, Cell 1.");
        builder.insertCell();
        builder.write("Row 2, Cell 2.");
        builder.insertCell();
        builder.write("Row 2, Cell 3.");
        builder.endRow();
        builder.endTable();

        doc.save(getArtifactsDir() + "DocumentBuilder.CreateFormattedTable.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.CreateFormattedTable.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertEquals(20.0d, table.getLeftIndent());

        Assert.assertEquals(HeightRule.AT_LEAST, table.getRows().get(0).getRowFormat().getHeightRule());
        Assert.assertEquals(40.0d, table.getRows().get(0).getRowFormat().getHeight());

        for (Cell c : (Iterable<Cell>) doc.getChildNodes(NodeType.CELL, true)) {
            Assert.assertEquals(ParagraphAlignment.CENTER, c.getFirstParagraph().getParagraphFormat().getAlignment());

            for (Run r : c.getFirstParagraph().getRuns()) {
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
        //ExSummary:Shows how to apply border and shading color while building a table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and set a default color/thickness for its borders.
        Table table = builder.startTable();
        table.setBorders(LineStyle.SINGLE, 2.0, Color.BLACK);

        // Create a row with two cells with different background colors.
        builder.insertCell();
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.RED);
        builder.writeln("Row 1, Cell 1.");
        builder.insertCell();
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.GREEN);
        builder.writeln("Row 1, Cell 2.");
        builder.endRow();

        // Reset cell formatting to disable the background colors
        // set a custom border thickness for all new cells created by the builder,
        // then build a second row.
        builder.getCellFormat().clearFormatting();
        builder.getCellFormat().getBorders().getLeft().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getRight().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getTop().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getBottom().setLineWidth(4.0);

        builder.insertCell();
        builder.writeln("Row 2, Cell 1.");
        builder.insertCell();
        builder.writeln("Row 2, Cell 2.");

        doc.save(getArtifactsDir() + "DocumentBuilder.TableBordersAndShading.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.TableBordersAndShading.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

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
        //ExSummary:Shows how to use unit conversion tools while specifying a preferred width for a cell.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.getCellFormat().setPreferredWidth(PreferredWidth.fromPoints(ConvertUtil.inchToPoint(3.0)));
        builder.insertCell();

        Assert.assertEquals(216.0d, table.getFirstRow().getFirstCell().getCellFormat().getPreferredWidth().getValue());
        //ExEnd
    }

    @Test
    public void insertHyperlinkToLocalBookmark() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.StartBookmark
        //ExFor:DocumentBuilder.EndBookmark
        //ExFor:DocumentBuilder.InsertHyperlink
        //ExSummary:Shows how to insert a hyperlink which references a local bookmark.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("Bookmark1");
        builder.write("Bookmarked text. ");
        builder.endBookmark("Bookmark1");
        builder.writeln("Text outside of the bookmark.");

        // Insert a HYPERLINK field that links to the bookmark. We can pass field switches
        // to the "InsertHyperlink" method as part of the argument containing the referenced bookmark's name.
        builder.getFont().setColor(Color.BLUE);
        builder.getFont().setUnderline(Underline.SINGLE);
        builder.insertHyperlink("Link to Bookmark1", "Bookmark1\" \\o \"Hyperlink Tip", true);

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertHyperlinkToLocalBookmark.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertHyperlinkToLocalBookmark.docx");
        FieldHyperlink hyperlink = (FieldHyperlink) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_HYPERLINK, " HYPERLINK \\l \"Bookmark1\" \\o \"Hyperlink Tip\" ", "Link to Bookmark1", hyperlink);
        Assert.assertEquals("Bookmark1", hyperlink.getSubAddress());
        Assert.assertEquals("Hyperlink Tip", hyperlink.getScreenTip());
        Assert.assertTrue(IterableUtils.matchesAny(doc.getRange().getBookmarks(), b -> b.getName().contains("Bookmark1")));
    }

    @Test
    public void cursorPosition() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Hello world!");

        // If the builder's cursor is at the end of the document,
        // there will be no nodes in front of it so that the current node will be null.
        Assert.assertNull(builder.getCurrentNode());

        Assert.assertEquals("Hello world!", builder.getCurrentParagraph().getText().trim());

        // Move to the beginning of the document and place the cursor at an existing node.
        builder.moveToDocumentStart();
        Assert.assertEquals(NodeType.RUN, builder.getCurrentNode().getNodeType());
    }

    @Test
    public void moveTo() throws Exception {
        //ExStart
        //ExFor:Story.LastParagraph
        //ExFor:DocumentBuilder.MoveTo(Node)
        //ExSummary:Shows how to move a DocumentBuilder's cursor position to a specified node.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Run 1. ");

        // The document builder has a cursor, which acts as the part of the document
        // where the builder appends new nodes when we use its document construction methods.
        // This cursor functions in the same way as Microsoft Word's blinking cursor,
        // and it also always ends up immediately after any node that the builder just inserted.
        // To append content to a different part of the document,
        // we can move the cursor to a different node with the "MoveTo" method.
        Assert.assertEquals(doc.getFirstSection().getBody().getLastParagraph(), builder.getCurrentParagraph()); //ExSkip
        builder.moveTo(doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0));
        Assert.assertEquals(doc.getFirstSection().getBody().getFirstParagraph(), builder.getCurrentParagraph()); //ExSkip

        // The cursor is now in front of the node that we moved it to.
        // Adding a second run will insert it in front of the first run.
        builder.writeln("Run 2. ");

        Assert.assertEquals("Run 2. \rRun 1.", doc.getText().trim());

        // Move the cursor to the end of the document to continue appending text to the end as before.
        builder.moveTo(doc.getLastSection().getBody().getLastParagraph());
        builder.writeln("Run 3. ");

        Assert.assertEquals("Run 2. \rRun 1. \rRun 3.", doc.getText().trim());
        Assert.assertEquals(doc.getFirstSection().getBody().getLastParagraph(), builder.getCurrentParagraph()); //ExSkip
        //ExEnd
    }

    @Test
    public void moveToParagraph() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.MoveToParagraph
        //ExSummary:Shows how to move a builder's cursor position to a specified paragraph.
        Document doc = new Document(getMyDir() + "Paragraphs.docx");
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        Assert.assertEquals(22, paragraphs.getCount());

        // Create document builder to edit the document. The builder's cursor,
        // which is the point where it will insert new nodes when we call its document construction methods,
        // is currently at the beginning of the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        Assert.assertEquals(0, paragraphs.indexOf(builder.getCurrentParagraph()));

        // Move that cursor to a different paragraph will place that cursor in front of that paragraph.
        builder.moveToParagraph(2, 0);
        Assert.assertEquals(2, paragraphs.indexOf(builder.getCurrentParagraph())); //ExSkip

        // Any new content that we add will be inserted at that point.
        builder.writeln("This is a new third paragraph. ");
        //ExEnd

        Assert.assertEquals(3, paragraphs.indexOf(builder.getCurrentParagraph()));

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals("This is a new third paragraph.", doc.getFirstSection().getBody().getParagraphs().get(2).getText().trim());
    }

    @Test
    public void moveToCell() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.MoveToCell
        //ExSummary:Shows how to move a document builder's cursor to a cell in a table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an empty 2x2 table.
        builder.startTable();
        builder.insertCell();
        builder.insertCell();
        builder.endRow();
        builder.insertCell();
        builder.insertCell();
        builder.endTable();

        // Because we have ended the table with the EndTable method,
        // the document builder's cursor is currently outside the table.
        // This cursor has the same function as Microsoft Word's blinking text cursor.
        // It can also be moved to a different location in the document using the builder's MoveTo methods.
        // We can move the cursor back inside the table to a specific cell.
        builder.moveToCell(0, 1, 1, 0);
        builder.write("Column 2, cell 2.");

        doc.save(getArtifactsDir() + "DocumentBuilder.MoveToCell.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.MoveToCell.docx");

        Table table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertEquals("Column 2, cell 2.", table.getRows().get(1).getCells().get(1).getText().trim());
    }

    @Test
    public void moveToBookmark() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.MoveToBookmark(String, Boolean, Boolean)
        //ExSummary:Shows how to move a document builder's node insertion point cursor to a bookmark.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // A valid bookmark consists of a BookmarkStart node, a BookmarkEnd node with a
        // matching bookmark name somewhere afterward, and contents enclosed by those nodes.
        builder.startBookmark("MyBookmark");
        builder.write("Hello world! ");
        builder.endBookmark("MyBookmark");

        // There are 4 ways of moving a document builder's cursor to a bookmark.
        // If we are between the BookmarkStart and BookmarkEnd nodes, the cursor will be inside the bookmark.
        // This means that any text added by the builder will become a part of the bookmark.
        // 1 -  Outside of the bookmark, in front of the BookmarkStart node:
        Assert.assertTrue(builder.moveToBookmark("MyBookmark", true, false));
        builder.write("1. ");

        Assert.assertEquals("Hello world! ", doc.getRange().getBookmarks().get("MyBookmark").getText());
        Assert.assertEquals("1. Hello world!", doc.getText().trim());

        // 2 -  Inside the bookmark, right after the BookmarkStart node:
        Assert.assertTrue(builder.moveToBookmark("MyBookmark", true, true));
        builder.write("2. ");

        Assert.assertEquals("2. Hello world! ", doc.getRange().getBookmarks().get("MyBookmark").getText());
        Assert.assertEquals("1. 2. Hello world!", doc.getText().trim());

        // 2 -  Inside the bookmark, right in front of the BookmarkEnd node:
        Assert.assertTrue(builder.moveToBookmark("MyBookmark", false, false));
        builder.write("3. ");

        Assert.assertEquals("2. Hello world! 3. ", doc.getRange().getBookmarks().get("MyBookmark").getText());
        Assert.assertEquals("1. 2. Hello world! 3.", doc.getText().trim());

        // 4 -  Outside of the bookmark, after the BookmarkEnd node:
        Assert.assertTrue(builder.moveToBookmark("MyBookmark", false, true));
        builder.write("4.");

        Assert.assertEquals("2. Hello world! 3. ", doc.getRange().getBookmarks().get("MyBookmark").getText());
        Assert.assertEquals("1. 2. Hello world! 3. 4.", doc.getText().trim());
        //ExEnd
    }

    @Test
    public void buildTable() throws Exception {
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
        //ExFor:AutoFitBehavior
        //ExSummary:Shows how to build a formatted 2x2 table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
        builder.write("Row 1, cell 1.");
        builder.insertCell();
        builder.write("Row 1, cell 2.");
        builder.endRow();

        // While building the table, the document builder will apply its current RowFormat/CellFormat property values
        // to the current row/cell that its cursor is in and any new rows/cells as it creates them.
        Assert.assertEquals(CellVerticalAlignment.CENTER, table.getRows().get(0).getCells().get(0).getCellFormat().getVerticalAlignment());
        Assert.assertEquals(CellVerticalAlignment.CENTER, table.getRows().get(0).getCells().get(1).getCellFormat().getVerticalAlignment());

        builder.insertCell();
        builder.getRowFormat().setHeight(100.0);
        builder.getRowFormat().setHeightRule(HeightRule.EXACTLY);
        builder.getCellFormat().setOrientation(TextOrientation.UPWARD);
        builder.write("Row 2, cell 1.");
        builder.insertCell();
        builder.getCellFormat().setOrientation(TextOrientation.DOWNWARD);
        builder.write("Row 2, cell 2.");
        builder.endRow();
        builder.endTable();

        // Previously added rows and cells are not retroactively affected by changes to the builder's formatting.
        Assert.assertEquals(0.0, table.getRows().get(0).getRowFormat().getHeight());
        Assert.assertEquals(HeightRule.AUTO, table.getRows().get(0).getRowFormat().getHeightRule());
        Assert.assertEquals(100.0, table.getRows().get(1).getRowFormat().getHeight());
        Assert.assertEquals(HeightRule.EXACTLY, table.getRows().get(1).getRowFormat().getHeightRule());
        Assert.assertEquals(TextOrientation.UPWARD, table.getRows().get(1).getCells().get(0).getCellFormat().getOrientation());
        Assert.assertEquals(TextOrientation.DOWNWARD, table.getRows().get(1).getCells().get(1).getCellFormat().getOrientation());

        doc.save(getArtifactsDir() + "DocumentBuilder.BuildTable.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.BuildTable.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertEquals(2, table.getRows().getCount());
        Assert.assertEquals(2, table.getRows().get(0).getCells().getCount());
        Assert.assertEquals(2, table.getRows().get(1).getCells().getCount());

        Assert.assertEquals(0.0, table.getRows().get(0).getRowFormat().getHeight());
        Assert.assertEquals(HeightRule.AUTO, table.getRows().get(0).getRowFormat().getHeightRule());
        Assert.assertEquals(100.0, table.getRows().get(1).getRowFormat().getHeight());
        Assert.assertEquals(HeightRule.EXACTLY, table.getRows().get(1).getRowFormat().getHeightRule());

        Assert.assertEquals("Row 1, cell 1.", table.getRows().get(0).getCells().get(0).getText().trim());
        Assert.assertEquals(CellVerticalAlignment.CENTER, table.getRows().get(0).getCells().get(0).getCellFormat().getVerticalAlignment());

        Assert.assertEquals("Row 1, cell 2.", table.getRows().get(0).getCells().get(1).getText().trim());

        Assert.assertEquals("Row 2, cell 1.", table.getRows().get(1).getCells().get(0).getText().trim());
        Assert.assertEquals(TextOrientation.UPWARD, table.getRows().get(1).getCells().get(0).getCellFormat().getOrientation());

        Assert.assertEquals("Row 2, cell 2.", table.getRows().get(1).getCells().get(1).getText().trim());
        Assert.assertEquals(TextOrientation.DOWNWARD, table.getRows().get(1).getCells().get(1).getCellFormat().getOrientation());
    }

    @Test
    public void tableCellVerticalRotatedFarEastTextOrientation() throws Exception {
        Document doc = new Document(getMyDir() + "Rotated cell text.docx");

        Table table = doc.getFirstSection().getBody().getTables().get(0);
        Cell cell = table.getFirstRow().getFirstCell();

        Assert.assertEquals(cell.getCellFormat().getOrientation(), TextOrientation.VERTICAL_ROTATED_FAR_EAST);

        doc = DocumentHelper.saveOpen(doc);

        table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        cell = table.getFirstRow().getFirstCell();

        Assert.assertEquals(cell.getCellFormat().getOrientation(), TextOrientation.VERTICAL_ROTATED_FAR_EAST);
    }

    @Test
    public void insertFloatingImage() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert an image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // There are two ways of using a document builder to source an image and then insert it as a floating shape.
        // 1 -  From a file in the local file system:
        builder.insertImage(getImageDir() + "Transparent background logo.png", RelativeHorizontalPosition.MARGIN, 100.0,
                RelativeVerticalPosition.MARGIN, 0.0, 200.0, 200.0, WrapType.SQUARE);

        // 2 -  From a URL:
        builder.insertImage(getAsposelogoUri().toString(), RelativeHorizontalPosition.MARGIN, 100.0,
                RelativeVerticalPosition.MARGIN, 250.0, 200.0, 200.0, WrapType.SQUARE);

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertFloatingImage.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertFloatingImage.docx");
        Shape image = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.PNG, image);
        Assert.assertEquals(100.0d, image.getLeft());
        Assert.assertEquals(0.0d, image.getTop());
        Assert.assertEquals(200.0d, image.getWidth());
        Assert.assertEquals(200.0d, image.getHeight());
        Assert.assertEquals(WrapType.SQUARE, image.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.MARGIN, image.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.MARGIN, image.getRelativeVerticalPosition());

        image = (Shape) doc.getChild(NodeType.SHAPE, 1, true);

        TestUtil.verifyImageInShape(320, 320, ImageType.PNG, image);
        Assert.assertEquals(100.0d, image.getLeft());
        Assert.assertEquals(250.0d, image.getTop());
        Assert.assertEquals(200.0d, image.getWidth());
        Assert.assertEquals(200.0d, image.getHeight());
        Assert.assertEquals(WrapType.SQUARE, image.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.MARGIN, image.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.MARGIN, image.getRelativeVerticalPosition());
    }

    @Test
    public void insertImageOriginalSize() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert an image from the local file system into a document while preserving its dimensions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The InsertImage method creates a floating shape with the passed image in its image data.
        // We can specify the dimensions of the shape can be passing them to this method.
        Shape imageShape = builder.insertImage(getImageDir() + "Logo.jpg", RelativeHorizontalPosition.MARGIN, 0.0,
                RelativeVerticalPosition.MARGIN, 0.0, -1, -1, WrapType.SQUARE);

        // Passing negative values as the intended dimensions will automatically define
        // the shape's dimensions based on the dimensions of its image.
        Assert.assertEquals(300.0d, imageShape.getWidth());
        Assert.assertEquals(300.0d, imageShape.getHeight());

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertImageOriginalSize.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertImageOriginalSize.docx");
        imageShape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, imageShape);
        Assert.assertEquals(0.0d, imageShape.getLeft());
        Assert.assertEquals(0.0d, imageShape.getTop());
        Assert.assertEquals(300.0d, imageShape.getWidth());
        Assert.assertEquals(300.0d, imageShape.getHeight());
        Assert.assertEquals(WrapType.SQUARE, imageShape.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.MARGIN, imageShape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.MARGIN, imageShape.getRelativeVerticalPosition());
    }

    @Test
    public void insertTextInput() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertTextInput
        //ExSummary:Shows how to insert a text input form field into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a form that prompts the user to enter text.
        builder.insertTextInput("TextInput", TextFormFieldType.REGULAR, "", "Enter your text here", 0);

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertTextInput.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertTextInput.docx");
        FormField formField = doc.getRange().getFormFields().get(0);

        Assert.assertTrue(formField.getEnabled());
        Assert.assertEquals("TextInput", formField.getName());
        Assert.assertEquals(0, formField.getMaxLength());
        Assert.assertEquals("Enter your text here", formField.getResult());
        Assert.assertEquals(FieldType.FIELD_FORM_TEXT_INPUT, formField.getType());
        Assert.assertEquals("", formField.getTextInputFormat());
        Assert.assertEquals(TextFormFieldType.REGULAR, formField.getTextInputType());
    }

    @Test
    public void insertComboBox() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertComboBox
        //ExSummary:Shows how to insert a combo box form field into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a form that prompts the user to pick one of the items from the menu.
        builder.write("Pick a fruit: ");
        String[] items = {"Apple", "Banana", "Cherry"};
        builder.insertComboBox("DropDown", items, 0);

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertComboBox.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertComboBox.docx");
        FormField formField = doc.getRange().getFormFields().get(0);

        Assert.assertTrue(formField.getEnabled());
        Assert.assertEquals("DropDown", formField.getName());
        Assert.assertEquals(0, formField.getDropDownSelectedIndex());
        Assert.assertEquals(items.length, formField.getDropDownItems().getCount());
        Assert.assertEquals(FieldType.FIELD_FORM_DROP_DOWN, formField.getType());
    }

    @Test(description = "WORDSNET-16868, WORDSJAVA-2406", enabled = false)
    public void signatureLineProviderId() throws Exception {
        //ExStart
        //ExFor:SignatureLine.IsSigned
        //ExFor:SignatureLine.IsValid
        //ExFor:SignatureLine.ProviderId
        //ExFor:SignatureLineOptions.ShowDate
        //ExFor:SignatureLineOptions.Email
        //ExFor:SignatureLineOptions.DefaultInstructions
        //ExFor:SignatureLineOptions.Instructions
        //ExFor:SignatureLineOptions.AllowComments
        //ExFor:DocumentBuilder.InsertSignatureLine(SignatureLineOptions)
        //ExFor:SignOptions.ProviderId
        //ExSummary:Shows how to sign a document with a personal certificate and a signature line.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        SignatureLineOptions signatureLineOptions = new SignatureLineOptions();
        signatureLineOptions.setSigner("vderyushev");
        signatureLineOptions.setSignerTitle("QA");
        signatureLineOptions.setEmail("vderyushev@aspose.com");
        signatureLineOptions.setShowDate(true);
        signatureLineOptions.setDefaultInstructions(false);
        signatureLineOptions.setInstructions("Please sign here.");
        signatureLineOptions.setAllowComments(true);

        SignatureLine signatureLine = builder.insertSignatureLine(signatureLineOptions).getSignatureLine();
        signatureLine.setProviderId(UUID.fromString("CF5A7BB4-8F3C-4756-9DF6-BEF7F13259A2"));

        Assert.assertFalse(signatureLine.isSigned());
        Assert.assertFalse(signatureLine.isValid());

        doc.save(getArtifactsDir() + "DocumentBuilder.SignatureLineProviderId.docx");

        Date currentDate = new Date();

        SignOptions signOptions = new SignOptions();
        signOptions.setSignatureLineId(signatureLine.getId());
        signOptions.setProviderId(signatureLine.getProviderId());
        signOptions.setComments("Document was signed by vderyushev");
        signOptions.setSignTime(currentDate);

        CertificateHolder certHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        DigitalSignatureUtil.sign(getArtifactsDir() + "DocumentBuilder.SignatureLineProviderId.docx",
                getArtifactsDir() + "DocumentBuilder.SignatureLineProviderId.Signed.docx", certHolder, signOptions);

        // Re-open our saved document, and verify that the "IsSigned" and "IsValid" properties both equal "true",
        // indicating that the signature line contains a signature.
        doc = new Document(getArtifactsDir() + "DocumentBuilder.SignatureLineProviderId.Signed.docx");
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        signatureLine = shape.getSignatureLine();

        Assert.assertTrue(signatureLine.isSigned());
        Assert.assertTrue(signatureLine.isValid());
        //ExEnd

        Assert.assertEquals("vderyushev", signatureLine.getSigner());
        Assert.assertEquals("QA", signatureLine.getSignerTitle());
        Assert.assertEquals("vderyushev@aspose.com", signatureLine.getEmail());
        Assert.assertTrue(signatureLine.getShowDate());
        Assert.assertFalse(signatureLine.getDefaultInstructions());
        Assert.assertEquals("Please sign here.", signatureLine.getInstructions());
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
    public void signatureLineInline() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertSignatureLine(SignatureLineOptions, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, WrapType)
        //ExSummary:Shows how to insert an inline signature line into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        SignatureLineOptions options = new SignatureLineOptions();
        options.setSigner("John Doe");
        options.setSignerTitle("Manager");
        options.setEmail("johndoe@aspose.com");
        options.setShowDate(true);
        options.setDefaultInstructions(false);
        options.setInstructions("Please sign here.");
        options.setAllowComments(true);

        builder.insertSignatureLine(options, RelativeHorizontalPosition.RIGHT_MARGIN, 2.0,
                RelativeVerticalPosition.PAGE, 3.0, WrapType.INLINE);

        // The signature line can be signed in Microsoft Word by double clicking it.
        doc.save(getArtifactsDir() + "DocumentBuilder.SignatureLineInline.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.SignatureLineInline.docx");

        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        SignatureLine signatureLine = shape.getSignatureLine();

        Assert.assertEquals(signatureLine.getSigner(), "John Doe");
        Assert.assertEquals(signatureLine.getSignerTitle(), "Manager");
        Assert.assertEquals(signatureLine.getEmail(), "johndoe@aspose.com");
        Assert.assertEquals(signatureLine.getShowDate(), true);
        Assert.assertEquals(signatureLine.getDefaultInstructions(), false);
        Assert.assertEquals(signatureLine.getInstructions(), "Please sign here.");
        Assert.assertEquals(signatureLine.getAllowComments(), true);
        Assert.assertEquals(signatureLine.isSigned(), false);
        Assert.assertEquals(signatureLine.isValid(), false);
    }

    @Test
    public void setParagraphFormatting() throws Exception {
        //ExStart
        //ExFor:ParagraphFormat.RightIndent
        //ExFor:ParagraphFormat.LeftIndent
        //ExSummary:Shows how to configure paragraph formatting to create off-center text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Center all text that the document builder writes, and set up indents.
        // The indent configuration below will create a body of text that will sit asymmetrically on the page.
        // The "center" that we align the text to will be the middle of the body of text, not the middle of the page.
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();
        paragraphFormat.setAlignment(ParagraphAlignment.CENTER);
        paragraphFormat.setLeftIndent(100.0);
        paragraphFormat.setRightIndent(50.0);
        paragraphFormat.setSpaceAfter(25.0);

        builder.writeln(
                "This paragraph demonstrates how left and right indentation affects word wrapping.");
        builder.writeln(
                "The space between the above paragraph and this one depends on the DocumentBuilder's paragraph format.");

        doc.save(getArtifactsDir() + "DocumentBuilder.SetParagraphFormatting.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.SetParagraphFormatting.docx");

        for (Paragraph paragraph : doc.getFirstSection().getBody().getParagraphs()) {
            Assert.assertEquals(ParagraphAlignment.CENTER, paragraph.getParagraphFormat().getAlignment());
            Assert.assertEquals(100.0d, paragraph.getParagraphFormat().getLeftIndent());
            Assert.assertEquals(50.0d, paragraph.getParagraphFormat().getRightIndent());
            Assert.assertEquals(25.0d, paragraph.getParagraphFormat().getSpaceAfter());
        }
    }

    @Test
    public void setCellFormatting() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.CellFormat
        //ExFor:CellFormat.Width
        //ExFor:CellFormat.LeftPadding
        //ExFor:CellFormat.RightPadding
        //ExFor:CellFormat.TopPadding
        //ExFor:CellFormat.BottomPadding
        //ExFor:DocumentBuilder.StartTable
        //ExFor:DocumentBuilder.EndTable
        //ExSummary:Shows how to format cells with a document builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Row 1, cell 1.");

        // Insert a second cell, and then configure cell text padding options.
        // The builder will apply these settings at its current cell, and any new cells creates afterwards.
        builder.insertCell();

        CellFormat cellFormat = builder.getCellFormat();
        cellFormat.setWidth(250.0);
        cellFormat.setLeftPadding(30.0);
        cellFormat.setRightPadding(30.0);
        cellFormat.setTopPadding(30.0);
        cellFormat.setBottomPadding(30.0);

        builder.write("Row 1, cell 2.");
        builder.endRow();
        builder.endTable();

        // The first cell was unaffected by the padding reconfiguration, and still holds the default values.
        Assert.assertEquals(0.0d, table.getFirstRow().getCells().get(0).getCellFormat().getWidth());
        Assert.assertEquals(5.4d, table.getFirstRow().getCells().get(0).getCellFormat().getLeftPadding());
        Assert.assertEquals(5.4d, table.getFirstRow().getCells().get(0).getCellFormat().getRightPadding());
        Assert.assertEquals(0.0d, table.getFirstRow().getCells().get(0).getCellFormat().getTopPadding());
        Assert.assertEquals(0.0d, table.getFirstRow().getCells().get(0).getCellFormat().getBottomPadding());

        Assert.assertEquals(250.0d, table.getFirstRow().getCells().get(1).getCellFormat().getWidth());
        Assert.assertEquals(30.0d, table.getFirstRow().getCells().get(1).getCellFormat().getLeftPadding());
        Assert.assertEquals(30.0d, table.getFirstRow().getCells().get(1).getCellFormat().getRightPadding());
        Assert.assertEquals(30.0d, table.getFirstRow().getCells().get(1).getCellFormat().getTopPadding());
        Assert.assertEquals(30.0d, table.getFirstRow().getCells().get(1).getCellFormat().getBottomPadding());

        // The first cell will still grow in the output document to match the size of its neighboring cell.
        doc.save(getArtifactsDir() + "DocumentBuilder.SetCellFormatting.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.SetCellFormatting.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertEquals(157.0d, table.getFirstRow().getCells().get(0).getCellFormat().getWidth());
        Assert.assertEquals(5.4d, table.getFirstRow().getCells().get(0).getCellFormat().getLeftPadding());
        Assert.assertEquals(5.4d, table.getFirstRow().getCells().get(0).getCellFormat().getRightPadding());
        Assert.assertEquals(0.0d, table.getFirstRow().getCells().get(0).getCellFormat().getTopPadding());
        Assert.assertEquals(0.0d, table.getFirstRow().getCells().get(0).getCellFormat().getBottomPadding());

        Assert.assertEquals(310.0d, table.getFirstRow().getCells().get(1).getCellFormat().getWidth());
        Assert.assertEquals(30.0d, table.getFirstRow().getCells().get(1).getCellFormat().getLeftPadding());
        Assert.assertEquals(30.0d, table.getFirstRow().getCells().get(1).getCellFormat().getRightPadding());
        Assert.assertEquals(30.0d, table.getFirstRow().getCells().get(1).getCellFormat().getTopPadding());
        Assert.assertEquals(30.0d, table.getFirstRow().getCells().get(1).getCellFormat().getBottomPadding());
    }

    @Test
    public void setRowFormatting() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.RowFormat
        //ExFor:HeightRule
        //ExFor:RowFormat.Height
        //ExFor:RowFormat.HeightRule
        //ExSummary:Shows how to format rows with a document builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Row 1, cell 1.");

        // Start a second row, and then configure its height. The builder will apply these settings to
        // its current row, as well as any new rows it creates afterwards.
        builder.endRow();

        RowFormat rowFormat = builder.getRowFormat();
        rowFormat.setHeight(100.0);
        rowFormat.setHeightRule(HeightRule.EXACTLY);

        builder.insertCell();
        builder.write("Row 2, cell 1.");
        builder.endTable();

        // The first row was unaffected by the padding reconfiguration and still holds the default values.
        Assert.assertEquals(0.0d, table.getRows().get(0).getRowFormat().getHeight());
        Assert.assertEquals(HeightRule.AUTO, table.getRows().get(0).getRowFormat().getHeightRule());

        Assert.assertEquals(100.0d, table.getRows().get(1).getRowFormat().getHeight());
        Assert.assertEquals(HeightRule.EXACTLY, table.getRows().get(1).getRowFormat().getHeightRule());

        doc.save(getArtifactsDir() + "DocumentBuilder.SetRowFormatting.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.SetRowFormatting.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertEquals(0.0d, table.getRows().get(0).getRowFormat().getHeight());
        Assert.assertEquals(HeightRule.AUTO, table.getRows().get(0).getRowFormat().getHeightRule());

        Assert.assertEquals(100.0d, table.getRows().get(1).getRowFormat().getHeight());
        Assert.assertEquals(HeightRule.EXACTLY, table.getRows().get(1).getRowFormat().getHeightRule());
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

        // Insert some text and mark it with a footnote with the IsAuto property set to "true" by default,
        // so the marker seen in the body text will be auto-numbered at "1",
        // and the footnote will appear at the bottom of the page.
        builder.write("This text will be referenced by a footnote.");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote comment regarding referenced text.");

        // Insert more text and mark it with an endnote with a custom reference mark,
        // which will be used in place of the number "2" and set "IsAuto" to false.
        builder.write("This text will be referenced by an endnote.");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Endnote comment regarding referenced text.", "CustomMark");

        // Footnotes always appear at the bottom of their referenced text,
        // so this page break will not affect the footnote.
        // On the other hand, endnotes are always at the end of the document
        // so that this page break will push the endnote down to the next page.
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
    public void applyBordersAndShading() throws Exception {
        //ExStart
        //ExFor:BorderCollection.Item(BorderType)
        //ExFor:Shading
        //ExFor:TextureIndex
        //ExFor:ParagraphFormat.Shading
        //ExFor:Shading.Texture
        //ExFor:Shading.BackgroundPatternColor
        //ExFor:Shading.ForegroundPatternColor
        //ExSummary:Shows how to decorate text with borders and shading.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        BorderCollection borders = builder.getParagraphFormat().getBorders();
        borders.setDistanceFromText(20.0);
        borders.getByBorderType(BorderType.LEFT).setLineStyle(LineStyle.DOUBLE);
        borders.getByBorderType(BorderType.RIGHT).setLineStyle(LineStyle.DOUBLE);
        borders.getByBorderType(BorderType.TOP).setLineStyle(LineStyle.DOUBLE);
        borders.getByBorderType(BorderType.BOTTOM).setLineStyle(LineStyle.DOUBLE);

        Shading shading = builder.getParagraphFormat().getShading();
        shading.setTexture(TextureIndex.TEXTURE_DIAGONAL_CROSS);
        shading.setBackgroundPatternColor(new Color(240, 128, 128));  // Light Coral
        shading.setForegroundPatternColor(new Color(255, 160, 122));  // Light Salmon

        builder.write("This paragraph is formatted with a double border and shading.");
        doc.save(getArtifactsDir() + "DocumentBuilder.ApplyBordersAndShading.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.ApplyBordersAndShading.docx");
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

        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Row 1, cell 1.");
        builder.insertCell();
        builder.write("Row 1, cell 2.");
        builder.endRow();
        builder.insertCell();
        builder.write("Row 2, cell 1.");
        builder.insertCell();
        builder.write("Row 2, cell 2.");
        builder.endTable();

        Assert.assertEquals(2, table.getRows().getCount());

        // Delete the first row of the first table in the document.
        builder.deleteRow(0, 0);

        Assert.assertEquals(1, table.getRows().getCount());
        Assert.assertEquals("Row 2, cell 1.\u0007Row 2, cell 2.", table.getText().trim());
        //ExEnd
    }

    @Test (dataProvider = "appendDocumentAndResolveStylesDataProvider")
    public void appendDocumentAndResolveStyles(boolean keepSourceNumbering) throws Exception
    {
        //ExStart
        //ExFor:Document.AppendDocument(Document, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how to manage list style clashes while appending a document.
        // Load a document with text in a custom style and clone it.
        Document srcDoc = new Document(getMyDir() + "Custom list numbering.docx");
        Document dstDoc = srcDoc.deepClone();

        // We now have two documents, each with an identical style named "CustomStyle".
        // Change the text color for one of the styles to set it apart from the other.
        dstDoc.getStyles().get("CustomStyle").getFont().setColor(Color.RED);

        // If there is a clash of list styles, apply the list format of the source document.
        // Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
        // Set the "KeepSourceNumbering" property to "true" import all clashing
        // list style numbering with the same appearance that it had in the source document.
        ImportFormatOptions options = new ImportFormatOptions();
        options.setKeepSourceNumbering(keepSourceNumbering);

        // Joining two documents that have different styles that share the same name causes a style clash.
        // We can specify an import format mode while appending documents to resolve this clash.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_DIFFERENT_STYLES, options);
        dstDoc.updateListLabels();

        dstDoc.save(getArtifactsDir() + "DocumentBuilder.AppendDocumentAndResolveStyles.docx");
        //ExEnd
    }

	@DataProvider(name = "appendDocumentAndResolveStylesDataProvider")
	public static Object[][] appendDocumentAndResolveStylesDataProvider() {
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "insertDocumentAndResolveStylesDataProvider")
    public void insertDocumentAndResolveStyles(boolean keepSourceNumbering) throws Exception
    {
        //ExStart
        //ExFor:Document.AppendDocument(Document, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how to manage list style clashes while inserting a document.
        Document dstDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(dstDoc);
        builder.insertBreak(BreakType.PARAGRAPH_BREAK);

        dstDoc.getLists().add(ListTemplate.NUMBER_DEFAULT);
        List list = dstDoc.getLists().get(0);

        builder.getListFormat().setList(list);

        for (int i = 1; i <= 15; i++)
            builder.write(MessageFormat.format("List Item {0}\n", i));

        Document attachDoc = (Document)dstDoc.deepClone(true);

        // If there is a clash of list styles, apply the list format of the source document.
        // Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
        // Set the "KeepSourceNumbering" property to "true" import all clashing
        // list style numbering with the same appearance that it had in the source document.
        ImportFormatOptions importOptions = new ImportFormatOptions();
        importOptions.setKeepSourceNumbering(keepSourceNumbering);

        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.insertDocument(attachDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING, importOptions);

        dstDoc.save(getArtifactsDir() + "DocumentBuilder.InsertDocumentAndResolveStyles.docx");
        //ExEnd
    }

	@DataProvider(name = "insertDocumentAndResolveStylesDataProvider")
	public static Object[][] insertDocumentAndResolveStylesDataProvider() {
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "loadDocumentWithListNumberingDataProvider")
    public void loadDocumentWithListNumbering(boolean keepSourceNumbering) throws Exception
    {
        //ExStart
        //ExFor:Document.AppendDocument(Document, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how to manage list style clashes while appending a clone of a document to itself.
        Document srcDoc = new Document(getMyDir() + "List item.docx");
        Document dstDoc = new Document(getMyDir() + "List item.docx");

        // If there is a clash of list styles, apply the list format of the source document.
        // Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
        // Set the "KeepSourceNumbering" property to "true" import all clashing
        // list style numbering with the same appearance that it had in the source document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);
        builder.moveToDocumentEnd();
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);

        ImportFormatOptions options = new ImportFormatOptions();
        options.setKeepSourceNumbering(keepSourceNumbering);
        builder.insertDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING, options);

        dstDoc.updateListLabels();
        //ExEnd
    }

	@DataProvider(name = "loadDocumentWithListNumberingDataProvider")
	public static Object[][] loadDocumentWithListNumberingDataProvider() {
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "ignoreTextBoxesDataProvider")
    public void ignoreTextBoxes(boolean ignoreTextBoxes) throws Exception
    {
        //ExStart
        //ExFor:ImportFormatOptions.IgnoreTextBoxes
        //ExSummary:Shows how to manage text box formatting while appending a document.
        // Create a document that will have nodes from another document inserted into it.
        Document dstDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        builder.writeln("Hello world!");

        // Create another document with a text box, which we will import into the first document.
        Document srcDoc = new Document();
        builder = new DocumentBuilder(srcDoc);

        Shape textBox = builder.insertShape(ShapeType.TEXT_BOX, 300.0, 100.0);
        builder.moveTo(textBox.getFirstParagraph());
        builder.getParagraphFormat().getStyle().getFont().setName("Courier New");
        builder.getParagraphFormat().getStyle().getFont().setSize(24.0);
        builder.write("Textbox contents");

        // Set a flag to specify whether to clear or preserve text box formatting
        // while importing them to other documents.
        ImportFormatOptions importFormatOptions = new ImportFormatOptions();
        importFormatOptions.setIgnoreTextBoxes(ignoreTextBoxes);

        // Import the text box from the source document into the destination document,
        // and then verify whether we have preserved the styling of its text contents.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING, importFormatOptions);
        Shape importedTextBox = (Shape) importer.importNode(textBox, true);
        dstDoc.getFirstSection().getBody().getParagraphs().get(1).appendChild(importedTextBox);

        if (ignoreTextBoxes) {
            Assert.assertEquals(12.0d, importedTextBox.getFirstParagraph().getRuns().get(0).getFont().getSize());
            Assert.assertEquals("Times New Roman", importedTextBox.getFirstParagraph().getRuns().get(0).getFont().getName());
        } else {
            Assert.assertEquals(24.0d, importedTextBox.getFirstParagraph().getRuns().get(0).getFont().getSize());
            Assert.assertEquals("Courier New", importedTextBox.getFirstParagraph().getRuns().get(0).getFont().getName());
        }

        dstDoc.save(getArtifactsDir() + "DocumentBuilder.IgnoreTextBoxes.docx");
        //ExEnd
    }

    @DataProvider(name = "ignoreTextBoxesDataProvider")
    public static Object[][] ignoreTextBoxesDataProvider() {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }

    @Test(dataProvider = "moveToFieldDataProvider")
    public void moveToField(boolean moveCursorToAfterTheField) throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.MoveToField
        //ExSummary:Shows how to move a document builder's node insertion point cursor to a specific field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a field using the DocumentBuilder and add a run of text after it.
        Field field = builder.insertField(" AUTHOR \"John Doe\" ");

        // The builder's cursor is currently at end of the document.
        Assert.assertNull(builder.getCurrentNode());

        // Move the cursor to the field while specifying whether to place that cursor before or after the field.
        builder.moveToField(field, moveCursorToAfterTheField);

        // Note that the cursor is outside of the field in both cases.
        // This means that we cannot edit the field using the builder like this.
        // To edit a field, we can use the builder's MoveTo method on a field's FieldStart
        // or FieldSeparator node to place the cursor inside.
        if (moveCursorToAfterTheField) {
            Assert.assertNull(builder.getCurrentNode());
            builder.write(" Text immediately after the field.");

            Assert.assertEquals("AUTHOR \"John Doe\" \u0014John Doe\u0015 Text immediately after the field.",
                    doc.getText().trim());
        } else {
            Assert.assertEquals(field.getStart(), builder.getCurrentNode());
            builder.write("Text immediately before the field. ");

            Assert.assertEquals("Text immediately before the field. \u0013 AUTHOR \"John Doe\" \u0014John Doe",
                    doc.getText().trim());
        }
        //ExEnd
    }

    @DataProvider(name = "moveToFieldDataProvider")
    public static Object[][] moveToFieldDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void insertOleObjectException() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Assert.assertThrows(RuntimeException.class, () -> builder.insertOleObject("", "checkbox", false, true, null));
    }

    @Test
    public void insertPieChart() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertChart(ChartType, Double, Double)
        //ExSummary:Shows how to insert a pie chart into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Chart chart = builder.insertChart(ChartType.PIE, ConvertUtil.pixelToPoint(300.0),
                ConvertUtil.pixelToPoint(300.0)).getChart();
        Assert.assertEquals(225.0d, ConvertUtil.pixelToPoint(300.0)); //ExSkip
        chart.getSeries().clear();
        chart.getSeries().add("My fruit",
                new String[]{"Apples", "Bananas", "Cherries"},
                new double[]{1.3, 2.2, 1.5});

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertPieChart.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertPieChart.docx");
        Shape chartShape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals("Chart Title", chartShape.getChart().getTitle().getText());
        Assert.assertEquals(225.0d, chartShape.getWidth());
        Assert.assertEquals(225.0d, chartShape.getHeight());
    }

    @Test
    public void insertChartRelativePosition() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertChart(ChartType, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to specify position and wrapping while inserting a chart.
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
        //ExFor:Field.Result
        //ExFor:Field.GetFieldCode
        //ExFor:Field.Type
        //ExFor:FieldType
        //ExSummary:Shows how to insert a field into a document using a field code.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Field dateField = builder.insertField("DATE \\* MERGEFORMAT");

        Assert.assertEquals(FieldType.FIELD_DATE, dateField.getType());
        Assert.assertEquals("DATE \\* MERGEFORMAT", dateField.getFieldCode());
        //ExEnd			
    }

    @Test(dataProvider = "insertFieldAndUpdateDataProvider")
    public void insertFieldAndUpdate(boolean updateInsertedFieldsImmediately) throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertField(FieldType, Boolean)
        //ExFor:Field.Update
        //ExSummary:Shows how to insert a field into a document using FieldType.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two fields while passing a flag which determines whether to update them as the builder inserts them.
        // In some cases, updating fields could be computationally expensive, and it may be a good idea to defer the update.
        doc.getBuiltInDocumentProperties().setAuthor("John Doe");
        builder.write("This document was written by ");
        builder.insertField(FieldType.FIELD_AUTHOR, updateInsertedFieldsImmediately);

        builder.insertParagraph();
        builder.write("\nThis is page ");
        builder.insertField(FieldType.FIELD_PAGE, updateInsertedFieldsImmediately);

        Assert.assertEquals(" AUTHOR ", doc.getRange().getFields().get(0).getFieldCode());
        Assert.assertEquals(" PAGE ", doc.getRange().getFields().get(1).getFieldCode());

        if (updateInsertedFieldsImmediately) {
            Assert.assertEquals("John Doe", doc.getRange().getFields().get(0).getResult());
            Assert.assertEquals("1", doc.getRange().getFields().get(1).getResult());
        } else {
            Assert.assertEquals("", doc.getRange().getFields().get(0).getResult());
            Assert.assertEquals("", doc.getRange().getFields().get(1).getResult());

            // We will need to update these fields using the update methods manually.
            doc.getRange().getFields().get(0).update();

            Assert.assertEquals("John Doe", doc.getRange().getFields().get(0).getResult());

            doc.updateFields();

            Assert.assertEquals("1", doc.getRange().getFields().get(1).getResult());
        }
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals("This document was written by \u0013 AUTHOR \u0014John Doe\u0015" +
                "\r\rThis is page \u0013 PAGE \u00141", doc.getText().trim());

        TestUtil.verifyField(FieldType.FIELD_AUTHOR, " AUTHOR ", "John Doe", doc.getRange().getFields().get(0));
        TestUtil.verifyField(FieldType.FIELD_PAGE, " PAGE ", "1", doc.getRange().getFields().get(1));
    }

    @DataProvider(name = "insertFieldAndUpdateDataProvider")
    public static Object[][] insertFieldAndUpdateDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    //ExStart
    //ExFor:IFieldResultFormatter
    //ExFor:IFieldResultFormatter.Format(Double, GeneralFormat)
    //ExFor:IFieldResultFormatter.Format(String, GeneralFormat)
    //ExFor:IFieldResultFormatter.FormatDateTime(DateTime, String, CalendarType)
    //ExFor:IFieldResultFormatter.FormatNumeric(Double, String)
    //ExFor:FieldOptions.ResultFormatter
    //ExFor:CalendarType
    //ExSummary:Shows how to automatically apply a custom format to field results as the fields are updated.
    @Test //ExSkip
    public void fieldResultFormatting() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldResultFormatter formatter = new FieldResultFormatter("$%d", "Date: %tb", "Item # %s:");
        doc.getFieldOptions().setResultFormatter(formatter);

        // Our field result formatter applies a custom format to newly created fields of three types of formats.
        // Field result formatters apply new formatting to fields as they are updated,
        // which happens as soon as we create them using this InsertField method overload.
        // 1 -  Numeric:
        builder.insertField(" = 2 + 3 \\# $###");

        Assert.assertEquals("$5", doc.getRange().getFields().get(0).getResult());
        Assert.assertEquals(1, formatter.countFormatInvocations(FieldResultFormatter.FormatInvocationType.NUMERIC));

        // 2 -  Date/time:
        builder.insertField("DATE \\@ \"d MMMM yyyy\"");

        Assert.assertTrue(doc.getRange().getFields().get(1).getResult().startsWith("Date: "));
        Assert.assertEquals(1, formatter.countFormatInvocations(FieldResultFormatter.FormatInvocationType.DATE_TIME));

        // 3 -  General:
        builder.insertField("QUOTE \"2\" \\* Ordinal");

        Assert.assertEquals("Item # 2:", doc.getRange().getFields().get(2).getResult());
        Assert.assertEquals(1, formatter.countFormatInvocations(FieldResultFormatter.FormatInvocationType.GENERAL));

        formatter.printFormatInvocations();
    }

    /// <summary>
    /// When fields with formatting are updated, this formatter will override their formatting
    /// with a custom format, while tracking every invocation.
    /// </summary>
    private static class FieldResultFormatter implements IFieldResultFormatter {
        public FieldResultFormatter(String numberFormat, String dateFormat, String generalFormat) {
            mNumberFormat = numberFormat;
            mDateFormat = dateFormat;
            mGeneralFormat = generalFormat;
        }

        public String formatNumeric(double value, String format) {
            if (mNumberFormat.isEmpty())
                return null;

            String newValue = String.format(mNumberFormat, (long) value);
            mFormatInvocations.add(new FormatInvocation(FormatInvocationType.NUMERIC, value, format, newValue));
            return newValue;
        }

        public String formatDateTime(Date value, String format, int calendarType) {
            if (mDateFormat.isEmpty())
                return null;

            String newValue = String.format(mDateFormat, value);
            mFormatInvocations.add(new FormatInvocation(FormatInvocationType.DATE_TIME, MessageFormat.format("{0} ({1})", value, calendarType), format, newValue));
            return newValue;
        }

        public String format(String value, int format) {
            return format((Object) value, format);
        }

        public String format(double value, int format) {
            return format((Object) value, format);
        }

        private String format(Object value, int format) {
            if (mGeneralFormat.isEmpty())
                return null;

            String newValue = String.format(mGeneralFormat, new DecimalFormat("#.####").format(value));
            mFormatInvocations.add(new FormatInvocation(FormatInvocationType.GENERAL, value, GeneralFormat.toString(format), newValue));
            return newValue;
        }

        public int countFormatInvocations(int formatInvocationType) {
            if (formatInvocationType == FormatInvocationType.ALL)
                return getFormatInvocations().size();

            return (int) IterableUtils.countMatches(getFormatInvocations(), i -> i.getFormatInvocationType() == formatInvocationType);
        }

        public void printFormatInvocations() {
            for (FormatInvocation f : getFormatInvocations())
                System.out.println(MessageFormat.format("Invocation type:\t{0}\n" +
                        "\tOriginal value:\t\t{1}\n" +
                        "\tOriginal format:\t{2}\n" +
                        "\tNew value:\t\t\t{3}\n", f.getFormatInvocationType(), f.getValue(), f.getOriginalFormat(), f.getNewValue()));
        }

        private final String mNumberFormat;
        private final String mDateFormat;
        private final String mGeneralFormat;

        private ArrayList<FormatInvocation> getFormatInvocations() {
            return mFormatInvocations;
        }

        private final ArrayList<FormatInvocation> mFormatInvocations = new ArrayList<>();

        private static class FormatInvocation {
            public int getFormatInvocationType() {
                return mFormatInvocationType;
            }

            private final int mFormatInvocationType;

            public Object getValue() {
                return mValue;
            }

            private final Object mValue;

            public String getOriginalFormat() {
                return mOriginalFormat;
            }

            private final String mOriginalFormat;

            public String getNewValue() {
                return mNewValue;
            }

            private final String mNewValue;

            public FormatInvocation(int formatInvocationType, Object value, String originalFormat, String newValue) {
                mValue = value;
                mFormatInvocationType = formatInvocationType;
                mOriginalFormat = originalFormat;
                mNewValue = newValue;
            }
        }

        public final class FormatInvocationType {
            private FormatInvocationType() {
            }

            public static final int NUMERIC = 0;
            public static final int DATE_TIME = 1;
            public static final int GENERAL = 2;
            public static final int ALL = 3;

            public static final int length = 4;
        }
    }
    //ExEnd

    @Test
    public void insertVideoWithUrl() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertOnlineVideo(String, Double, Double)
        //ExSummary:Shows how to insert an online video into a document using a URL.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertOnlineVideo("https://youtu.be/t_1LYZ102RA", 360.0, 270.0);

        // We can watch the video from Microsoft Word by clicking on the shape.
        doc.save(getArtifactsDir() + "DocumentBuilder.InsertVideoWithUrl.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertVideoWithUrl.docx");
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(480, 360, ImageType.JPEG, shape);
        TestUtil.verifyWebResponseStatusCode(200, new URL(shape.getHRef()));

        Assert.assertEquals(360.0d, shape.getWidth());
        Assert.assertEquals(270.0d, shape.getHeight());
    }

    @Test
    public void insertUnderline() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.Underline
        //ExSummary:Shows how to format text inserted by a document builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.setUnderline(Underline.DASH);
        builder.getFont().setColor(Color.BLUE);
        builder.getFont().setSize(32.0);

        // The builder applies formatting to its current paragraph and any new text added by it afterward.
        builder.writeln("Large, blue, and underlined text.");

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertUnderline.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertUnderline.docx");
        Run firstRun = doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0);

        Assert.assertEquals("Large, blue, and underlined text.", firstRun.getText().trim());
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

        // A Story is a type of node that has child Paragraph nodes, such as a Body.
        Assert.assertEquals(builder.getCurrentStory(), doc.getFirstSection().getBody());
        Assert.assertEquals(builder.getCurrentStory(), builder.getCurrentParagraph().getParentNode());
        Assert.assertEquals(StoryType.MAIN_TEXT, builder.getCurrentStory().getStoryType());

        builder.getCurrentStory().appendParagraph("Text added to current Story.");

        // A Story can also contain tables.
        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Row 1, cell 1");
        builder.insertCell();
        builder.write("Row 1, cell 2");
        builder.endTable();

        Assert.assertTrue(builder.getCurrentStory().getTables().contains(table));
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        Assert.assertEquals(1, doc.getFirstSection().getBody().getTables().getCount());
        Assert.assertEquals("Row 1, cell 1\u0007Row 1, cell 2\u0007\u0007\rText added to current Story.", doc.getFirstSection().getBody().getText().trim());
    }

    @Test
    public void insertOleObjects() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertOleObject(Stream, String, Boolean, Stream)
        //ExSummary:Shows how to use document builder to embed OLE objects in a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Microsoft Excel spreadsheet from the local file system
        // into the document while keeping its default appearance.
        InputStream spreadsheetStream = new FileInputStream(getMyDir() + "Spreadsheet.xlsx");
        InputStream representingImage = new FileInputStream(getImageDir() + "Logo.jpg");
        try {
            builder.writeln("Spreadsheet Ole object:");
            builder.insertOleObject(spreadsheetStream, "OleObject.xlsx", false, representingImage);
        } finally {
            if (spreadsheetStream != null) spreadsheetStream.close();
        }

        // Double-click these objects in Microsoft Word to open
        // the linked files using their respective applications.
        doc.save(getArtifactsDir() + "DocumentBuilder.InsertOleObjects.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertOleObjects.docx");

        Assert.assertEquals(1, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        Assert.assertEquals("", shape.getOleFormat().getIconCaption());
        Assert.assertFalse(shape.getOleFormat().getOleIcon());
    }

    @Test
    public void insertStyleSeparator() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertStyleSeparator
        //ExSummary:Shows how to work with style separators.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Each paragraph can only have one style.
        // The InsertStyleSeparator method allows us to work around this limitation.
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.write("This text is in a Heading style. ");
        builder.insertStyleSeparator();

        Style paraStyle = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "MyParaStyle");
        paraStyle.getFont().setBold(false);
        paraStyle.getFont().setSize(8.0);
        paraStyle.getFont().setName("Arial");

        builder.getParagraphFormat().setStyleName(paraStyle.getName());
        builder.write("This text is in a custom style. ");

        // Calling the InsertStyleSeparator method creates another paragraph,
        // which can have a different style to the previous. There will be no break between paragraphs.
        // The text in the output document will look like one paragraph with two styles.
        Assert.assertEquals(2, doc.getFirstSection().getBody().getParagraphs().getCount());
        Assert.assertEquals("Heading 1", doc.getFirstSection().getBody().getParagraphs().get(0).getParagraphFormat().getStyle().getName());
        Assert.assertEquals("MyParaStyle", doc.getFirstSection().getBody().getParagraphs().get(1).getParagraphFormat().getStyle().getName());

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertStyleSeparator.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertStyleSeparator.docx");

        Assert.assertEquals(2, doc.getFirstSection().getBody().getParagraphs().getCount());
        Assert.assertEquals("This text is in a Heading style. \r This text is in a custom style.",
                doc.getText().trim());
        Assert.assertEquals("Heading 1", doc.getFirstSection().getBody().getParagraphs().get(0).getParagraphFormat().getStyle().getName());
        Assert.assertEquals("MyParaStyle", doc.getFirstSection().getBody().getParagraphs().get(1).getParagraphFormat().getStyle().getName());
        Assert.assertEquals(" ", doc.getFirstSection().getBody().getParagraphs().get(1).getRuns().get(0).getText());
        TestUtil.docPackageFileContainsString("w:rPr><w:vanish /><w:specVanish /></w:rPr>",
                getArtifactsDir() + "DocumentBuilder.InsertStyleSeparator.docx", "document.xml");
        TestUtil.docPackageFileContainsString("<w:t xml:space=\"preserve\"> </w:t>",
                getArtifactsDir() + "DocumentBuilder.InsertStyleSeparator.docx", "document.xml");
    }

    @Test(enabled = false, description = "Bug: does not insert headers and footers, all lists (bullets, numbering, multilevel) breaks")
    public void insertDocument() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertDocument(Document, ImportFormatMode)
        //ExFor:ImportFormatMode
        //ExSummary:Shows how to insert a document into another document.
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
    public void smartStyleBehavior() throws Exception {
        //ExStart
        //ExFor:ImportFormatOptions
        //ExFor:ImportFormatOptions.SmartStyleBehavior
        //ExFor:DocumentBuilder.InsertDocument(Document, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how to resolve duplicate styles while inserting documents.
        Document dstDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        Style myStyle = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "MyStyle");
        myStyle.getFont().setSize(14.0);
        myStyle.getFont().setName("Courier New");
        myStyle.getFont().setColor(Color.BLUE);

        builder.getParagraphFormat().setStyleName(myStyle.getName());
        builder.writeln("Hello world!");

        // Clone the document and edit the clone's "MyStyle" style, so it is a different color than that of the original.
        // If we insert the clone into the original document, the two styles with the same name will cause a clash.
        Document srcDoc = dstDoc.deepClone();
        srcDoc.getStyles().get("MyStyle").getFont().setColor(Color.RED);

        // When we enable SmartStyleBehavior and use the KeepSourceFormatting import format mode,
        // Aspose.Words will resolve style clashes by converting source document styles.
        // with the same names as destination styles into direct paragraph attributes.
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

    @Test
    public void emphasesWarningSourceMarkdown() throws Exception {
        Document doc = new Document(getMyDir() + "Emphases markdown warning.docx");

        WarningInfoCollection warnings = new WarningInfoCollection();
        doc.setWarningCallback(warnings);
        doc.save(getArtifactsDir() + "DocumentBuilder.EmphasesWarningSourceMarkdown.md");

        for (WarningInfo warningInfo : warnings) {
            if (warningInfo.getSource() == WarningSource.MARKDOWN)
                Assert.assertEquals("The (*, 0:11) cannot be properly written into Markdown.", warningInfo.getDescription());
        }
    }

    @Test
    public void doNotIgnoreHeaderFooter() throws Exception {
        //ExStart
        //ExFor:ImportFormatOptions.IgnoreHeaderFooter
        //ExSummary:Shows how to specifies ignoring or not source formatting of headers/footers content.
        Document dstDoc = new Document(getMyDir() + "Document.docx");
        Document srcDoc = new Document(getMyDir() + "Header and footer types.docx");

        ImportFormatOptions importFormatOptions = new ImportFormatOptions();
        importFormatOptions.setIgnoreHeaderFooter(false);

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING, importFormatOptions);

        dstDoc.save(getArtifactsDir() + "DocumentBuilder.DoNotIgnoreHeaderFooter.docx");
        //ExEnd
    }

    /// <summary>
    /// All markdown tests work with the same file. That's why we need order for them.
    /// </summary>
    @Test(priority = 1)
    public void markdownDocumentEmphases() throws Exception {
        DocumentBuilder builder = new DocumentBuilder();

        // Bold and Italic are represented as Font.Bold and Font.Italic.
        builder.getFont().setItalic(true);
        builder.writeln("This text will be italic");

        // Use clear formatting if we don't want to combine styles between paragraphs.
        builder.getFont().clearFormatting();

        builder.getFont().setBold(true);
        builder.writeln("This text will be bold");

        builder.getFont().clearFormatting();

        builder.getFont().setItalic(true);
        builder.write("You ");
        builder.getFont().setBold(true);
        builder.write("can");
        builder.getFont().setBold(false);
        builder.writeln(" combine them");

        builder.getFont().clearFormatting();

        builder.getFont().setStrikeThrough(true);
        builder.writeln("This text will be strikethrough");

        // Markdown treats asterisks (*), underscores (_) and tilde (~) as indicators of emphasis.
        builder.getDocument().save(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file. That's why we need order for them.
    /// </summary>
    @Test(priority = 2)
    public void markdownDocumentInlineCode() throws Exception {
        Document doc = new Document(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare our created document for further work
        // and clear paragraph formatting not to use the previous styles.
        builder.moveToDocumentEnd();
        builder.getParagraphFormat().clearFormatting();
        builder.writeln("\n");

        // Style with name that starts from word InlineCode, followed by optional dot (.) and number of backticks (`).
        // If number of backticks is missed, then one backtick will be used by default.
        Style inlineCode1BackTicks = doc.getStyles().add(StyleType.CHARACTER, "InlineCode");
        builder.getFont().setStyle(inlineCode1BackTicks);
        builder.writeln("Text with InlineCode style with one backtick");

        // Use optional dot (.) and number of backticks (`).
        // There will be 3 backticks.
        Style inlineCode3BackTicks = doc.getStyles().add(StyleType.CHARACTER, "InlineCode.3");
        builder.getFont().setStyle(inlineCode3BackTicks);
        builder.writeln("Text with InlineCode style with 3 backticks");

        builder.getDocument().save(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file. That's why we need order for them.
    /// </summary>
    @Test(description = "WORDSNET-19850", priority = 3)
    public void markdownDocumentHeadings() throws Exception {
        Document doc = new Document(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare our created document for further work
        // and clear paragraph formatting not to use the previous styles.
        builder.moveToDocumentEnd();
        builder.getParagraphFormat().clearFormatting();
        builder.writeln("\n");

        // By default, Heading styles in Word may have bold and italic formatting.
        // If we do not want text to be emphasized, set these properties explicitly to false.
        // Thus we can't use 'builder.Font.ClearFormatting()' because Bold/Italic will be set to true.
        builder.getFont().setBold(false);
        builder.getFont().setItalic(false);

        // Create for one heading for each level.
        builder.getParagraphFormat().setStyleName("Heading 1");
        builder.getFont().setItalic(true);
        builder.writeln("This is an italic H1 tag");

        // Reset our styles from the previous paragraph to not combine styles between paragraphs.
        builder.getFont().setBold(false);
        builder.getFont().setItalic(false);

        // Structure-enhanced text heading can be added through style inheritance.
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
    /// All markdown tests work with the same file. That's why we need order for them.
    /// </summary>
    @Test(priority = 4)
    public void markdownDocumentBlockquotes() throws Exception {
        Document doc = new Document(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare our created document for further work
        // and clear paragraph formatting not to use the previous styles.
        builder.moveToDocumentEnd();
        builder.getParagraphFormat().clearFormatting();
        builder.writeln("\n");

        // By default, the document stores blockquote style for the first level.
        builder.getParagraphFormat().setStyleName("Quote");
        builder.writeln("Blockquote");

        // Create styles for nested levels through style inheritance.
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
    /// All markdown tests work with the same file. That's why we need order for them.
    /// </summary>
    @Test(priority = 5)
    public void markdownDocumentIndentedCode() throws Exception {
        Document doc = new Document(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare our created document for further work
        // and clear paragraph formatting not to use the previous styles.
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
    /// All markdown tests work with the same file. That's why we need order for them.
    /// </summary>
    @Test(priority = 6)
    public void markdownDocumentFencedCode() throws Exception {
        Document doc = new Document(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare our created document for further work
        // and clear paragraph formatting not to use the previous styles.
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
    /// All markdown tests work with the same file. That's why we need order for them.
    /// </summary>
    @Test(priority = 7)
    public void markdownDocumentHorizontalRule() throws Exception {
        Document doc = new Document(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare our created document for further work
        // and clear paragraph formatting not to use the previous styles.
        builder.moveToDocumentEnd();
        builder.getParagraphFormat().clearFormatting();
        builder.writeln("\n");

        // Insert HorizontalRule that will be present in .md file as '-----'.
        builder.insertHorizontalRule();

        builder.getDocument().save(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file. That's why we need order for them.
    /// </summary>
    @Test(priority = 8)
    public void markdownDocumentBulletedList() throws Exception {
        Document doc = new Document(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare our created document for further work
        // and clear paragraph formatting not to use the previous styles.
        builder.moveToDocumentEnd();
        builder.getParagraphFormat().clearFormatting();
        builder.writeln("\n");

        // Bulleted lists are represented using paragraph numbering.
        builder.getListFormat().applyBulletDefault();
        // There can be 3 types of bulleted lists.
        // The only diff in a numbering format of the very first level are ‘-’, ‘+’ or ‘*’ respectively.
        builder.getListFormat().getList().getListLevels().get(0).setNumberFormat("-");

        builder.writeln("Item 1");
        builder.writeln("Item 2");
        builder.getListFormat().listIndent();
        builder.writeln("Item 2a");
        builder.writeln("Item 2b");

        builder.getDocument().save(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file. That's why we need order for them.
    /// </summary>
    @Test(dataProvider = "loadMarkdownDocumentAndAssertContentDataProvider", priority = 9)
    public void loadMarkdownDocumentAndAssertContent(String text, String styleName, boolean isItalic, boolean isBold) throws Exception {
        // Load created document from previous tests.
        Document doc = new Document(getArtifactsDir() + "DocumentBuilder.MarkdownDocument.md");
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        for (Paragraph paragraph : paragraphs) {
            if (paragraph.getRuns().getCount() != 0) {
                // Check that all document text has the necessary styles.
                if (paragraph.getRuns().get(0).getText().equals(text) && !text.contains("InlineCode")) {
                    Assert.assertEquals(styleName, paragraph.getParagraphFormat().getStyle().getName());
                    Assert.assertEquals(isItalic, paragraph.getRuns().get(0).getFont().getItalic());
                    Assert.assertEquals(isBold, paragraph.getRuns().get(0).getFont().getBold());
                } else if (paragraph.getRuns().get(0).getText().equals(text) && text.contains("InlineCode")) {
                    Assert.assertEquals(styleName, paragraph.getRuns().get(0).getFont().getStyleName());
                }
            }

            // Check that document also has a HorizontalRule present as a shape.
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

    @Test(dataProvider = "markdownDocumentTableContentAlignmentDataProvider")
    public void markdownDocumentTableContentAlignment(int tableContentAlignment) throws Exception {
        DocumentBuilder builder = new DocumentBuilder();

        builder.insertCell();
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);
        builder.write("Cell1");
        builder.insertCell();
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.write("Cell2");

        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        saveOptions.setTableContentAlignment(tableContentAlignment);

        builder.getDocument().save(getArtifactsDir() + "DocumentBuilder.MarkdownDocumentTableContentAlignment.md", saveOptions);

        Document doc = new Document(getArtifactsDir() + "DocumentBuilder.MarkdownDocumentTableContentAlignment.md");
        Table table = doc.getFirstSection().getBody().getTables().get(0);

        switch (tableContentAlignment) {
            case TableContentAlignment.AUTO:
                Assert.assertEquals(ParagraphAlignment.RIGHT,
                        table.getFirstRow().getCells().get(0).getFirstParagraph().getParagraphFormat().getAlignment());
                Assert.assertEquals(ParagraphAlignment.CENTER,
                        table.getFirstRow().getCells().get(1).getFirstParagraph().getParagraphFormat().getAlignment());
                break;
            case TableContentAlignment.LEFT:
                Assert.assertEquals(ParagraphAlignment.LEFT,
                        table.getFirstRow().getCells().get(0).getFirstParagraph().getParagraphFormat().getAlignment());
                Assert.assertEquals(ParagraphAlignment.LEFT,
                        table.getFirstRow().getCells().get(1).getFirstParagraph().getParagraphFormat().getAlignment());
                break;
            case TableContentAlignment.CENTER:
                Assert.assertEquals(ParagraphAlignment.CENTER,
                        table.getFirstRow().getCells().get(0).getFirstParagraph().getParagraphFormat().getAlignment());
                Assert.assertEquals(ParagraphAlignment.CENTER,
                        table.getFirstRow().getCells().get(1).getFirstParagraph().getParagraphFormat().getAlignment());
                break;
            case TableContentAlignment.RIGHT:
                Assert.assertEquals(ParagraphAlignment.RIGHT,
                        table.getFirstRow().getCells().get(0).getFirstParagraph().getParagraphFormat().getAlignment());
                Assert.assertEquals(ParagraphAlignment.RIGHT,
                        table.getFirstRow().getCells().get(1).getFirstParagraph().getParagraphFormat().getAlignment());
                break;
        }
    }

    @DataProvider(name = "markdownDocumentTableContentAlignmentDataProvider")
    public static Object[][] markdownDocumentTableContentAlignmentDataProvider() {
        return new Object[][]
                {
                        {TableContentAlignment.LEFT},
                        {TableContentAlignment.RIGHT},
                        {TableContentAlignment.CENTER},
                        {TableContentAlignment.AUTO},
                };
    }

    //ExStart
    //ExFor:MarkdownSaveOptions.ImageSavingCallback
    //ExFor:IImageSavingCallback
    //ExSummary:Shows how to rename the image name during saving into Markdown document.
    @Test //ExSkip
    public void renameImages() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        MarkdownSaveOptions options = new MarkdownSaveOptions();

        // If we convert a document that contains images into Markdown, we will end up with one Markdown file which links to several images.
        // Each image will be in the form of a file in the local file system.
        // There is also a callback that can customize the name and file system location of each image.
        options.setImageSavingCallback(new SavedImageRename("DocumentBuilder.HandleDocument.md"));

        // The ImageSaving() method of our callback will be run at this time.
        doc.save(getArtifactsDir() + "DocumentBuilder.HandleDocument.md", options);

        Assert.assertEquals(1, DocumentHelper.directoryGetFiles(getArtifactsDir(),"*.*").stream().filter(f -> f.endsWith(".jpeg")).count());
        Assert.assertEquals(8, DocumentHelper.directoryGetFiles(getArtifactsDir(),"*.*").stream().filter(f -> f.endsWith(".png")).count());
    }

    /// <summary>
    /// Renames saved images that are produced when an Markdown document is saved.
    /// </summary>
    public static class SavedImageRename implements IImageSavingCallback
    {
        public SavedImageRename(String outFileName)
        {
            mOutFileName = outFileName;
        }

        public void imageSaving(ImageSavingArgs args) throws Exception {
            String imageFileName = MessageFormat.format("{0} shape {1}, of type {2}.{3}", mOutFileName, ++mCount, args.getCurrentShape().getShapeType(), FilenameUtils.getExtension(args.getImageFileName()));

            args.setImageFileName(imageFileName);
            args.setImageStream(new FileOutputStream(getArtifactsDir() + imageFileName));

            Assert.assertTrue(args.isImageAvailable());
            Assert.assertFalse(args.getKeepImageStreamOpen());
        }

        private int mCount;
        private String mOutFileName;
    }
    //ExEnd

    @Test
    public void insertOnlineVideo() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertOnlineVideo(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert an online video into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        String videoUrl = "https://vimeo.com/52477838";

        // Insert a shape that plays a video from the web when clicked in Microsoft Word.
        // This rectangular shape will contain an image based on the first frame of the linked video
        // and a "play button" visual prompt. The video has an aspect ratio of 16:9.
        // We will set the shape's size to that ratio, so the image does not appear stretched.
        builder.insertOnlineVideo(videoUrl, RelativeHorizontalPosition.LEFT_MARGIN, 0.0,
                RelativeVerticalPosition.TOP_MARGIN, 0.0, 320.0, 180.0, WrapType.SQUARE);

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
        TestUtil.verifyWebResponseStatusCode(200, new URL(shape.getHRef()));
    }

    @Test
    public void insertOnlineVideoCustomThumbnail() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertOnlineVideo(String, String, Byte[], Double, Double)
        //ExFor:DocumentBuilder.InsertOnlineVideo(String, String, Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert an online video into a document with a custom thumbnail.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        String videoUrl = "https://vimeo.com/52477838";
        String videoEmbedCode = "<iframe src=\"https://player.vimeo.com/video/52477838\" width=\"640\" height=\"360\" frameborder=\"0\" " +
                "title=\"Aspose\" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe>";

        byte[] thumbnailImageBytes = IOUtils.toByteArray(getAsposelogoUri().toURL().openStream());

        BufferedImage image = ImageIO.read(new ByteArrayInputStream(thumbnailImageBytes));

        // Below are two ways of creating a shape with a custom thumbnail, which links to an online video
        // that will play when we click on the shape in Microsoft Word.
        // 1 -  Insert an inline shape at the builder's node insertion cursor:
        builder.insertOnlineVideo(videoUrl, videoEmbedCode, thumbnailImageBytes, image.getWidth(), image.getHeight());

        builder.insertBreak(BreakType.PAGE_BREAK);

        // 2 -  Insert a floating shape:
        double left = builder.getPageSetup().getRightMargin() - image.getWidth();
        double top = builder.getPageSetup().getBottomMargin() - image.getHeight();

        builder.insertOnlineVideo(videoUrl, videoEmbedCode, thumbnailImageBytes,
                RelativeHorizontalPosition.RIGHT_MARGIN, left, RelativeVerticalPosition.BOTTOM_MARGIN, top,
                image.getWidth(), image.getHeight(), WrapType.SQUARE);

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertOnlineVideoCustomThumbnail.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.InsertOnlineVideoCustomThumbnail.docx");
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(320, 320, ImageType.PNG, shape);
        Assert.assertEquals(320.0d, shape.getWidth());
        Assert.assertEquals(320.0d, shape.getHeight());
        Assert.assertEquals(0.0d, shape.getLeft());
        Assert.assertEquals(0.0d, shape.getTop());
        Assert.assertEquals(WrapType.INLINE, shape.getWrapType());
        Assert.assertEquals(RelativeVerticalPosition.PARAGRAPH, shape.getRelativeVerticalPosition());
        Assert.assertEquals(RelativeHorizontalPosition.COLUMN, shape.getRelativeHorizontalPosition());

        Assert.assertEquals("https://vimeo.com/52477838", shape.getHRef());

        shape = (Shape) doc.getChild(NodeType.SHAPE, 1, true);

        TestUtil.verifyImageInShape(320, 320, ImageType.PNG, shape);
        Assert.assertEquals(320.0d, shape.getWidth());
        Assert.assertEquals(320.0d, shape.getHeight());
        Assert.assertEquals(-248.0d, shape.getLeft());
        Assert.assertEquals(-248.0d, shape.getTop());
        Assert.assertEquals(WrapType.SQUARE, shape.getWrapType());
        Assert.assertEquals(RelativeVerticalPosition.BOTTOM_MARGIN, shape.getRelativeVerticalPosition());
        Assert.assertEquals(RelativeHorizontalPosition.RIGHT_MARGIN, shape.getRelativeHorizontalPosition());

        Assert.assertEquals("https://vimeo.com/52477838", shape.getHRef());
        TestUtil.verifyWebResponseStatusCode(200, new URL(shape.getHRef()));
    }

    @Test
    public void insertOleObjectAsIcon() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertOleObjectAsIcon(String, String, Boolean, String, String)
        //ExFor:DocumentBuilder.InsertOleObjectAsIcon(Stream, String, String, String)
        //ExSummary:Shows how to insert an embedded or linked OLE object as icon into the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // If 'iconFile' and 'iconCaption' are omitted, this overloaded method selects
        // the icon according to 'progId' and uses the filename for the icon caption.
        builder.insertOleObjectAsIcon(getMyDir() + "Presentation.pptx", "Package", false, getImageDir() + "Logo icon.ico", "My embedded file");

        builder.insertBreak(BreakType.LINE_BREAK);

        try (FileInputStream stream = new FileInputStream(getMyDir() + "Presentation.pptx")) {
            // If 'iconFile' and 'iconCaption' are omitted, this overloaded method selects
            // the icon according to the file extension and uses the filename for the icon caption.
            Shape shape = builder.insertOleObjectAsIcon(stream, "PowerPoint.Application", getImageDir() + "Logo icon.ico",
                    "My embedded file stream");

            OlePackage setOlePackage = shape.getOleFormat().getOlePackage();
            setOlePackage.setFileName("Presentation.pptx");
            setOlePackage.setDisplayName("Presentation.pptx");
        }

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertOleObjectAsIcon.docx");
        //ExEnd
    }
}
