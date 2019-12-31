// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Font;
import java.awt.Color;
import com.aspose.words.Underline;
import com.aspose.words.Document;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.BreakType;
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;
import com.aspose.words.Field;
import com.aspose.ms.System.msConsole;
import com.aspose.words.Shape;
import com.aspose.words.NodeType;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.words.SaveFormat;
import com.aspose.words.FieldType;
import com.aspose.words.FieldIf;
import com.aspose.words.ControlChar;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.WrapType;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.TextFormFieldType;
import com.aspose.words.FormFieldCollection;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.CellVerticalAlignment;
import com.aspose.ms.System.Drawing.msColor;
import com.aspose.words.HeightRule;
import com.aspose.words.LineStyle;
import com.aspose.words.TextOrientation;
import com.aspose.words.Table;
import com.aspose.words.TableStyleOptions;
import com.aspose.words.AutoFitBehavior;
import com.aspose.words.PreferredWidth;
import com.aspose.words.PreferredWidthType;
import com.aspose.words.Cell;
import com.aspose.words.ConvertUtil;
import com.aspose.words.Node;
import com.aspose.words.Paragraph;
import com.aspose.words.ParagraphFormat;
import com.aspose.words.SignatureLineOptions;
import com.aspose.words.SignatureLine;
import com.aspose.ms.System.Guid;
import com.aspose.words.SignOptions;
import com.aspose.ms.System.DateTime;
import com.aspose.words.CertificateHolder;
import com.aspose.words.DigitalSignatureUtil;
import com.aspose.words.CellFormat;
import com.aspose.words.RowFormat;
import com.aspose.words.Orientation;
import com.aspose.words.PaperSize;
import com.aspose.words.FootnoteType;
import com.aspose.ms.System.msString;
import com.aspose.words.NumberStyle;
import com.aspose.words.FootnoteNumberingRule;
import com.aspose.words.BorderCollection;
import com.aspose.words.BorderType;
import com.aspose.words.Shading;
import com.aspose.words.TextureIndex;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.ImportFormatOptions;
import com.aspose.words.NodeImporter;
import com.aspose.words.ParagraphCollection;
import com.aspose.words.ChartType;
import com.aspose.words.IFieldResultFormatter;
import com.aspose.ms.System.Collections.msArrayList;
import com.aspose.words.CalendarType;
import com.aspose.words.GeneralFormat;
import java.util.ArrayList;
import com.aspose.words.StoryType;
import com.aspose.ms.System.IO.Stream;
import com.aspose.ms.System.IO.File;
import com.aspose.ms.System.IO.FileMode;
import java.awt.image.BufferedImage;
import com.aspose.words.Style;
import com.aspose.words.StyleType;
import org.testng.annotations.DataProvider;


@Test
public class ExDocumentBuilder extends ApiExampleBase
{
    @Test
    public void writeAndFont() throws Exception
    {
        //ExStart
        //ExFor:Font.Size
        //ExFor:Font.Bold
        //ExFor:Font.Name
        //ExFor:Font.Color
        //ExFor:Font.Underline
        //ExFor:DocumentBuilder.#ctor
        //ExSummary:Inserts formatted text using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder();

        // Specify font formatting before adding text
        Font font = builder.getFont();
        font.setSize(16.0);
        font.setBold(true);
        font.setColor(Color.BLUE);
        font.setName("Arial");
        font.setUnderline(Underline.DASH);

        builder.write("Sample text.");
        //ExEnd
    }

    @Test
    public void headersAndFooters() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.#ctor(Document)
        //ExFor:DocumentBuilder.MoveToHeaderFooter
        //ExFor:DocumentBuilder.MoveToSection
        //ExFor:DocumentBuilder.InsertBreak
        //ExFor:DocumentBuilder.Writeln
        //ExFor:HeaderFooterType
        //ExFor:PageSetup.DifferentFirstPageHeaderFooter
        //ExFor:PageSetup.OddAndEvenPagesHeaderFooter
        //ExFor:BreakType
        //ExSummary:Creates headers and footers in a document using DocumentBuilder.
        // Create a blank document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Specify that we want headers and footers different for first, even and odd pages
        builder.getPageSetup().setDifferentFirstPageHeaderFooter(true);
        builder.getPageSetup().setOddAndEvenPagesHeaderFooter(true);

        // Create the headers
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
        builder.write("Header First");
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_EVEN);
        builder.write("Header Even");
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.write("Header Odd");

        // Create three pages in the document
        builder.moveToSection(0);
        builder.writeln("Page1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page2");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page3");

        doc.save(getArtifactsDir() + "DocumentBuilder.HeadersAndFooters.doc");
        //ExEnd
    }

    @Test
    public void insertMergeField() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertField(String)
        //ExFor:DocumentBuilder.MoveToMergeField(String, Boolean, Boolean)
        //ExSummary:Shows how to insert merge fields and move between them.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField("MERGEFIELD MyMergeField1 \\* MERGEFORMAT");
        builder.insertField("MERGEFIELD MyMergeField2 \\* MERGEFORMAT");

        msAssert.areEqual(2, doc.getRange().getFields().getCount());

        // The second merge field starts immediately after the end of the first
        // We'll move the builder's cursor to the end of the first so we can split them by text
        builder.moveToMergeField("MyMergeField1", true, false);

        builder.write(" Text between our two merge fields. ");

        doc.save(getArtifactsDir() + "DocumentBuilder.MergeFields.docx");
        //ExEnd			
    }

    @Test
    public void insertFieldFieldCode() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertField(String)
        //ExFor:Field
        //ExFor:Field.Update
        //ExFor:Field.Result
        //ExFor:Field.GetFieldCode
        //ExFor:Field.Type
        //ExFor:Field.Remove
        //ExFor:FieldType
        //ExSummary:Inserts a field into a document using DocumentBuilder and FieldCode.
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
        msConsole.writeLine("FieldResult: {0}", dateField.getResult());

        // Display the field code which defines the behavior of the field. This can been seen in Microsoft Word by pressing ALT+F9
        msConsole.writeLine("FieldCode: {0}", dateField.getFieldCode());

        // The field type defines what type of field in the Document this is. In this case the type is "FieldDate" 
        msConsole.writeLine("FieldType: {0}", dateField.getType());

        // Finally let's completely remove the field from the document. This can easily be done by invoking the Remove method on the object
        dateField.remove();
        //ExEnd			
    }

    @Test
    public void insertHorizontalRule() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertHorizontalRule
        //ExFor:ShapeBase.IsHorizontalRule
        //ExSummary:Shows how to insert horizontal rule shape in a document.
        // Use a document builder to insert a horizontal rule
        DocumentBuilder builder = new DocumentBuilder();
        builder.insertHorizontalRule();

        // Get the rule from the document's shape collection and verify it
        Shape horizontalRule = (Shape)builder.getDocument().getChild(NodeType.SHAPE, 0, true);
        Assert.assertTrue(horizontalRule.isHorizontalRule());
        //ExEnd
    }

    @Test
    public void fieldLocale() throws Exception
    {
        //ExStart
        //ExFor:Field.LocaleId
        //ExSummary: Get or sets locale for fields
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Field field = builder.insertField("DATE \\* MERGEFORMAT");
        field.setLocaleId(2064);

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        Field newField = doc.getRange().getFields().get(0);
        msAssert.areEqual(2064, newField.getLocaleId());
        //ExEnd
    }

    @Test (dataProvider = "getFieldCodeDataProvider")
    public void getFieldCode(boolean containsNestedFields) throws Exception
    {
        //ExStart
        //ExFor:Field.GetFieldCode
        //ExFor:Field.GetFieldCode(bool)
        //ExSummary:Shows how to get text between field start and field separator (or field end if there is no separator).
        Document doc = new Document(getMyDir() + "Field.FieldCode.docx");

        for (Field field : doc.getRange().getFields())
        {
            if (field.getType() == FieldType.FIELD_IF)
            {
                FieldIf fieldIf = (FieldIf)field;

                String fieldCode = fieldIf.getFieldCode();
                msAssert.areEqual(" IF " + ControlChar.FIELD_START_CHAR + " MERGEFIELD Q223 " + ControlChar.FIELD_SEPARATOR_CHAR + ControlChar.FIELD_END_CHAR + " > 0 \" (and additionally London Weighting of  " + ControlChar.FIELD_START_CHAR + " MERGEFIELD  Q223 \\f £ " + ControlChar.FIELD_SEPARATOR_CHAR + ControlChar.FIELD_END_CHAR + " per hour) \" \"\" ",fieldCode); //ExSkip

                if (containsNestedFields)
                {
                    fieldCode = fieldIf.getFieldCode(true);
                    msAssert.areEqual(" IF " + ControlChar.FIELD_START_CHAR + " MERGEFIELD Q223 " + ControlChar.FIELD_SEPARATOR_CHAR + ControlChar.FIELD_END_CHAR + " > 0 \" (and additionally London Weighting of  " + ControlChar.FIELD_START_CHAR + " MERGEFIELD  Q223 \\f £ " + ControlChar.FIELD_SEPARATOR_CHAR + ControlChar.FIELD_END_CHAR + " per hour) \" \"\" ",fieldCode); //ExSkip
                }
                else
                {
                    fieldCode = fieldIf.getFieldCode(false);
                    msAssert.areEqual(" IF  > 0 \" (and additionally London Weighting of   per hour) \" \"\" ",fieldCode); //ExSkip
                }
            }
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "getFieldCodeDataProvider")
	public static Object[][] getFieldCodeDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test
    public void documentBuilderAndSave() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello World!");

        doc.save(getArtifactsDir() + "DocumentBuilderAndSave.docx");
    }

    @Test
    public void insertHyperlink() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertHyperlink
        //ExFor:Font.ClearFormatting
        //ExFor:Font.Color
        //ExFor:Font.Underline
        //ExFor:Underline
        //ExSummary:Inserts a hyperlink into a document using DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Please make sure to visit ");

        // Specify font formatting for the hyperlink
        builder.getFont().setColor(Color.BLUE);
        builder.getFont().setUnderline(Underline.SINGLE);
        // Insert the link.
        builder.insertHyperlink("Aspose Website", "http://www.aspose.com", false);

        // Revert to default formatting
        builder.getFont().clearFormatting();

        builder.write(" for more information.");

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertHyperlink.doc");
        //ExEnd
    }

    @Test
    public void pushPopFont() throws Exception
    {
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
        builder.write("To go to an important location, click ");

        // Save the font formatting so we use different formatting for hyperlink and restore old formatting later
        builder.pushFont();

        // Set new font formatting for the hyperlink and insert the hyperlink
        // The "Hyperlink" style is a Microsoft Word built-in style so we don't have to worry to 
        // create it, it will be created automatically if it does not yet exist in the document
        builder.getFont().setStyleIdentifier(StyleIdentifier.HYPERLINK);
        builder.insertHyperlink("here", "http://www.google.com", false);

        // Restore the formatting that was before the hyperlink.
        builder.popFont();

        builder.writeln(". We hope you enjoyed the example.");

        doc.save(getArtifactsDir() + "DocumentBuilder.PushPopFont.doc");
        //ExEnd
    }

                @Test
    public void insertWatermarkNetStandard2() throws Exception
    {
        //ExStart
        //ExFor:HeaderFooterType
        //ExFor:DocumentBuilder.MoveToHeaderFooter
        //ExFor:PageSetup.PageWidth
        //ExFor:PageSetup.PageHeight
        //ExFor:DocumentBuilder.InsertImage(Image)
        //ExFor:WrapType
        //ExFor:RelativeHorizontalPosition
        //ExFor:RelativeVerticalPosition
        //ExSummary:Inserts a watermark image into a document using DocumentBuilder (.NetStandard 2.0).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The best place for the watermark image is in the header or footer so it is shown on every page
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);

        SKBitmap image = SKBitmap.Decode(getImageDir() + "Watermark.png");
        try /*JAVA: was using*/
        {
            // Insert a floating picture
            Shape shape = builder.InsertImage(image);
            shape.setWrapType(WrapType.NONE);
            shape.setBehindText(true);

            shape.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
            shape.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);

            // Calculate image left and top position so it appears in the center of the page
            shape.setLeft((builder.getPageSetup().getPageWidth() - shape.getWidth()) / 2.0);
            shape.setTop((builder.getPageSetup().getPageHeight() - shape.getHeight()) / 2.0);
        }
        finally { if (image != null) image.close(); }

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertWatermark.NetStandard2.doc");
        //ExEnd
    }

    @Test
    public void insertOleObjectNetStandard2() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertOleObject(String, Boolean, Boolean, Image)
        //ExFor:DocumentBuilder.InsertOleObject(String, String, Boolean, Boolean, Image)
        //ExSummary:Shows how to insert an OLE object into a document (.NetStandard 2.0).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        SKBitmap representingImage = SKBitmap.Decode(getImageDir() + "Aspose.Words.gif");
        try /*JAVA: was using*/
        {
            // OleObject
            builder.InsertOleObject(getMyDir() + "Document.Spreadsheet.xlsx", false, false, representingImage);
            // OleObject with ProgId
            builder.InsertOleObject(getMyDir() + "Document.Spreadsheet.xlsx", "Excel.Sheet", false, false,
                representingImage);
        }
        finally { if (representingImage != null) representingImage.close(); }

        doc.save(getArtifactsDir() + "Document.InsertedOleObject.NetStandard2.docx");
        //ExEnd
    }
    
    @Test
    public void insertHtml() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder
        //ExFor:DocumentBuilder.InsertHtml(String)
        //ExSummary:Inserts HTML into a document. The formatting specified in the HTML is applied.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        final String HTML = "<P align='right'>Paragraph right</P>" + "<b>Implicit paragraph left</b>" +
                            "<div align='center'>Div center</div>" + "<h1 align='left'>Heading 1 left.</h1>";

        builder.insertHtml(HTML);

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertHtml.doc");
        //ExEnd
    }

    @Test
    public void insertHtmlWithCurrentDocumentFormatting() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertHtml(String, Boolean)
        //ExSummary:Inserts HTML into a document using. The current document formatting at the insertion position is applied to the inserted text. 
        Document doc = new Document();

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertHtml(
            "<P align='right'>Paragraph right</P>" + "<b>Implicit paragraph left</b>" +
            "<div align='center'>Div center</div>" + "<h1 align='left'>Heading 1 left.</h1>", true);

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertHtml.doc");
        //ExEnd
    }

    @Test
    public void insertMathMl() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertHtml(String)
        //ExSummary:Inserts MathMl into a document using.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        final String MATH_ML =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><msub><mi>a</mi><mrow><mn>1</mn></mrow></msub><mo>+</mo><msub><mi>b</mi><mrow><mn>1</mn></mrow></msub></mrow></math>";

        builder.insertHtml(MATH_ML);
        //ExEnd

        doc.save(getArtifactsDir() + "MathML.docx");
        doc.save(getArtifactsDir() + "MathML.pdf");

        Assert.assertTrue(DocumentHelper.compareDocs(getGoldsDir() + "MathML Gold.docx", getArtifactsDir() + "MathML.docx"));
    }

    @Test
    public void insertTextAndBookmark() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder
        //ExFor:DocumentBuilder.StartBookmark
        //ExFor:DocumentBuilder.EndBookmark
        //ExSummary:Adds some text into the document and encloses the text in a bookmark using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder();

        builder.startBookmark("MyBookmark");
        builder.writeln("Text inside a bookmark.");
        builder.endBookmark("MyBookmark");
        //ExEnd
    }

    @Test
    public void createForm() throws Exception
    {
        //ExStart
        //ExFor:TextFormFieldType
        //ExFor:DocumentBuilder.InsertTextInput
        //ExFor:DocumentBuilder.InsertComboBox
        //ExSummary:Builds a sample form to fill.
        DocumentBuilder builder = new DocumentBuilder();

        // Insert a text form field for input a name
        builder.insertTextInput("", TextFormFieldType.REGULAR, "", "Enter your name here", 30);

        // Insert two blank lines
        builder.writeln("");
        builder.writeln("");

        String[] items =
        {
            "-- Select your favorite footwear --", "Sneakers", "Oxfords", "Flip-flops", "Other",
            "I prefer to be barefoot"
        };

        // Insert a combo box to select a footwear type
        builder.insertComboBox("", items, 0);

        // Insert 2 blank lines
        builder.writeln("");
        builder.writeln("");

        builder.getDocument().save(getArtifactsDir() + "DocumentBuilder.CreateForm.doc");
        //ExEnd
    }

    @Test
    public void insertCheckBox() throws Exception
    {
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

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        // Get checkboxes from the document
        FormFieldCollection formFields = doc.getRange().getFormFields();

        // Check that is the right checkbox
        msAssert.areEqual("", formFields.get(0).getName());

        //Assert that parameters sets correctly
        msAssert.areEqual(false, formFields.get(0).getChecked());
        msAssert.areEqual(false, formFields.get(0).getDefault());
        msAssert.areEqual(10, formFields.get(0).getCheckBoxSize());

        // Check that is the right checkbox
        // Please pay attention that MS Word allows strings with at most 20 characters
        msAssert.areEqual("CheckBox_Default", formFields.get(1).getName());

        //Assert that parameters sets correctly
        msAssert.areEqual(true, formFields.get(1).getChecked());
        msAssert.areEqual(true, formFields.get(1).getDefault());
        msAssert.areEqual(50, formFields.get(1).getCheckBoxSize());

        // Check that is the right checkbox
        // Please pay attention that MS Word allows strings with at most 20 characters
        msAssert.areEqual("CheckBox_OnlyChecked", formFields.get(2).getName());

        // Assert that parameters sets correctly
        msAssert.areEqual(true, formFields.get(2).getChecked());
        msAssert.areEqual(true, formFields.get(2).getDefault());
        msAssert.areEqual(100, formFields.get(2).getCheckBoxSize());
    }

    @Test
    public void insertCheckBoxEmptyName() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Checking that the checkbox insertion with an empty name working correctly
        builder.insertCheckBox("", true, false, 1);
        builder.insertCheckBox("", false, 1);
    }

    @Test
    public void workingWithNodes() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.MoveTo(Node)
        //ExFor:DocumentBuilder.MoveToBookmark(String)
        //ExFor:DocumentBuilder.CurrentParagraph
        //ExFor:DocumentBuilder.CurrentNode
        //ExFor:DocumentBuilder.MoveToDocumentStart
        //ExFor:DocumentBuilder.MoveToDocumentEnd
        //ExFor:DocumentBuilder.IsAtEndOfParagraph
        //ExFor:DocumentBuilder.IsAtStartOfParagraph
        //ExSummary:Shows how to move between nodes and manipulate current ones.
        Document doc = new Document(getMyDir() + "DocumentBuilder.WorkingWithNodes.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move to a bookmark and delete the parent paragraph
        builder.moveToBookmark("ParaToDelete");
        builder.getCurrentParagraph().remove();

        FindReplaceOptions options = new FindReplaceOptions();
        {
            options.setMatchCase(false);
            options.setFindWholeWordsOnly(true);
        }

        // Move to a particular paragraph's run and replace all occurrences of "bad" with "good" within this run
        builder.moveTo(doc.getLastSection().getBody().getParagraphs().get(0).getRuns().get(0));
        Assert.assertTrue(builder.isAtStartOfParagraph());
        Assert.assertFalse(builder.isAtEndOfParagraph());
        builder.getCurrentNode().getRange().replace("bad", "good", options);

        // Mark the beginning of the document
        builder.moveToDocumentStart();
        builder.writeln("Start of document.");

        // builder.WriteLn puts an end to its current paragraph after writing the text and starts a new one
        msAssert.areEqual(2, doc.getFirstSection().getBody().getParagraphs().getCount());
        Assert.assertTrue(builder.isAtStartOfParagraph());
        Assert.assertTrue(builder.isAtEndOfParagraph());

        // builder.Write doesn't end the paragraph
        builder.write("Second paragraph.");

        msAssert.areEqual(2, doc.getFirstSection().getBody().getParagraphs().getCount());
        Assert.assertFalse(builder.isAtStartOfParagraph());
        Assert.assertTrue(builder.isAtEndOfParagraph());

        // Mark the ending of the document
        builder.moveToDocumentEnd();
        builder.writeln("End of document.");

        doc.save(getArtifactsDir() + "DocumentBuilder.WorkingWithNodes.doc");
        //ExEnd
    }

    @Test
    public void fillingDocument() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.MoveToMergeField(String)
        //ExFor:DocumentBuilder.Bold
        //ExFor:DocumentBuilder.Italic
        //ExSummary:Fills document merge fields with some data.
        Document doc = new Document(getMyDir() + "DocumentBuilder.FillingDocument.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToMergeField("TeamLeaderName");
        builder.setBold(true);
        builder.writeln("Roman Korchagin");

        builder.moveToMergeField("SoftwareDeveloper1Name");
        builder.setItalic(true);
        builder.writeln("Dmitry Vorobyev");

        builder.moveToMergeField("SoftwareDeveloper2Name");
        builder.setItalic(true);
        builder.writeln("Vladimir Averkin");

        doc.save(getArtifactsDir() + "DocumentBuilder.FillingDocument.doc");
        //ExEnd
    }

    @Test
    public void insertToc() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertTableOfContents
        //ExFor:Document.UpdateFields
        //ExFor:DocumentBuilder.#ctor(Document)
        //ExFor:ParagraphFormat.StyleIdentifier
        //ExFor:DocumentBuilder.InsertBreak
        //ExFor:BreakType
        //ExSummary:Demonstrates how to insert a Table of contents (TOC) into a document using heading styles as entries.
        // Use a blank document
        Document doc = new Document();
        // Create a document builder to insert content with into document
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert a table of contents at the beginning of the document
        builder.insertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        // Start the actual document content on the second page
        builder.insertBreak(BreakType.PAGE_BREAK);
        // Build a document with complex structure by applying different heading styles thus creating TOC entries
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

        // Call the method below to update the TOC
        doc.updateFields();
        //ExEnd

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertToc.docx");
    }

    @Test
    public void insertTable() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder
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
        //ExFor:RowFormat.HeightRule
        //ExFor:RowFormat.Height
        //ExFor:HeightRule
        //ExFor:Shading.BackgroundPatternColor
        //ExFor:Shading.ClearFormatting
        //ExSummary:Shows how to build a nice bordered table.
        DocumentBuilder builder = new DocumentBuilder();

        // Start building a table
        builder.startTable();
        
        // Set the appropriate paragraph, cell, and row formatting. The formatting properties are preserved
        // until they are explicitly modified so there's no need to set them for each row or cell
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);

        builder.getCellFormat().clearFormatting();
        builder.getCellFormat().setWidth(150.0);
        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
        builder.getCellFormat().getShading().setBackgroundPatternColor(msColor.getGreenYellow());
        builder.getCellFormat().setWrapText(false);
        builder.getCellFormat().setFitText(true);

        builder.getRowFormat().clearFormatting();
        builder.getRowFormat().setHeightRule(HeightRule.EXACTLY);
        builder.getRowFormat().setHeight(50.0);
        builder.getRowFormat().getBorders().setLineStyle(LineStyle.ENGRAVE_3_D);
        builder.getRowFormat().getBorders().setColor(msColor.getOrange());

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

        builder.getDocument().save(getArtifactsDir() + "DocumentBuilder.InsertTable.docx");
        //ExEnd
    }

    @Test
    public void insertTableWithTableStyle() throws Exception
    {
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

        doc.save(getArtifactsDir() + "DocumentBuilder.SetTableStyle.docx");
        //ExEnd

        // Verify that the style was set by expanding to direct formatting
        doc.expandTableStylesToDirectFormatting();
        msAssert.areEqual("Medium Shading 1 Accent 1", table.getStyle().getName());
        msAssert.areEqual(TableStyleOptions.FIRST_COLUMN | TableStyleOptions.ROW_BANDS | TableStyleOptions.FIRST_ROW,
            table.getStyleOptions());
        msAssert.areEqual(189, (table.getFirstRow().getFirstCell().getCellFormat().getShading().getBackgroundPatternColor().getBlue() & 0xFF));
        msAssert.areEqual(Color.WHITE.getRGB(), table.getFirstRow().getFirstCell().getFirstParagraph().getRuns().get(0).getFont().getColor().getRGB());
        msAssert.areNotEqual(Color.LightBlue.getRGB(),
            (table.getLastRow().getFirstCell().getCellFormat().getShading().getBackgroundPatternColor().getBlue() & 0xFF));
        msAssert.areEqual(msColor.Empty.getRGB(), table.getLastRow().getFirstCell().getFirstParagraph().getRuns().get(0).getFont().getColor().getRGB());
    }

    @Test
    public void insertTableSetHeadingRow() throws Exception
    {
        //ExStart
        //ExFor:RowFormat.HeadingFormat
        //ExSummary:Shows how to build a table which include heading rows that repeat on subsequent pages. 
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
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
        for (int i = 0; i < 50; i++)
        {
            builder.insertCell();
            builder.getRowFormat().setHeadingFormat(false);
            builder.write("Column 1 Text");
            builder.insertCell();
            builder.write("Column 2 Text");
            builder.endRow();
        }

        doc.save(getArtifactsDir() + "Table.HeadingRow.doc");
        //ExEnd

        Assert.assertTrue(table.getFirstRow().getRowFormat().getHeadingFormat());
        Assert.assertTrue(table.getRows().get(1).getRowFormat().getHeadingFormat());
        Assert.assertFalse(table.getRows().get(2).getRowFormat().getHeadingFormat());
    }

    @Test
    public void insertTableWithPreferredWidth() throws Exception
    {
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

        doc.save(getArtifactsDir() + "Table.PreferredWidth.doc");
        //ExEnd

        // Verify the correct settings were applied
        msAssert.areEqual(PreferredWidthType.PERCENT, table.getPreferredWidth().getType());
        msAssert.areEqual(50, table.getPreferredWidth().getValue());
    }

    @Test
    public void insertCellsWithDifferentPreferredCellWidths() throws Exception
    {
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
        builder.getCellFormat().setPreferredWidth(PreferredWidth.fromPoints(40.0));
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.LightYellow);
        builder.writeln("Cell at 40 points width");

        PreferredWidth width = builder.getCellFormat().getPreferredWidth();
        msConsole.writeLine($"Width \"{width.GetHashCode()}\": {width.ToString()}");

        // Insert a relative (percent) sized cell
        builder.insertCell();
        builder.getCellFormat().setPreferredWidth(PreferredWidth.fromPercent(20.0));
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.LightBlue);
        builder.writeln("Cell at 20% width");

        // Each cell had its own PreferredWidth
        Assert.assertFalse(builder.getCellFormat().getPreferredWidth().equals(width));

        width = builder.getCellFormat().getPreferredWidth();
        msConsole.writeLine($"Width \"{width.GetHashCode()}\": {width.ToString()}");

        // Insert a auto sized cell
        builder.insertCell();
        builder.getCellFormat().setPreferredWidth(PreferredWidth.AUTO);
        builder.getCellFormat().getShading().setBackgroundPatternColor(msColor.getLightGreen());
        builder.writeln(
            "Cell automatically sized. The size of this cell is calculated from the table preferred width.");
        builder.writeln("In this case the cell will fill up the rest of the available space.");

        doc.save(getArtifactsDir() + "Table.CellPreferredWidths.docx");
        //ExEnd

        // Verify the correct settings were applied
        msAssert.areEqual(PreferredWidthType.POINTS, table.getFirstRow().getFirstCell().getCellFormat().getPreferredWidth().getType());
        msAssert.areEqual(PreferredWidthType.PERCENT, table.getFirstRow().getCells().get(1).getCellFormat().getPreferredWidth().getType());
        msAssert.areEqual(PreferredWidthType.AUTO, table.getFirstRow().getCells().get(2).getCellFormat().getPreferredWidth().getType());
    }

    @Test
    public void insertTableFromHtml() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the table from HTML. Note that AutoFitSettings does not apply to tables
        // inserted from HTML.
        builder.insertHtml("<table>" + "<tr>" + "<td>Row 1, Cell 1</td>" + "<td>Row 1, Cell 2</td>" + "</tr>" +
                           "<tr>" + "<td>Row 2, Cell 2</td>" + "<td>Row 2, Cell 2</td>" + "</tr>" + "</table>");

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertTableFromHtml.doc");

        // Verify the table was constructed properly
        msAssert.areEqual(1, doc.getChildNodes(NodeType.TABLE, true).getCount());
        msAssert.areEqual(2, doc.getChildNodes(NodeType.ROW, true).getCount());
        msAssert.areEqual(4, doc.getChildNodes(NodeType.CELL, true).getCount());
    }

    @Test
    public void buildNestedTableUsingDocumentBuilder() throws Exception
    {
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

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertNestedTable.doc");
        //ExEnd

        msAssert.areEqual(2, doc.getChildNodes(NodeType.TABLE, true).getCount());
        msAssert.areEqual(4, doc.getChildNodes(NodeType.CELL, true).getCount());
        msAssert.areEqual(1, cell.getTables().get(0).getCount());
        msAssert.areEqual(2, cell.getTables().get(0).getFirstRow().getCells().getCount());
    }

    @Test
    public void buildSimpleTable() throws Exception
    {
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
        builder.write("Row 2, Cell 1 Content");

        // Build the second cell.
        builder.insertCell();
        builder.write("Row 2, Cell 2 Content.");
        builder.endRow();

        // Signal that we have finished building the table
        builder.endTable();

        // Save the document to disk
        doc.save(getArtifactsDir() + "DocumentBuilder.CreateSimpleTable.doc");
        //ExEnd

        // Verify that the cell count of the table is four
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        Assert.assertNotNull(table);
        msAssert.areEqual(4, table.getChildNodes(NodeType.CELL, true).getCount());
    }

    @Test
    public void buildFormattedTable() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder
        //ExFor:DocumentBuilder.Write
        //ExFor:DocumentBuilder.InsertCell
        //ExFor:RowFormat.Height
        //ExFor:RowFormat.HeightRule
        //ExFor:Table.LeftIndent
        //ExFor:Shading.BackgroundPatternColor
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
        builder.getCellFormat().getShading().setBackgroundPatternColor(new Color((198), (217), (241)));
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.getFont().setSize(16.0);
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
        builder.getFont().setSize(12.0);
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

        doc.save(getArtifactsDir() + "DocumentBuilder.CreateFormattedTable.doc");
        //ExEnd

        // Verify that the cell style is different compared to default
        msAssert.areNotEqual(table.getLeftIndent(), 0.0);
        msAssert.areNotEqual(table.getFirstRow().getRowFormat().getHeightRule(), HeightRule.AUTO);
        msAssert.areNotEqual(table.getFirstRow().getFirstCell().getCellFormat().getShading().getBackgroundPatternColor(), msColor.Empty);
        msAssert.areNotEqual(table.getFirstRow().getFirstCell().getFirstParagraph().getParagraphFormat().getAlignment(),
            ParagraphAlignment.LEFT);
    }

    @Test
    public void setCellShadingAndBorders() throws Exception
    {
        //ExStart
        //ExFor:Shading
        //ExFor:Shading.BackgroundPatternColor
        //ExFor:Table.SetBorders
        //ExFor:BorderCollection.Left
        //ExFor:BorderCollection.Right
        //ExFor:BorderCollection.Top
        //ExFor:BorderCollection.Bottom
        //ExSummary:Shows how to format table and cell with different borders and shadings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();

        // Set the borders for the entire table
        table.setBorders(LineStyle.SINGLE, 2.0, Color.BLACK);
        // Set the cell shading for this cell
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.RED);
        builder.writeln("Cell #1");

        builder.insertCell();
        // Specify a different cell shading for the second cell
        builder.getCellFormat().getShading().setBackgroundPatternColor(msColor.getGreen());
        builder.writeln("Cell #2");

        // End this row.
        builder.endRow();

        // Clear the cell formatting from previous operations
        builder.getCellFormat().clearFormatting();

        // Create the second row
        builder.insertCell();

        // Create larger borders for the first cell of this row. This will be different
        // compared to the borders set for the table
        builder.getCellFormat().getBorders().getLeft().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getRight().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getTop().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getBottom().setLineWidth(4.0);
        builder.writeln("Cell #3");

        builder.insertCell();
        // Clear the cell formatting from the previous cell
        builder.getCellFormat().clearFormatting();
        builder.writeln("Cell #4");

        doc.save(getArtifactsDir() + "Table.SetBordersAndShading.doc");
        //ExEnd

        // Verify the table was created correctly
        msAssert.areEqual(Color.RED.getRGB(),
            table.getFirstRow().getFirstCell().getCellFormat().getShading().getBackgroundPatternColor().getRGB());
        msAssert.areEqual(msColor.getGreen().getRGB(),
            table.getFirstRow().getCells().get(1).getCellFormat().getShading().getBackgroundPatternColor().getRGB());
        msAssert.areEqual(msColor.getGreen().getRGB(),
            table.getFirstRow().getCells().get(1).getCellFormat().getShading().getBackgroundPatternColor().getRGB());
        msAssert.areEqual(msColor.Empty.getRGB(),
            table.getLastRow().getFirstCell().getCellFormat().getShading().getBackgroundPatternColor().getRGB());

        msAssert.areEqual(Color.BLACK.getRGB(), table.getFirstRow().getFirstCell().getCellFormat().getBorders().getLeft().getColor().getRGB());
        msAssert.areEqual(Color.BLACK.getRGB(), table.getFirstRow().getFirstCell().getCellFormat().getBorders().getLeft().getColor().getRGB());
        msAssert.areEqual(LineStyle.SINGLE, table.getFirstRow().getFirstCell().getCellFormat().getBorders().getLeft().getLineStyle());
        msAssert.areEqual(2.0, table.getFirstRow().getFirstCell().getCellFormat().getBorders().getLeft().getLineWidth());
        msAssert.areEqual(4.0, table.getLastRow().getFirstCell().getCellFormat().getBorders().getLeft().getLineWidth());
    }

    @Test
    public void setPreferredTypeConvertUtil() throws Exception
    {
        //ExStart
        //ExFor:PreferredWidth.FromPoints
        //ExSummary:Shows how to specify a cell preferred width by converting inches to points.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.getCellFormat().setPreferredWidth(PreferredWidth.fromPoints(ConvertUtil.inchToPoint(3.0)));
        builder.insertCell();
        //ExEnd

        msAssert.areEqual(216.0, table.getFirstRow().getFirstCell().getCellFormat().getPreferredWidth().getValue());
    }

    @Test
    public void insertHyperlinkToLocalBookmark() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.StartBookmark
        //ExFor:DocumentBuilder.EndBookmark
        //ExFor:DocumentBuilder.InsertHyperlink
        //ExSummary:Inserts a hyperlink referencing local bookmark.
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

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertHyperlinkToLocalBookmark.doc");
        //ExEnd
    }

    @Test
    public void documentBuilderCtor() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Hello World!");
    }

    @Test
    public void documentBuilderCursorPosition() throws Exception
    {
        Document doc = new Document(getMyDir() + "DocumentBuilder.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        Node curNode = builder.getCurrentNode();
        Paragraph curParagraph = builder.getCurrentParagraph();
    }

    @Test
    public void documentBuilderMoveToNode() throws Exception
    {
        //ExStart
        //ExFor:Story.LastParagraph
        //ExFor:DocumentBuilder.MoveTo(Node)
        //ExSummary:Shows how to move a cursor position to a specified node.
        Document doc = new Document(getMyDir() + "DocumentBuilder.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveTo(doc.getFirstSection().getBody().getLastParagraph());
        //ExEnd
    }

    @Test
    public void documentBuilderMoveToDocumentStartEnd() throws Exception
    {
        Document doc = new Document(getMyDir() + "DocumentBuilder.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToDocumentEnd();
        builder.writeln("This is the end of the document.");

        builder.moveToDocumentStart();
        builder.writeln("This is the beginning of the document.");
    }

    @Test
    public void documentBuilderMoveToSection() throws Exception
    {
        Document doc = new Document(getMyDir() + "DocumentBuilder.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Parameters are 0-index. Moves to third section
        builder.moveToSection(2);
        builder.writeln("This is the 3rd section.");
    }

    @Test
    public void documentBuilderMoveToParagraph() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.MoveToParagraph
        //ExSummary:Shows how to move a cursor position to the specified paragraph.
        Document doc = new Document(getMyDir() + "DocumentBuilder.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Parameters are 0-index. Moves to third paragraph
        builder.moveToParagraph(2, 0);
        builder.writeln("This is the 3rd paragraph.");
        //ExEnd
    }

    @Test
    public void documentBuilderMoveToTableCell() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.MoveToCell
        //ExSummary:Shows how to move a cursor position to the specified table cell.
        Document doc = new Document(getMyDir() + "DocumentBuilder.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // All parameters are 0-index. Moves to the 2nd table, 3rd row, 5th cell
        builder.moveToCell(1, 2, 4, 0);
        builder.writeln("Hello World!");
        //ExEnd
    }

    @Test
    public void documentBuilderMoveToBookmark() throws Exception
    {
        Document doc = new Document(getMyDir() + "DocumentBuilder.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToBookmark("CoolBookmark");
        builder.writeln("This is a very cool bookmark.");
    }

    @Test
    public void documentBuilderMoveToBookmarkEnd() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.MoveToBookmark(String, Boolean, Boolean)
        //ExSummary:Shows how to move a cursor position to just after the bookmark end.
        Document doc = new Document(getMyDir() + "DocumentBuilder.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToBookmark("CoolBookmark", false, true);
        builder.writeln("This is a very cool bookmark.");
        //ExEnd
    }

    @Test
    public void documentBuilderMoveToMergeField() throws Exception
    {
        Document doc = new Document(getMyDir() + "DocumentBuilder.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToMergeField("NiceMergeField");
        builder.writeln("This is a very nice merge field.");
    }

    @Test
    public void documentBuilderInsertParagraph() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertParagraph
        //ExFor:ParagraphFormat.FirstLineIndent
        //ExFor:ParagraphFormat.Alignment
        //ExFor:ParagraphFormat.KeepTogether
        //ExFor:ParagraphFormat.AddSpaceBetweenFarEastAndAlpha
        //ExFor:ParagraphFormat.AddSpaceBetweenFarEastAndDigit
        //ExFor:Paragraph.IsEndOfDocument
        //ExSummary:Shows how to insert a paragraph into the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Specify font formatting
        Font font = builder.getFont();
        font.setSize(16.0);
        font.setBold(true);
        font.setColor(Color.BLUE);
        font.setName("Arial");
        font.setUnderline(Underline.DASH);

        // Specify paragraph formatting
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();
        paragraphFormat.setFirstLineIndent(8.0);
        paragraphFormat.setAlignment(ParagraphAlignment.JUSTIFY);
        paragraphFormat.setAddSpaceBetweenFarEastAndAlpha(true);
        paragraphFormat.setAddSpaceBetweenFarEastAndDigit(true);
        paragraphFormat.setKeepTogether(true);

        builder.writeln("A whole paragraph.");

        // We can use this flag to ensure that we're at the end of the document
        Assert.assertTrue(builder.getCurrentParagraph().isEndOfDocument());
        //ExEnd
    }

    @Test
    public void documentBuilderBuildTable() throws Exception
    {
        //ExStart
        //ExFor:Table
        //ExFor:DocumentBuilder.StartTable
        //ExFor:DocumentBuilder.InsertCell
        //ExFor:DocumentBuilder.EndRow
        //ExFor:DocumentBuilder.EndTable
        //ExFor:DocumentBuilder.CellFormat
        //ExFor:DocumentBuilder.RowFormat
        //ExFor:DocumentBuilder.Write
        //ExFor:DocumentBuilder.Writeln(String)
        //ExFor:RowFormat.Height
        //ExFor:RowFormat.HeightRule
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
        // Use fixed column widths
        table.autoFit(AutoFitBehavior.FIXED_COLUMN_WIDTHS);

        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
        builder.write("This is row 1 cell 1");

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
        builder.writeln("This is row 2 cell 1");

        // Insert a cell
        builder.insertCell();
        builder.getCellFormat().setOrientation(TextOrientation.DOWNWARD);
        builder.writeln("This is row 2 cell 2");

        builder.endRow();

        builder.endTable();
        //ExEnd
    }

    @Test
    public void tableCellVerticalRotatedFarEastTextOrientation() throws Exception
    {
        Document doc = new Document(getMyDir() + "DocumentBuilder.TableCellVerticalRotatedFarEastTextOrientation.docx");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        Cell cell = table.getFirstRow().getFirstCell();

        msAssert.areEqual(TextOrientation.VERTICAL_ROTATED_FAR_EAST, cell.getCellFormat().getOrientation());

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        cell = table.getFirstRow().getFirstCell();

        msAssert.areEqual(TextOrientation.VERTICAL_ROTATED_FAR_EAST, cell.getCellFormat().getOrientation());
    }

    @Test
    public void documentBuilderInsertBreak() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("This is page 1.");
        builder.insertBreak(BreakType.PAGE_BREAK);

        builder.writeln("This is page 2.");
        builder.insertBreak(BreakType.PAGE_BREAK);

        builder.writeln("This is page 3.");
    }

    @Test
    public void documentBuilderInsertInlineImage() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertImage(getImageDir() + "Watermark.png");
    }

    @Test
    public void documentBuilderInsertFloatingImage() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert a floating image from a file or URL.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertImage(getImageDir() + "Watermark.png", RelativeHorizontalPosition.MARGIN, 100.0,
            RelativeVerticalPosition.MARGIN, 100.0, 200.0, 100.0, WrapType.SQUARE);
        //ExEnd
    }

    @Test
    public void insertImageFromUrl() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String)
        //ExSummary:Shows how to insert an image into a document from a web address.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertImage(getAsposeLogoUrl());

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertImageFromUrl.doc");
        //ExEnd

        // Verify that the image was inserted into the document
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        Assert.assertNotNull(shape);
        Assert.assertTrue(shape.hasImage());
    }

    @Test
    public void documentBuilderInsertImageSourceSize() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert a floating image from a file or URL and retain the original image size in the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Pass a negative value to the width and height values to specify using the size of the source image
        builder.insertImage(getImageDir() + "LogoSmall.png", RelativeHorizontalPosition.MARGIN, 200.0,
            RelativeVerticalPosition.MARGIN, 100.0, -1, -1, WrapType.SQUARE);
        //ExEnd

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertImageOriginalSize.doc");
    }

    @Test
    public void documentBuilderInsertBookmark() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("FineBookmark");
        builder.writeln("This is just a fine bookmark.");
        builder.endBookmark("FineBookmark");
    }

    @Test
    public void documentBuilderInsertTextInputFormField() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertTextInput
        //ExSummary:Shows how to insert a text input form field into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertTextInput("TextInput", TextFormFieldType.REGULAR, "", "Hello", 0);
        //ExEnd
    }

    @Test
    public void documentBuilderInsertComboBoxFormField() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertComboBox
        //ExSummary:Shows how to insert a combobox form field into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        String[] items = { "One", "Two", "Three" };
        builder.insertComboBox("DropDown", items, 0);
        //ExEnd
    }

    @Test
    public void documentBuilderInsertToc() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table of contents at the beginning of the document
        builder.insertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // The newly inserted table of contents will be initially empty
        // It needs to be populated by updating the fields in the document
        doc.updateFields();
    }

    @Test (description = "WORDSNET-16868")
    public void createAndSignSignatureLineUsingProviderId() throws Exception
    {
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

        SignatureLineOptions signatureLineOptions = new SignatureLineOptions();
        {
            signatureLineOptions.setSigner("vderyushev");
            signatureLineOptions.setSignerTitle("QA");
            signatureLineOptions.setEmail("vderyushev@aspose.com");
            signatureLineOptions.setShowDate(true);
            signatureLineOptions.setDefaultInstructions(false);
            signatureLineOptions.setInstructions("You need more info about signature line");
            signatureLineOptions.setAllowComments(true);
        }

        SignatureLine signatureLine = builder.insertSignatureLine(signatureLineOptions).getSignatureLine();
        signatureLine.setProviderIdInternal(Guid.parse("CF5A7BB4-8F3C-4756-9DF6-BEF7F13259A2"));
        
        doc.save(getArtifactsDir() + "DocumentBuilder.SignatureLineProviderId In.docx");

        SignOptions signOptions = new SignOptions();
        signOptions.setSignatureLineIdInternal(signatureLine.getIdInternal());
        signOptions.setProviderIdInternal(signatureLine.getProviderIdInternal());
        signOptions.setComments("Document was signed by vderyushev");
        signOptions.setSignTimeInternal(DateTime.getNow());

        CertificateHolder certHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        DigitalSignatureUtil.sign(getArtifactsDir() + "DocumentBuilder.SignatureLineProviderId In.docx", getArtifactsDir() + "DocumentBuilder.SignatureLineProviderId Out.docx", certHolder, signOptions);
        //ExEnd

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "DocumentBuilder.SignatureLineProviderId Out.docx", getGoldsDir() + "DocumentBuilder.SignatureLineProviderId Gold.docx"));
    }

    @Test
    public void insertSignatureLineCurrentPosition() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertSignatureLine(SignatureLineOptions, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, WrapType)
        //ExSummary:Shows how to insert signature line at the specified position.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        SignatureLineOptions options = new SignatureLineOptions();
        {
            options.setSigner("John Doe");
            options.setSignerTitle("Manager");
            options.setEmail("johndoe@aspose.com");
            options.setShowDate(true);
            options.setDefaultInstructions(false);
            options.setInstructions("You need more info about signature line");
            options.setAllowComments(true);
        }

        builder.insertSignatureLine(options, RelativeHorizontalPosition.RIGHT_MARGIN, 2.0,
            RelativeVerticalPosition.PAGE, 3.0, WrapType.INLINE);
        //ExEnd

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        SignatureLine signatureLine = shape.getSignatureLine();

        msAssert.areEqual("John Doe", signatureLine.getSigner());
        msAssert.areEqual("Manager", signatureLine.getSignerTitle());
        msAssert.areEqual("johndoe@aspose.com", signatureLine.getEmail());
        msAssert.areEqual(true, signatureLine.getShowDate());
        msAssert.areEqual(false, signatureLine.getDefaultInstructions());
        msAssert.areEqual("You need more info about signature line", signatureLine.getInstructions());
        msAssert.areEqual(true, signatureLine.getAllowComments());
        msAssert.areEqual(false, signatureLine.isSigned());
        msAssert.areEqual(false, signatureLine.isValid());
    }

    @Test
    public void documentBuilderSetFontFormatting() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set font formatting properties
        Font font = builder.getFont();
        font.setBold(true);
        font.setColor(msColor.getDarkBlue());
        font.setItalic(true);
        font.setName("Arial");
        font.setSize(24.0);
        font.setSpacing(5.0);
        font.setUnderline(Underline.DOUBLE);

        // Output formatted text
        builder.writeln("I'm a very nice formatted String.");
    }

    @Test
    public void documentBuilderSetParagraphFormatting() throws Exception
    {
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
            "I'm a very nice formatted paragraph. I'm intended to demonstrate how the left and right indents affect word wrapping.");
        builder.writeln(
            "I'm another nice formatted paragraph. I'm intended to demonstrate how the space after paragraph looks like.");
        //ExEnd
    }

    @Test
    public void documentBuilderSetCellFormatting() throws Exception
    {
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

        builder.writeln("I'm a wonderful formatted cell.");

        builder.endRow();
        builder.endTable();
        //ExEnd
    }

    @Test
    public void documentBuilderSetRowFormatting() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.RowFormat
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

        builder.writeln("I'm a wonderful formatted row.");

        builder.endRow();
        builder.endTable();
        //ExEnd
    }

    @Test
    public void documentBuilderSetListFormatting() throws Exception
    {
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
    public void documentBuilderSetSectionFormatting() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set page properties
        builder.getPageSetup().setOrientation(Orientation.LANDSCAPE);
        builder.getPageSetup().setLeftMargin(50.0);
        builder.getPageSetup().setPaperSize(PaperSize.PAPER_10_X_14);
    }

    @Test
    public void insertFootnote() throws Exception
    {
        //ExStart
        //ExFor:FootnoteType
        //ExFor:Document.FootnoteOptions
        //ExFor:DocumentBuilder.InsertFootnote(FootnoteType,String)
        //ExFor:DocumentBuilder.InsertFootnote(FootnoteType,String,String)
        //ExSummary:Shows how to add a footnote to a paragraph in the document using DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i <= 100; i++)
        {
            builder.write("Some text " + i);

            builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote text " + i);
            builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote text " + i, "242");
        }
        //ExEnd

        msAssert.areEqual("Footnote text 0",
            msString.trim(doc.getChildNodes(NodeType.FOOTNOTE, true).get(0).toString(SaveFormat.TEXT)));

        doc.getFootnoteOptions().setNumberStyle(NumberStyle.ARABIC);
        doc.getFootnoteOptions().setStartNumber(1);
        doc.getFootnoteOptions().setRestartRule(FootnoteNumberingRule.RESTART_PAGE);

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertFootnote.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "DocumentBuilder.InsertFootnote.docx", getGoldsDir() + "DocumentBuilder.InsertFootnote Gold.docx"));
    }

    @Test
    public void documentBuilderApplyParagraphStyle() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set paragraph style
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.TITLE);

        builder.write("Hello");
    }

    @Test
    public void documentBuilderApplyBordersAndShading() throws Exception
    {
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
        shading.setBackgroundPatternColor(msColor.getLightCoral());
        shading.setForegroundPatternColor(Color.LightSalmon);

        builder.write("I'm a formatted paragraph with double border and nice shading.");
        //ExEnd
    }

    @Test
    public void deleteRow() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.DeleteRow
        //ExSummary:Shows how to delete a row from a table.
        Document doc = new Document(getMyDir() + "DocumentBuilder.DocWithTable.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Delete the first row of the first table in the document
        builder.deleteRow(0, 0);
        //ExEnd
    }

    @Test (enabled = false, description = "Bug: does not insert headers and footers, all lists (bullets, numbering, multilevel) breaks")
    public void insertDocument() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertDocument(Document, ImportFormatMode)
        //ExFor:ImportFormatMode
        //ExSummary:Shows how to insert a document content into another document keep formatting of inserted document.
        Document doc = new Document(getMyDir() + "Document.docx");

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveToDocumentEnd();
        builder.insertBreak(BreakType.PAGE_BREAK);

        Document docToInsert = new Document(getMyDir() + "DocumentBuilder.KeepSourceFormatting.docx");

        builder.insertDocument(docToInsert, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        builder.getDocument().save(getArtifactsDir() + "DocumentBuilder.InsertDocument.docx");
        //ExEnd

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "DocumentBuilder.InsertDocument.docx", getGoldsDir() + "DocumentBuilder.InsertDocument Gold.docx"));
    }

    @Test
    public void keepSourceNumbering() throws Exception
    {
        //ExStart
        //ExFor:ImportFormatOptions.KeepSourceNumbering
        //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how the numbering will be imported when it clashes in source and destination documents.
        Document dstDoc = new Document(getMyDir() + "DocumentBuilder.KeepSourceNumbering.DestinationDocument.docx");
        Document srcDoc = new Document(getMyDir() + "DocumentBuilder.KeepSourceNumbering.SourceDocument.docx");
        
        ImportFormatOptions importFormatOptions = new ImportFormatOptions();
        // Keep source list formatting when importing numbered paragraphs
        importFormatOptions.setKeepSourceNumbering(true);
        
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING, importFormatOptions);
        
        ParagraphCollection srcParas = srcDoc.getFirstSection().getBody().getParagraphs();
        for (Node node : (Iterable<Node>) srcParas)
        {
            Paragraph srcPara = (Paragraph) node;
            Node importedNode = importer.importNode(srcPara, true);
            dstDoc.getFirstSection().getBody().appendChild(importedNode);
        }
 
        dstDoc.save(getArtifactsDir() + "DocumentBuilder.KeepSourceNumbering.ResultDocument.docx");
        //ExEnd
    }

    @Test
    public void ignoreTextBoxes() throws Exception
    {
        //ExStart
        //ExFor:ImportFormatOptions.IgnoreTextBoxes
        //ExSummary:Shows how to manage formatting in the text boxes of the source destination during the import.
        Document dstDoc = new Document(getMyDir() + "DocumentBuilder.IgnoreTextBoxes.DestinationDocument.docx");
        Document srcDoc = new Document(getMyDir() + "DocumentBuilder.IgnoreTextBoxes.SourceDocument.docx");
        
        ImportFormatOptions importFormatOptions = new ImportFormatOptions();
        // Keep the source text boxes formatting when importing
        importFormatOptions.setIgnoreTextBoxes(false);

        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING, importFormatOptions);
 
        ParagraphCollection srcParas = srcDoc.getFirstSection().getBody().getParagraphs();
        for (Node node : (Iterable<Node>) srcParas)
        {
            Paragraph srcPara = (Paragraph) node;
            Node importedNode = importer.importNode(srcPara, true);
            dstDoc.getFirstSection().getBody().appendChild(importedNode);
        }
 
        dstDoc.save(getArtifactsDir() + "DocumentBuilder.IgnoreTextBoxes.ResultDocument.docx");
        //ExEnd
    }

    @Test
    public void moveToFieldEx() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.MoveToField
        //ExSummary:Shows how to move document builder's cursor to a specific field.
        Document doc = new Document(getMyDir() + "Document.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        Field field = builder.insertField("MERGEFIELD field");

        builder.moveToField(field, true);
        //ExEnd
    }          

    @Test
    public void insertOleObjectException() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Assert.That(() => builder.insertOleObject("", "checkbox", false, true, null),
            Throws.<IllegalArgumentException>TypeOf());
    }

    @Test
    public void insertChartDouble() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertChart(ChartType, Double, Double)
        //ExSummary:Shows how to insert a chart into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertChart(ChartType.PIE, ConvertUtil.pixelToPoint(300.0), ConvertUtil.pixelToPoint(300.0));

        doc.save(getArtifactsDir() + "Document.InsertedChartDouble.doc");
        //ExEnd
    }

    @Test
    public void insertChartRelativePosition() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertChart(ChartType, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert a chart into a document and specify position and size.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertChart(ChartType.PIE, RelativeHorizontalPosition.MARGIN, 100.0, RelativeVerticalPosition.MARGIN,
            100.0, 200.0, 100.0, WrapType.SQUARE);

        doc.save(getArtifactsDir() + "Document.InsertedChartRelativePosition.doc");
        //ExEnd
    }

    @Test
    public void insertFieldFieldType() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertField(FieldType, Boolean)
        //ExSummary:Shows how to insert a field into a document using FieldType.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("This field was inserted/updated at ");
        builder.insertField(FieldType.FIELD_TIME, true);

        doc.save(getArtifactsDir() + "Document.InsertedField.doc");
        //ExEnd
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
    public void fieldResultFormatting() throws Exception
    {
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
        ((FieldResultFormatter)doc.getFieldOptions().getResultFormatter()).printInvocations();

        // Our formatter has also overridden the formats that were originally applied in the fields
        msAssert.areEqual("$5", doc.getRange().getFields().get(0).getResult());
        Assert.assertTrue(doc.getRange().getFields().get(1).getResult().startsWith("Date: "));
        msAssert.areEqual("Item # 2:", doc.getRange().getFields().get(2).getResult());
    }

    /// <summary>
    /// Custom IFieldResult implementation that applies formats and tracks format invocations
    /// </summary>
    private static class FieldResultFormatter implements IFieldResultFormatter
    {
        public FieldResultFormatter(String numberFormat, String dateFormat, String generalFormat)
        {
            mNumberFormat = numberFormat;
            mDateFormat = dateFormat;
            mGeneralFormat = generalFormat;
        }

        public String formatNumeric(double value, String format)
        {
            msArrayList.add(mNumberFormatInvocations, new Object[] { value, format });

            return msString.isNullOrEmpty(mNumberFormat) ? null : msString.format(mNumberFormat, value);
        }

        public String formatDateTime(DateTime value, String format, /*CalendarType*/int calendarType)
        {
            msArrayList.add(mDateFormatInvocations, new Object[] { value, format, calendarType });

            return msString.isNullOrEmpty(mDateFormat) ? null : msString.format(mDateFormat, value);
        }

        public String format(String value, /*GeneralFormat*/int format)
        {
            return format((Object)value, format);
        }

        public String format(double value, /*GeneralFormat*/int format)
        {
            return format((Object)value, format);
        }

        private String format(Object value, /*GeneralFormat*/int format)
        {
            msArrayList.add(mGeneralFormatInvocations, new Object[] { value, format });

            return msString.isNullOrEmpty(mGeneralFormat) ? null : msString.format(mGeneralFormat, value);
        }

        public void printInvocations()
        {
            msConsole.writeLine("Number format invocations ({0}):", mNumberFormatInvocations.size());
            for (Object[] s : (Iterable<Object[]>) mNumberFormatInvocations)
            {
                msConsole.writeLine("\tValue: " + s[0] + ", original format: " + s[1]);
            }

            msConsole.writeLine("Date format invocations ({0}):", mDateFormatInvocations.size());
            for (Object[] s : (Iterable<Object[]>) mDateFormatInvocations)
            {
                msConsole.writeLine("\tValue: " + s[0] + ", original format: " + s[1] + ", calendar type: " + s[2]);
            }

            msConsole.writeLine("General format invocations ({0}):", mGeneralFormatInvocations.size());
            for (Object[] s : (Iterable<Object[]>) mGeneralFormatInvocations)
            {
                msConsole.writeLine("\tValue: " + s[0] + ", original format: " + s[1]);
            }
        }

        private /*final*/ String mNumberFormat;
        private /*final*/ String mDateFormat;
        private /*final*/ String mGeneralFormat;

        private /*final*/ ArrayList mNumberFormatInvocations = new ArrayList();
        private /*final*/ ArrayList mDateFormatInvocations = new ArrayList();
        private /*final*/ ArrayList mGeneralFormatInvocations = new ArrayList();

    }
    //ExEnd

    @Test
    public void insertVideoWithUrl() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertOnlineVideo(String, Double, Double)
        //ExSummary:Show how to insert online video into a document using video url
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Pass direct url from youtu.be.
        final String URL = "https://youtu.be/t_1LYZ102RA";

        final double WIDTH = 360.0;
        final double HEIGHT = 270.0;

        builder.insertOnlineVideo(URL, WIDTH, HEIGHT);
        //ExEnd
    }

    @Test
    public void insertUnderline() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.Underline
        //ExSummary:Shows how to set and edit a document builder's underline.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a new style for our underline
        builder.setUnderline(Underline.DASH);

        // Same object as DocumentBuilder.Font.Underline
        msAssert.areEqual(builder.getUnderline(), builder.getFont().getUnderline());
        msAssert.areEqual(Underline.DASH, builder.getFont().getUnderline());

        // These properties will be applied to the underline as well
        builder.getFont().setColor(Color.BLUE);
        builder.getFont().setSize(32.0);

        builder.writeln("Underlined text.");

        doc.save(getArtifactsDir() + "DocumentBuilder.Underline.docx");         
        //ExEnd
    }

    @Test
    public void addTextToCurrentStory() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.CurrentStory
        //ExSummary:Shows how to work with a document builder's current story.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The body of the current section is the same object as the current story
        msAssert.areEqual(builder.getCurrentStory(), doc.getFirstSection().getBody());
        msAssert.areEqual(builder.getCurrentStory(), builder.getCurrentParagraph().getParentNode());

        msAssert.areEqual(StoryType.MAIN_TEXT, builder.getCurrentStory().getStoryType());

        builder.getCurrentStory().appendParagraph("Text added to current Story.");

        // A story can contain tables too
        Table table = builder.startTable();

        builder.insertCell();
        builder.write("This is row 1 cell 1");
        builder.insertCell();
        builder.write("This is row 1 cell 2");

        builder.endRow();

        builder.insertCell();
        builder.writeln("This is row 2 cell 1");
        builder.insertCell();
        builder.writeln("This is row 2 cell 2");

        builder.endRow();
        builder.endTable();

        // The table we just made is automatically placed in the story
        Assert.assertTrue(builder.getCurrentStory().getTables().contains(table));

        doc.save(getArtifactsDir() + "DocumentBuilder.CurrentStory.docx");
        //ExEnd
    }

    @Test
    public void builderInsertOleObject() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertOleObject(Stream, String, Boolean, Image)
        //ExSummary:Shows how to use document builder to embed Ole objects in a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Let's take a spreadsheet from our system and insert it into the document
        Stream spreadsheetStream = File.open(getMyDir() + "MySpreadsheet.xlsx", FileMode.OPEN);
        try /*JAVA: was using*/
        {
            // The spreadsheet can be activated by double clicking the panel that you'll see in the document immediately under the text we will add
            // We did not set the area to double click as an icon nor did we change its appearance so it looks like a simple panel
            builder.writeln("Spreadsheet Ole object:");
            builder.insertOleObjectInternal(spreadsheetStream, "MyOleObject.xlsx", false, null);

            // A powerpoint presentation is another type of object we can embed in our document
            // This time we'll also exercise some control over how it looks 
            Stream powerpointStream = File.open(getMyDir() + "MyPresentation.pptx", FileMode.OPEN);
            try /*JAVA: was using*/
            {
                // If we insert the Ole object as an icon, we are still provided with a default icon
                // If that is not suitable, we can make the icon to look like any image
                WebClient webClient = new WebClient();
                try /*JAVA: was using*/
                {
                    byte[] imgBytes = webClient.DownloadData(getAsposeLogoUrl());

                                                                
                    MemoryStream stream = new MemoryStream(imgBytes);
                    try /*JAVA: was using*/
                    {
                        BufferedImage image = BufferedImage.FromStream(stream);
                        try /*JAVA: was using*/
                        {
                            // If we double click the image, the powerpoint presentation will open
                            builder.insertParagraph();
                            builder.writeln("Powerpoint Ole object:");
                            builder.insertOleObjectInternal(powerpointStream, "MyOleObject.pptx", true, image);
                        }
                        finally { if (image != null) image.flush(); }
                    }
                    finally { if (stream != null) stream.close(); }

                                    }
                finally { if (webClient != null) webClient.close(); }
            }
            finally { if (powerpointStream != null) powerpointStream.close(); }
        }
        finally { if (spreadsheetStream != null) spreadsheetStream.close(); }

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertOleObject.docx");
        //ExEnd
    }

    @Test
    public void builderInsertStyleSeparator() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertStyleSeparator
        //ExSummary:Shows how to use and separate multiple styles in a paragraph
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("This text is in the default style. ");

        builder.insertStyleSeparator();

        // Create a custom style
        Style myStyle = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "MyStyle");
        myStyle.getFont().setSize(14.0);
        myStyle.getFont().setName("Courier New");
        myStyle.getFont().setColor(Color.BLUE);

        // Append text with custom style
        builder.getParagraphFormat().setStyleName(myStyle.getName());
        builder.write("This is text in the same paragraph but with my custom style.");

        doc.save(getArtifactsDir() + "DocumentBuilder.StyleSeparator.docx");
        //ExEnd
    }

    @Test
    public void insertStyleSeparator() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertStyleSeparator
        //ExSummary:Shows how to separate styles from two different paragraphs used in one logical printed paragraph.
        DocumentBuilder builder = new DocumentBuilder(new Document());

        Style paraStyle = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "MyParaStyle");
        paraStyle.getFont().setBold(false);
        paraStyle.getFont().setSize(8.0);
        paraStyle.getFont().setName("Arial");

        // Append text with "Heading 1" style
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.write("Heading 1");
        builder.insertStyleSeparator();

        // Append text with another style
        builder.getParagraphFormat().setStyleName(paraStyle.getName());
        builder.write("This is text with some other formatting ");
        //ExEnd

        builder.getDocument().save(getArtifactsDir() + "DocumentBuilder.InsertStyleSeparator.docx");
    }

    @Test
    public void withoutStyleSeparator() throws Exception
    {
        DocumentBuilder builder = new DocumentBuilder(new Document());

        Style paraStyle = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "MyParaStyle");
        paraStyle.getFont().setBold(false);
        paraStyle.getFont().setSize(8.0);
        paraStyle.getFont().setName("Arial");

        // Append text with "Heading 1" style
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.write("Heading 1");

        // Append text with another style
        builder.getParagraphFormat().setStyleName(paraStyle.getName());
        builder.write("This is text with some other formatting ");

        builder.getDocument().save(getArtifactsDir() + "DocumentBuilder.InsertTextWithoutStyleSeparator.docx");
    }

    @Test
    public void resolveStyleBehaviorWhileInsertDocument() throws Exception
    {
        //ExStart
        //ExFor:ImportFormatOptions
        //ExFor:ImportFormatOptions.SmartStyleBehavior
        //ExFor:DocumentBuilder.InsertDocument(Document, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how to resolve styles behavior while inserting documents.
        Document destDoc = new Document(getMyDir() + "DocumentBuilder.SmartStyleBehavior.DestinationDocument.docx");
        Document sourceDoc1 = new Document(getMyDir() + "DocumentBuilder.SmartStyleBehavior.SourceDocument01.docx");
        Document sourceDoc2 = new Document(getMyDir() + "DocumentBuilder.SmartStyleBehavior.SourceDocument02.docx");

        DocumentBuilder builder = new DocumentBuilder(destDoc);

        builder.moveToDocumentEnd();
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.moveToDocumentEnd();

        ImportFormatOptions importFormatOptions = new ImportFormatOptions();
        importFormatOptions.setSmartStyleBehavior(true);
        
        // When SmartStyleBehavior is enabled,
        // a source style will be expanded into a direct attributes inside a destination document,
        // if KeepSourceFormatting importing mode is used
        builder.insertDocument(sourceDoc1, ImportFormatMode.KEEP_SOURCE_FORMATTING, importFormatOptions);
        
        builder.moveToDocumentEnd();
        builder.insertBreak(BreakType.PAGE_BREAK);
        
        // When SmartStyleBehavior is disabled,
        // a source style will be expanded only if it is numbered.
        // Existing destination attributes will not be overridden, including lists
        builder.insertDocument(sourceDoc2, ImportFormatMode.USE_DESTINATION_STYLES);

        destDoc.save(getArtifactsDir() + "DocumentBuilder.SmartStyleBehavior.ResultDocument.docx");
        //ExEnd
    }

    @Test
    public void resolveStyleBehaviorWhileAppendDocument() throws Exception
    {
        //ExStart
        //ExFor:Document.AppendDocument(Document, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how to resolve styles behavior while append document.
        Document srcDoc = new Document(getMyDir() + "DocumentBuilder.ResolveStyleBehaviorWhileAppendDocument.Source.docx");
        Document dstDoc = new Document(getMyDir() + "DocumentBuilder.ResolveStyleBehaviorWhileAppendDocument.Destination.docx");

        ImportFormatOptions options = new ImportFormatOptions();
        // Specify that if numbering clashes in source and destination documents
        // then a numbering from the source document will be used
        options.setKeepSourceNumbering(true);
        dstDoc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES, options);
        dstDoc.updateListLabels();
        //ExEnd

        Paragraph para = dstDoc.getSections().get(1).getBody().getLastParagraph();
        String paraText = para.getText();

        msAssert.areEqual("1.", para.getListLabel().getLabelString());
        msAssert.isTrue(paraText.startsWith("13->13"), paraText);
    }

    @Test
    public void markdownTest() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
    
        builder.getParagraphFormat().setStyleName("Quote");
        builder.writeln("12345");
        
        Style quoteLevel2 = doc.getStyles().add(StyleType.PARAGRAPH, "Quote1");
        builder.getParagraphFormat().setStyle(quoteLevel2);
        doc.getStyles().get("Quote1").setBaseStyleName("Quote");
        builder.writeln("123456");
                
        // Save to md file
        Document mdDoc = saveOpen(doc, getArtifactsDir() + "AddedParagraphStyle.md");

        ParagraphCollection paragraphs = mdDoc.getFirstSection().getBody().getParagraphs();

        for (Paragraph paragraph : (Iterable<Paragraph>) paragraphs)
        {
            if (paragraph.getRuns().getCount() != 0)
            {
                if ("Blockquote 1".equals(paragraph.getRuns().get(0).getText()))
                {
                    msAssert.areEqual("Quote1", paragraph.getParagraphFormat().getStyle().getName());
                }
            }
        }
    }
    
    private static Document saveOpen(Document doc, String path) throws Exception
    {
        doc.save(path);
        return new Document(path);
    }

            }
