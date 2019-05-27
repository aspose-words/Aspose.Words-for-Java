package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.Paragraph;
import com.aspose.words.FieldType;
import com.aspose.words.Run;
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;
import com.aspose.ms.System.DateTime;
import com.aspose.ms.System.msString;
import com.aspose.words.ParagraphCollection;
import com.aspose.ms.System.msConsole;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.ParagraphFormat;
import com.aspose.words.Node;


@Test
class ExParagraph !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void insertField() throws Exception
    {
        //ExStart
        //ExFor:Paragraph.InsertField(string, Node, bool)
        //ExFor:Paragraph.InsertField(FieldType, bool, Node, bool)
        //ExFor:Paragraph.InsertField(string, string, Node, bool)
        //ExSummary:Shows how to insert field using several methods: "field code", "field code and field value", "field code and field value after a run of text"
        Document doc = new Document();

        // Get the first paragraph of the document
        Paragraph para = doc.getFirstSection().getBody().getFirstParagraph();

        // Inserting field using field code
        // Note: All methods support inserting field after some node. Just set "true" in the "isAfter" parameter
        para.insertField(" AUTHOR ", null, false);

        // Using field type
        // Note:
        // 1. For inserting field using field type, you can choose, update field before or after you open the document ("updateField" parameter)
        // 2. For other methods it's works automatically
        para.insertField(FieldType.FIELD_AUTHOR, false, null, true);

        // Using field code and field value
        para.insertField(" AUTHOR ", "Test Field Value", null, false);

        // Add a run of text
        Run run = new Run(doc); { run.setText(" Hello World!"); }
        para.appendChild(run);

        // Using field code and field value before a run of text
        // Note: For inserting field before/after a run of text you can use all methods above, just add ref on your text ("refNode" parameter)
        para.insertField(" AUTHOR ", "Test Field Value", run, false);
        //ExEnd
    }

    @Test
    public void insertFieldBeforeTextInParagraph() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        insertFieldUsingFieldCode(doc, " AUTHOR ", null, false, 1);

        msAssert.areEqual("\u0013 AUTHOR \u0014Test Author\u0015Hello World!\r",
            DocumentHelper.getParagraphText(doc, 1));
    }

    @Test
    public void insertFieldAfterTextInParagraph() throws Exception
    {
        String date = DateTime.getToday().toString("d");

        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        insertFieldUsingFieldCode(doc, " DATE ", null, true, 1);

        msAssert.areEqual(msString.format("Hello World!\u0013 DATE \u0014{0}\u0015\r", date),
            DocumentHelper.getParagraphText(doc, 1));
    }

    @Test
    public void insertFieldBeforeTextInParagraphWithoutUpdateField() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        insertFieldUsingFieldType(doc, FieldType.FIELD_AUTHOR, false, null, false, 1);

        msAssert.areEqual("\u0013 AUTHOR \u0014\u0015Hello World!\r", DocumentHelper.getParagraphText(doc, 1));
    }

    @Test
    public void insertFieldAfterTextInParagraphWithoutUpdateField() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        insertFieldUsingFieldType(doc, FieldType.FIELD_AUTHOR, false, null, true, 1);

        msAssert.areEqual("Hello World!\u0013 AUTHOR \u0014\u0015\r", DocumentHelper.getParagraphText(doc, 1));
    }

    @Test
    public void insertFieldWithoutSeparator() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        insertFieldUsingFieldType(doc, FieldType.FIELD_LIST_NUM, true, null, false, 1);

        msAssert.areEqual("\u0013 LISTNUM \u0015Hello World!\r", DocumentHelper.getParagraphText(doc, 1));
    }

    @Test
    public void insertFieldBeforeParagraphWithoutDocumentAuthor() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();
        doc.getBuiltInDocumentProperties().setAuthor("");

        insertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", null, null, false, 1);

        msAssert.areEqual("\u0013 AUTHOR \u0014\u0015Hello World!\r", DocumentHelper.getParagraphText(doc, 1));
    }

    @Test
    public void insertFieldAfterParagraphWithoutChangingDocumentAuthor() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        insertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", null, null, true, 1);

        msAssert.areEqual("Hello World!\u0013 AUTHOR \u0014\u0015\r", DocumentHelper.getParagraphText(doc, 1));
    }

    @Test
    public void insertFieldBeforeRunText() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        //Add some text into the paragraph
        Run run = DocumentHelper.insertNewRun(doc, " Hello World!", 1);

        insertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", "Test Field Value", run, false, 1);

        msAssert.areEqual("Hello World!\u0013 AUTHOR \u0014Test Field Value\u0015 Hello World!\r",
            DocumentHelper.getParagraphText(doc, 1));
    }

    @Test
    public void insertFieldAfterRunText() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        // Add some text into the paragraph
        Run run = DocumentHelper.insertNewRun(doc, " Hello World!", 1);

        insertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", "", run, true, 1);

        msAssert.areEqual("Hello World! Hello World!\u0013 AUTHOR \u0014\u0015\r",
            DocumentHelper.getParagraphText(doc, 1));
    }

    @Test (description = "WORDSNET-12396")
    public void insertFieldEmptyParagraphWithoutUpdateField() throws Exception
    {
        Document doc = DocumentHelper.createDocumentWithoutDummyText();

        insertFieldUsingFieldType(doc, FieldType.FIELD_AUTHOR, false, null, false, 1);

        msAssert.areEqual("\u0013 AUTHOR \u0014\u0015\f", DocumentHelper.getParagraphText(doc, 1));
    }

    @Test (description = "WORDSNET-12397")
    public void insertFieldEmptyParagraphWithUpdateField() throws Exception
    {
        Document doc = DocumentHelper.createDocumentWithoutDummyText();

        insertFieldUsingFieldType(doc, FieldType.FIELD_AUTHOR, true, null, false, 0);

        msAssert.areEqual("\u0013 AUTHOR \u0014Test Author\u0015\r", DocumentHelper.getParagraphText(doc, 0));
    }

    @Test
    public void getFormatRevision() throws Exception
    {
        //ExStart
        //ExFor:Paragraph.IsFormatRevision
        //ExSummary:Shows how to get information about whether this object was formatted in Microsoft Word while change tracking was enabled
        Document doc = new Document(getMyDir() + "Paragraph.IsFormatRevision.docx");

        Paragraph firstParagraph = DocumentHelper.getParagraph(doc, 0);
        Assert.assertTrue(firstParagraph.isFormatRevision());
        //ExEnd

        Paragraph secondParagraph = DocumentHelper.getParagraph(doc, 1);
        Assert.assertFalse(secondParagraph.isFormatRevision());
    }

    @Test
    public void getFrameProperties() throws Exception
    {
        //ExStart
        //ExFor:Paragraph.FrameFormat
        //ExFor:FrameFormat
        //ExFor:FrameFormat.IsFrame
        //ExFor:FrameFormat.Width
        //ExFor:FrameFormat.Height
        //ExFor:FrameFormat.HeightRule
        //ExFor:FrameFormat.HorizontalAlignment
        //ExFor:FrameFormat.VerticalAlignment
        //ExFor:FrameFormat.HorizontalPosition
        //ExFor:FrameFormat.RelativeHorizontalPosition
        //ExFor:FrameFormat.HorizontalDistanceFromText
        //ExFor:FrameFormat.VerticalPosition
        //ExFor:FrameFormat.RelativeVerticalPosition
        //ExFor:FrameFormat.VerticalDistanceFromText
        //ExSummary:Shows how to get information about formatting properties of paragraph as frame.
        Document doc = new Document(getMyDir() + "Paragraph.Frame.docx");

        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        for (Paragraph paragraph : paragraphs.<Paragraph>OfType() !!Autoporter error: Undefined expression type )
        {
            if (paragraph.getFrameFormat().isFrame())
            {
                msConsole.writeLine("Width: " + paragraph.getFrameFormat().getWidth());
                msConsole.writeLine("Height: " + paragraph.getFrameFormat().getHeight());
                msConsole.writeLine("HeightRule: " + paragraph.getFrameFormat().getHeightRule());
                msConsole.writeLine("HorizontalAlignment: " + paragraph.getFrameFormat().getHorizontalAlignment());
                msConsole.writeLine("VerticalAlignment: " + paragraph.getFrameFormat().getVerticalAlignment());
                msConsole.writeLine("HorizontalPosition: " + paragraph.getFrameFormat().getHorizontalPosition());
                msConsole.writeLine("RelativeHorizontalPosition: " +
                                  paragraph.getFrameFormat().getRelativeHorizontalPosition());
                msConsole.writeLine("HorizontalDistanceFromText: " +
                                  paragraph.getFrameFormat().getHorizontalDistanceFromText());
                msConsole.writeLine("VerticalPosition: " + paragraph.getFrameFormat().getVerticalPosition());
                msConsole.writeLine("RelativeVerticalPosition: " + paragraph.getFrameFormat().getRelativeVerticalPosition());
                msConsole.writeLine("VerticalDistanceFromText: " + paragraph.getFrameFormat().getVerticalDistanceFromText());
            }
        }
        //ExEnd

        if (paragraphs.get(0).getFrameFormat().isFrame())
        {
            msAssert.areEqual(233.3, paragraphs.get(0).getFrameFormat().getWidth());
            msAssert.areEqual(138.8, paragraphs.get(0).getFrameFormat().getHeight());
            msAssert.areEqual(21.05, paragraphs.get(0).getFrameFormat().getHorizontalPosition());
            msAssert.areEqual(RelativeHorizontalPosition.PAGE, paragraphs.get(0).getFrameFormat().getRelativeHorizontalPosition());
            msAssert.areEqual(9, paragraphs.get(0).getFrameFormat().getHorizontalDistanceFromText());
            msAssert.areEqual(-17.65, paragraphs.get(0).getFrameFormat().getVerticalPosition());
            msAssert.areEqual(RelativeVerticalPosition.PARAGRAPH, paragraphs.get(0).getFrameFormat().getRelativeVerticalPosition());
            msAssert.areEqual(0, paragraphs.get(0).getFrameFormat().getVerticalDistanceFromText());
        }
        else
        {
            Assert.fail("There are no frames in the document.");
        }
    }

    @Test
    public void asianTypographyProperties() throws Exception
    {
        //ExStart
        //ExFor:ParagraphFormat.FarEastLineBreakControl
        //ExFor:ParagraphFormat.WordWrap
        //ExFor:ParagraphFormat.HangingPunctuation
        //ExSummary:Shows how to set special properties for Asian typography. 
        Document doc = new Document(getMyDir() + "Document.docx");

        ParagraphFormat format = doc.getFirstSection().getBody().getParagraphs().get(0).getParagraphFormat();
        format.setFarEastLineBreakControl(true);
        format.setWordWrap(false);
        format.setHangingPunctuation(true);

        doc.save(getArtifactsDir() + "Paragraph.AsianTypographyProperties.docx");
        //ExEnd
    }

    /// <summary>
    /// Insert field into the first paragraph of the current document using field type
    /// </summary>
    private static void insertFieldUsingFieldType(Document doc, /*FieldType*/int fieldType, boolean updateField, Node refNode,
        boolean isAfter, int paraIndex) throws Exception
    {
        Paragraph para = DocumentHelper.getParagraph(doc, paraIndex);
        para.insertField(fieldType, updateField, refNode, isAfter);
    }

    /// <summary>
    /// Insert field into the first paragraph of the current document using field code
    /// </summary>
    private static void insertFieldUsingFieldCode(Document doc, String fieldCode, Node refNode, boolean isAfter,
        int paraIndex) throws Exception
    {
        Paragraph para = DocumentHelper.getParagraph(doc, paraIndex);
        para.insertField(fieldCode, refNode, isAfter);
    }

    /// <summary>
    /// Insert field into the first paragraph of the current document using field code and field String
    /// </summary>
    private static void insertFieldUsingFieldCodeFieldString(Document doc, String fieldCode, String fieldValue,
        Node refNode, boolean isAfter, int paraIndex)
    {
        Paragraph para = DocumentHelper.getParagraph(doc, paraIndex);
        para.insertField(fieldCode, fieldValue, refNode, isAfter);
    }

    @Test
    public void dropCapPosition() throws Exception
    {
        //ExStart
        //ExFor:DropCapPosition
        //ExSummary:Shows how to set the position of a drop cap.
        // Create a blank document
        Document doc = new Document();

        // Every paragraph has its own drop cap setting
        Paragraph para = doc.getFirstSection().getBody().getFirstParagraph();

        // By default, it is "none", for no drop caps
        msAssert.areEqual(com.aspose.words.DropCapPosition.NONE, para.getParagraphFormat().getDropCapPosition());

        // Move the first capital to outside the text margin
        para.getParagraphFormat().setDropCapPosition(com.aspose.words.DropCapPosition.MARGIN);

        // This text will be affected
        para.getRuns().add(new Run(doc, "Hello World!"));

        doc.save(getArtifactsDir() + "Paragraph.DropCap.docx");
        //ExEnd
    }
}
