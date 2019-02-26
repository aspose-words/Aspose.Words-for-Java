//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.Paragraph;
import com.aspose.words.FieldType;
import com.aspose.words.Run;
import org.testng.Assert;
import com.aspose.words.ParagraphCollection;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.Node;

import java.text.DateFormat;
import java.text.MessageFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.Date;
import java.util.Locale;

public class ExParagraph extends ApiExampleBase
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
        Run run = new Run(doc);
        {
            run.setText(" Hello World!");
        }
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

        Assert.assertEquals(DocumentHelper.getParagraphText(doc, 1), "\u0013 AUTHOR \u0014Test Author\u0015Hello World!\r");
    }

    @Test(enabled = false)
    public void insertFieldAfterTextInParagraph() throws Exception
    {
        LocalDateTime ldt = LocalDateTime.now();
        String date = DateTimeFormatter.ofPattern("MM/dd/yyyy", Locale.US).format(ldt);

        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        insertFieldUsingFieldCode(doc, " DATE ", null, true, 1);

        Assert.assertEquals(DocumentHelper.getParagraphText(doc, 1), MessageFormat.format("Hello World!\u0013 DATE \u0014{0}\u0015\r", date));
    }

    @Test
    public void insertFieldBeforeTextInParagraphWithoutUpdateField() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        insertFieldUsingFieldType(doc, FieldType.FIELD_AUTHOR, false, null, false, 1);

        Assert.assertEquals(DocumentHelper.getParagraphText(doc, 1), "\u0013 AUTHOR \u0014\u0015Hello World!\r");
    }

    @Test
    public void insertFieldAfterTextInParagraphWithoutUpdateField() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        insertFieldUsingFieldType(doc, FieldType.FIELD_AUTHOR, false, null, true, 1);

        Assert.assertEquals(DocumentHelper.getParagraphText(doc, 1), "Hello World!\u0013 AUTHOR \u0014\u0015\r");
    }

    @Test
    public void insertFieldWithoutSeparator() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        insertFieldUsingFieldType(doc, FieldType.FIELD_LIST_NUM, true, null, false, 1);

        Assert.assertEquals(DocumentHelper.getParagraphText(doc, 1), "\u0013 LISTNUM \u0015Hello World!\r");
    }

    @Test
    public void insertFieldBeforeParagraphWithoutDocumentAuthor() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();
        doc.getBuiltInDocumentProperties().setAuthor("");

        insertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", null, null, false, 1);

        Assert.assertEquals(DocumentHelper.getParagraphText(doc, 1), "\u0013 AUTHOR \u0014\u0015Hello World!\r");
    }

    @Test
    public void insertFieldAfterParagraphWithoutChangingDocumentAuthor() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        insertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", null, null, true, 1);

        Assert.assertEquals(DocumentHelper.getParagraphText(doc, 1), "Hello World!\u0013 AUTHOR \u0014\u0015\r");
    }

    @Test
    public void insertFieldBeforeRunText() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        //Add some text into the paragraph
        Run run = DocumentHelper.insertNewRun(doc, " Hello World!", 1);

        insertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", "Test Field Value", run, false, 1);

        Assert.assertEquals(DocumentHelper.getParagraphText(doc, 1), "Hello World!\u0013 AUTHOR \u0014Test Field Value\u0015 Hello World!\r");
    }

    @Test
    public void insertFieldAfterRunText() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        // Add some text into the paragraph
        Run run = DocumentHelper.insertNewRun(doc, " Hello World!", 1);

        insertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", "", run, true, 1);

        Assert.assertEquals(DocumentHelper.getParagraphText(doc, 1), "Hello World! Hello World!\u0013 AUTHOR \u0014\u0015\r");
    }

    @Test(description = "WORDSNET-12396")
    public void insertFieldEmptyParagraphWithoutUpdateField() throws Exception
    {
        Document doc = DocumentHelper.createDocumentWithoutDummyText();

        insertFieldUsingFieldType(doc, FieldType.FIELD_AUTHOR, false, null, false, 1);

        Assert.assertEquals(DocumentHelper.getParagraphText(doc, 1), "\u0013 AUTHOR \u0014\u0015\f");
    }

    @Test(description = "WORDSNET-12397")
    public void insertFieldEmptyParagraphWithUpdateField() throws Exception
    {
        Document doc = DocumentHelper.createDocumentWithoutDummyText();

        insertFieldUsingFieldType(doc, FieldType.FIELD_AUTHOR, true, null, false, 0);

        Assert.assertEquals(DocumentHelper.getParagraphText(doc, 0), "\u0013 AUTHOR \u0014Test Author\u0015\r");
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
        //ExFor:FrameFormat.IsFrame
        //ExFor:FrameFormat.Width
        //ExFor:FrameFormat.Height
        //ExFor:FrameFormat.HorizontalPosition
        //ExFor:FrameFormat.RelativeHorizontalPosition
        //ExFor:FrameFormat.HorizontalDistanceFromText
        //ExFor:FrameFormat.VerticalPosition
        //ExFor:FrameFormat.RelativeVerticalPosition
        //ExFor:FrameFormat.VerticalDistanceFromText
        //ExSummary:Shows how to get information about formatting properties of paragraph as frame.
        Document doc = new Document(getMyDir() + "Paragraph.Frame.docx");

        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        for (Paragraph paragraph : (Iterable<Paragraph>) paragraphs)
        {
            if (paragraph.getFrameFormat().isFrame())
            {
                System.out.println("Width: " + paragraph.getFrameFormat().getWidth());
                System.out.println("Height: " + paragraph.getFrameFormat().getHeight());
                System.out.println("HorizontalPosition: " + paragraph.getFrameFormat().getHorizontalPosition());
                System.out.println("RelativeHorizontalPosition: " + paragraph.getFrameFormat().getRelativeHorizontalPosition());
                System.out.println("HorizontalDistanceFromText: " + paragraph.getFrameFormat().getHorizontalDistanceFromText());
                System.out.println("VerticalPosition: " + paragraph.getFrameFormat().getVerticalPosition());
                System.out.println("RelativeVerticalPosition: " + paragraph.getFrameFormat().getRelativeVerticalPosition());
                System.out.println("VerticalDistanceFromText: " + paragraph.getFrameFormat().getVerticalDistanceFromText());
            }
        }
        //ExEnd

        if (paragraphs.get(0).getFrameFormat().isFrame())
        {
            Assert.assertEquals(paragraphs.get(0).getFrameFormat().getWidth(), 233.3);
            Assert.assertEquals(paragraphs.get(0).getFrameFormat().getHeight(), 138.8);
            Assert.assertEquals(paragraphs.get(0).getFrameFormat().getHorizontalPosition(), 21.05);
            Assert.assertEquals(paragraphs.get(0).getFrameFormat().getRelativeHorizontalPosition(), RelativeHorizontalPosition.PAGE);
            Assert.assertEquals(paragraphs.get(0).getFrameFormat().getHorizontalDistanceFromText(), 9.0);
            Assert.assertEquals(paragraphs.get(0).getFrameFormat().getVerticalPosition(), -17.65);
            Assert.assertEquals(paragraphs.get(0).getFrameFormat().getRelativeVerticalPosition(), RelativeVerticalPosition.PARAGRAPH);
            Assert.assertEquals(paragraphs.get(0).getFrameFormat().getVerticalDistanceFromText(), 0.0);
        }
        else
        {
            Assert.fail("There are no frames in the document.");
        }
    }

    /**
     *  Insert field into the first paragraph of the current document using field type
     */
    private static void insertFieldUsingFieldType(Document doc, /*FieldType*/int fieldType, boolean updateField, Node refNode, boolean isAfter, int paraIndex) throws Exception
    {
        Paragraph para = DocumentHelper.getParagraph(doc, paraIndex);
        para.insertField(fieldType, updateField, refNode, isAfter);
    }

    /**
     *  Insert field into the first paragraph of the current document using field code
     */
    private static void insertFieldUsingFieldCode(Document doc, String fieldCode, Node refNode, boolean isAfter, int paraIndex) throws Exception
    {
        Paragraph para = DocumentHelper.getParagraph(doc, paraIndex);
        para.insertField(fieldCode, refNode, isAfter);
    }

    /**
     *  Insert field into the first paragraph of the current document using field code and field String
     */
    private static void insertFieldUsingFieldCodeFieldString(Document doc, String fieldCode, String fieldValue, Node refNode, boolean isAfter, int paraIndex)
    {
        Paragraph para = DocumentHelper.getParagraph(doc, paraIndex);
        para.insertField(fieldCode, fieldValue, refNode, isAfter);
    }
}
