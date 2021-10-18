// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.words.Font;
import java.awt.Color;
import com.aspose.words.Underline;
import com.aspose.words.ParagraphFormat;
import com.aspose.words.ParagraphAlignment;
import org.testng.Assert;
import com.aspose.words.Paragraph;
import com.aspose.words.FieldType;
import java.util.Date;
import com.aspose.ms.System.DateTime;
import com.aspose.ms.System.TimeSpan;
import com.aspose.words.Run;
import com.aspose.words.Field;
import java.text.MessageFormat;
import com.aspose.words.NodeType;
import com.aspose.words.ParagraphCollection;
import com.aspose.words.HeightRule;
import com.aspose.words.HorizontalAlignment;
import com.aspose.words.VerticalAlignment;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.Node;
import com.aspose.words.Body;
import com.aspose.words.BreakType;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.TabStopCollection;
import com.aspose.words.TabAlignment;
import com.aspose.words.TabLeader;


@Test
class ExParagraph !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
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

        Font font = builder.getFont();
        font.setSize(16.0);
        font.setBold(true);
        font.setColor(Color.BLUE);
        font.setName("Arial");
        font.setUnderline(Underline.DASH);

        ParagraphFormat paragraphFormat = builder.getParagraphFormat();
        paragraphFormat.setFirstLineIndent(8.0);
        paragraphFormat.setAlignment(ParagraphAlignment.JUSTIFY);
        paragraphFormat.setAddSpaceBetweenFarEastAndAlpha(true);
        paragraphFormat.setAddSpaceBetweenFarEastAndDigit(true);
        paragraphFormat.setKeepTogether(true);

        // The "Writeln" method ends the paragraph after appending text
        // and then starts a new line, adding a new paragraph.
        builder.writeln("Hello world!");

        Assert.assertTrue(builder.getCurrentParagraph().isEndOfDocument());
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        Paragraph paragraph = doc.getFirstSection().getBody().getFirstParagraph();

        Assert.assertEquals(8, paragraph.getParagraphFormat().getFirstLineIndent());
        Assert.assertEquals(ParagraphAlignment.JUSTIFY, paragraph.getParagraphFormat().getAlignment());
        Assert.assertTrue(paragraph.getParagraphFormat().getAddSpaceBetweenFarEastAndAlpha());
        Assert.assertTrue(paragraph.getParagraphFormat().getAddSpaceBetweenFarEastAndDigit());
        Assert.assertTrue(paragraph.getParagraphFormat().getKeepTogether());
        Assert.assertEquals("Hello world!", paragraph.getText().trim());

        Font runFont = paragraph.getRuns().get(0).getFont();

        Assert.assertEquals(16.0d, runFont.getSize());
        Assert.assertTrue(runFont.getBold());
        Assert.assertEquals(Color.BLUE.getRGB(), runFont.getColor().getRGB());
        Assert.assertEquals("Arial", runFont.getName());
        Assert.assertEquals(Underline.DASH, runFont.getUnderline());
    }

    @Test
    public void appendField() throws Exception
    {
        //ExStart
        //ExFor:Paragraph.AppendField(FieldType, Boolean)
        //ExFor:Paragraph.AppendField(String)
        //ExFor:Paragraph.AppendField(String, String)
        //ExSummary:Shows various ways of appending fields to a paragraph.
        Document doc = new Document();
        Paragraph paragraph = doc.getFirstSection().getBody().getFirstParagraph();

        // Below are three ways of appending a field to the end of a paragraph.
        // 1 -  Append a DATE field using a field type, and then update it:
        paragraph.appendField(FieldType.FIELD_DATE, true);

        // 2 -  Append a TIME field using a field code: 
        paragraph.appendField(" TIME  \\@ \"HH:mm:ss\" ");

        // 3 -  Append a QUOTE field using a field code, and get it to display a placeholder value:
        paragraph.appendField(" QUOTE \"Real value\"", "Placeholder value");

        Assert.assertEquals("Placeholder value", doc.getRange().getFields().get(2).getResult());

        // This field will display its placeholder value until we update it.
        doc.updateFields();

        Assert.assertEquals("Real value", doc.getRange().getFields().get(2).getResult());

        doc.save(getArtifactsDir() + "Paragraph.AppendField.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Paragraph.AppendField.docx");

        TestUtil.verifyField(FieldType.FIELD_DATE, " DATE ", new Date(), doc.getRange().getFields().get(0), new TimeSpan(0, 0, 0, 0));
        TestUtil.verifyField(FieldType.FIELD_TIME, " TIME  \\@ \"HH:mm:ss\" ", new Date(), doc.getRange().getFields().get(1), new TimeSpan(0, 0, 0, 5));
        TestUtil.verifyField(FieldType.FIELD_QUOTE, " QUOTE \"Real value\"", "Real value", doc.getRange().getFields().get(2));
    }

    @Test
    public void insertField() throws Exception
    {
        //ExStart
        //ExFor:Paragraph.InsertField(string, Node, bool)
        //ExFor:Paragraph.InsertField(FieldType, bool, Node, bool)
        //ExFor:Paragraph.InsertField(string, string, Node, bool)
        //ExSummary:Shows various ways of adding fields to a paragraph.
        Document doc = new Document();
        Paragraph para = doc.getFirstSection().getBody().getFirstParagraph();

        // Below are three ways of inserting a field into a paragraph.
        // 1 -  Insert an AUTHOR field into a paragraph after one of the paragraph's child nodes:
        Run run = new Run(doc); { run.setText("This run was written by "); }
        para.appendChild(run);

        doc.getBuiltInDocumentProperties().get("Author").setValue("John Doe");
        para.insertField(FieldType.FIELD_AUTHOR, true, run, true);

        // 2 -  Insert a QUOTE field after one of the paragraph's child nodes:
        run = new Run(doc); { run.setText("."); }
        para.appendChild(run);

        Field field = para.insertField(" QUOTE \" Real value\" ", run, true);

        // 3 -  Insert a QUOTE field before one of the paragraph's child nodes,
        // and get it to display a placeholder value:
        para.insertField(" QUOTE \" Real value.\"", " Placeholder value.", field.getStart(), false);

        Assert.assertEquals(" Placeholder value.", doc.getRange().getFields().get(1).getResult());

        // This field will display its placeholder value until we update it.
        doc.updateFields();

        Assert.assertEquals(" Real value.", doc.getRange().getFields().get(1).getResult());

        doc.save(getArtifactsDir() + "Paragraph.InsertField.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Paragraph.InsertField.docx");

        TestUtil.verifyField(FieldType.FIELD_AUTHOR, " AUTHOR ", "John Doe", doc.getRange().getFields().get(0));
        TestUtil.verifyField(FieldType.FIELD_QUOTE, " QUOTE \" Real value.\"", " Real value.", doc.getRange().getFields().get(1));
        TestUtil.verifyField(FieldType.FIELD_QUOTE, " QUOTE \" Real value\" ", " Real value", doc.getRange().getFields().get(2));
    }

    @Test
    public void insertFieldBeforeTextInParagraph() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        insertFieldUsingFieldCode(doc, " AUTHOR ", null, false, 1);

        Assert.assertEquals("\u0013 AUTHOR \u0014Test Author\u0015Hello World!\r",
            DocumentHelper.getParagraphText(doc, 1));
    }

    @Test
    public void insertFieldAfterTextInParagraph() throws Exception
    {
        String date = DateTime.getToday().toString("d");

        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        insertFieldUsingFieldCode(doc, " DATE ", null, true, 1);

        Assert.assertEquals(MessageFormat.format("Hello World!\u0013 DATE \u0014{0}\u0015\r", date),
            DocumentHelper.getParagraphText(doc, 1));
    }

    @Test
    public void insertFieldBeforeTextInParagraphWithoutUpdateField() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        insertFieldUsingFieldType(doc, FieldType.FIELD_AUTHOR, false, null, false, 1);

        Assert.assertEquals("\u0013 AUTHOR \u0014\u0015Hello World!\r", DocumentHelper.getParagraphText(doc, 1));
    }

    @Test
    public void insertFieldAfterTextInParagraphWithoutUpdateField() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        insertFieldUsingFieldType(doc, FieldType.FIELD_AUTHOR, false, null, true, 1);

        Assert.assertEquals("Hello World!\u0013 AUTHOR \u0014\u0015\r", DocumentHelper.getParagraphText(doc, 1));
    }

    @Test
    public void insertFieldWithoutSeparator() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        insertFieldUsingFieldType(doc, FieldType.FIELD_LIST_NUM, true, null, false, 1);

        Assert.assertEquals("\u0013 LISTNUM \u0015Hello World!\r", DocumentHelper.getParagraphText(doc, 1));
    }

    @Test
    public void insertFieldBeforeParagraphWithoutDocumentAuthor() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();
        doc.getBuiltInDocumentProperties().setAuthor("");

        insertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", null, null, false, 1);

        Assert.assertEquals("\u0013 AUTHOR \u0014\u0015Hello World!\r", DocumentHelper.getParagraphText(doc, 1));
    }

    @Test
    public void insertFieldAfterParagraphWithoutChangingDocumentAuthor() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        insertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", null, null, true, 1);

        Assert.assertEquals("Hello World!\u0013 AUTHOR \u0014\u0015\r", DocumentHelper.getParagraphText(doc, 1));
    }

    @Test
    public void insertFieldBeforeRunText() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        //Add some text into the paragraph
        Run run = DocumentHelper.insertNewRun(doc, " Hello World!", 1);

        insertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", "Test Field Value", run, false, 1);

        Assert.assertEquals("Hello World!\u0013 AUTHOR \u0014Test Field Value\u0015 Hello World!\r",
            DocumentHelper.getParagraphText(doc, 1));
    }

    @Test
    public void insertFieldAfterRunText() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        // Add some text into the paragraph
        Run run = DocumentHelper.insertNewRun(doc, " Hello World!", 1);

        insertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", "", run, true, 1);

        Assert.assertEquals("Hello World! Hello World!\u0013 AUTHOR \u0014\u0015\r",
            DocumentHelper.getParagraphText(doc, 1));
    }

    @Test (description = "WORDSNET-12396")
    public void insertFieldEmptyParagraphWithoutUpdateField() throws Exception
    {
        Document doc = DocumentHelper.createDocumentWithoutDummyText();

        insertFieldUsingFieldType(doc, FieldType.FIELD_AUTHOR, false, null, false, 1);

        Assert.assertEquals("\u0013 AUTHOR \u0014\u0015\f", DocumentHelper.getParagraphText(doc, 1));
    }

    @Test (description = "WORDSNET-12397")
    public void insertFieldEmptyParagraphWithUpdateField() throws Exception
    {
        Document doc = DocumentHelper.createDocumentWithoutDummyText();

        insertFieldUsingFieldType(doc, FieldType.FIELD_AUTHOR, true, null, false, 0);

        Assert.assertEquals("\u0013 AUTHOR \u0014Test Author\u0015\r", DocumentHelper.getParagraphText(doc, 0));
    }

    @Test
    public void compositeNodeChildren() throws Exception
    {
        //ExStart
        //ExFor:CompositeNode.Count
        //ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
        //ExFor:CompositeNode.InsertAfter(Node, Node)
        //ExFor:CompositeNode.InsertBefore(Node, Node)
        //ExFor:CompositeNode.PrependChild(Node) 
        //ExFor:Paragraph.GetText
        //ExFor:Run
        //ExSummary:Shows how to add, update and delete child nodes in a CompositeNode's collection of children.
        Document doc = new Document();

        // An empty document, by default, has one paragraph.
        Assert.assertEquals(1, doc.getFirstSection().getBody().getParagraphs().getCount());

        // Composite nodes such as our paragraph can contain other composite and inline nodes as children.
        Paragraph paragraph = doc.getFirstSection().getBody().getFirstParagraph();
        Run paragraphText = new Run(doc, "Initial text. ");
        paragraph.appendChild(paragraphText);

        // Create three more run nodes.
        Run run1 = new Run(doc, "Run 1. ");
        Run run2 = new Run(doc, "Run 2. ");
        Run run3 = new Run(doc, "Run 3. ");

        // The document body will not display these runs until we insert them into a composite node
        // that itself is a part of the document's node tree, as we did with the first run.
        // We can determine where the text contents of nodes that we insert
        // appears in the document by specifying an insertion location relative to another node in the paragraph.
        Assert.assertEquals("Initial text.", paragraph.getText().trim());

        // Insert the second run into the paragraph in front of the initial run.
        paragraph.insertBefore(run2, paragraphText);

        Assert.assertEquals("Run 2. Initial text.", paragraph.getText().trim());

        // Insert the third run after the initial run.
        paragraph.insertAfter(run3, paragraphText);

        Assert.assertEquals("Run 2. Initial text. Run 3.", paragraph.getText().trim());

        // Insert the first run to the start of the paragraph's child nodes collection.
        paragraph.prependChild(run1);

        Assert.assertEquals("Run 1. Run 2. Initial text. Run 3.", paragraph.getText().trim());
        Assert.assertEquals(4, paragraph.getChildNodes(NodeType.ANY, true).getCount());

        // We can modify the contents of the run by editing and deleting existing child nodes.
        ((Run)paragraph.getChildNodes(NodeType.RUN, true).get(1)).setText("Updated run 2. ");
        paragraph.getChildNodes(NodeType.RUN, true).remove(paragraphText);

        Assert.assertEquals("Run 1. Updated run 2. Run 3.", paragraph.getText().trim());
        Assert.assertEquals(3, paragraph.getChildNodes(NodeType.ANY, true).getCount());
        //ExEnd
    }

    @Test
    public void revisions() throws Exception
    {
        //ExStart
        //ExFor:Paragraph.IsMoveFromRevision
        //ExFor:Paragraph.IsMoveToRevision
        //ExFor:ParagraphCollection
        //ExFor:ParagraphCollection.Item(Int32)
        //ExFor:Story.Paragraphs
        //ExSummary:Shows how to check whether a paragraph is a move revision.
        Document doc = new Document(getMyDir() + "Revisions.docx");

        // This document contains "Move" revisions, which appear when we highlight text with the cursor,
        // and then drag it to move it to another location
        // while tracking revisions in Microsoft Word via "Review" -> "Track changes".
        Assert.AreEqual(6, doc.getRevisions().Count(r => r.RevisionType == RevisionType.Moving));

        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        // Move revisions consist of pairs of "Move from", and "Move to" revisions. 
        // These revisions are potential changes to the document that we can either accept or reject.
        // Before we accept/reject a move revision, the document
        // must keep track of both the departure and arrival destinations of the text.
        // The second and the fourth paragraph define one such revision, and thus both have the same contents.
        Assert.assertEquals(paragraphs.get(1).getText(), paragraphs.get(3).getText());

        // The "Move from" revision is the paragraph where we dragged the text from.
        // If we accept the revision, this paragraph will disappear,
        // and the other will remain and no longer be a revision.
        Assert.assertTrue(paragraphs.get(1).isMoveFromRevision());

        // The "Move to" revision is the paragraph where we dragged the text to.
        // If we reject the revision, this paragraph instead will disappear, and the other will remain.
        Assert.assertTrue(paragraphs.get(3).isMoveToRevision());
        //ExEnd
    }

    @Test
    public void getFormatRevision() throws Exception
    {
        //ExStart
        //ExFor:Paragraph.IsFormatRevision
        //ExSummary:Shows how to check whether a paragraph is a format revision.
        Document doc = new Document(getMyDir() + "Format revision.docx");

        // This paragraph is a "Format" revision, which occurs when we change the formatting of existing text
        // while tracking revisions in Microsoft Word via "Review" -> "Track changes".
        Assert.assertTrue(doc.getFirstSection().getBody().getFirstParagraph().isFormatRevision());
        //ExEnd
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
        //ExSummary:Shows how to get information about formatting properties of paragraphs that are frames.
        Document doc = new Document(getMyDir() + "Paragraph frame.docx");

        Paragraph paragraphFrame = doc.getFirstSection().getBody().getParagraphs().<Paragraph>OfType().First(p => p.FrameFormat.IsFrame);

        Assert.assertEquals(233.3d, paragraphFrame.getFrameFormat().getWidth());
        Assert.assertEquals(138.8d, paragraphFrame.getFrameFormat().getHeight());
        Assert.assertEquals(HeightRule.AT_LEAST, paragraphFrame.getFrameFormat().getHeightRule());
        Assert.assertEquals(HorizontalAlignment.DEFAULT, paragraphFrame.getFrameFormat().getHorizontalAlignment());
        Assert.assertEquals(VerticalAlignment.DEFAULT, paragraphFrame.getFrameFormat().getVerticalAlignment());
        Assert.assertEquals(34.05d, paragraphFrame.getFrameFormat().getHorizontalPosition());
        Assert.assertEquals(RelativeHorizontalPosition.PAGE, paragraphFrame.getFrameFormat().getRelativeHorizontalPosition());
        Assert.assertEquals(9.0d, paragraphFrame.getFrameFormat().getHorizontalDistanceFromText());
        Assert.assertEquals(20.5d, paragraphFrame.getFrameFormat().getVerticalPosition());
        Assert.assertEquals(RelativeVerticalPosition.PARAGRAPH, paragraphFrame.getFrameFormat().getRelativeVerticalPosition());
        Assert.assertEquals(0.0d, paragraphFrame.getFrameFormat().getVerticalDistanceFromText());
        //ExEnd
    }

    /// <summary>
    /// Insert field into the first paragraph of the current document using field type.
    /// </summary>
    private static void insertFieldUsingFieldType(Document doc, /*FieldType*/int fieldType, boolean updateField, Node refNode,
        boolean isAfter, int paraIndex) throws Exception
    {
        Paragraph para = DocumentHelper.getParagraph(doc, paraIndex);
        para.insertField(fieldType, updateField, refNode, isAfter);
    }

    /// <summary>
    /// Insert field into the first paragraph of the current document using field code.
    /// </summary>
    private static void insertFieldUsingFieldCode(Document doc, String fieldCode, Node refNode, boolean isAfter,
        int paraIndex) throws Exception
    {
        Paragraph para = DocumentHelper.getParagraph(doc, paraIndex);
        para.insertField(fieldCode, refNode, isAfter);
    }

    /// <summary>
    /// Insert field into the first paragraph of the current document using field code and field String.
    /// </summary>
    private static void insertFieldUsingFieldCodeFieldString(Document doc, String fieldCode, String fieldValue,
        Node refNode, boolean isAfter, int paraIndex)
    {
        Paragraph para = DocumentHelper.getParagraph(doc, paraIndex);
        para.insertField(fieldCode, fieldValue, refNode, isAfter);
    }

    @Test
    public void isRevision() throws Exception
    {
        //ExStart
        //ExFor:Paragraph.IsDeleteRevision
        //ExFor:Paragraph.IsInsertRevision
        //ExSummary:Shows how to work with revision paragraphs.
        Document doc = new Document();
        Body body = doc.getFirstSection().getBody();
        Paragraph para = body.getFirstParagraph();

        para.appendChild(new Run(doc, "Paragraph 1. "));
        body.appendParagraph("Paragraph 2. ");
        body.appendParagraph("Paragraph 3. ");

        // The above paragraphs are not revisions.
        // Paragraphs that we add after starting revision tracking will register as "Insert" revisions.
        doc.startTrackRevisionsInternal("John Doe", new Date());

        para = body.appendParagraph("Paragraph 4. ");

        Assert.assertTrue(para.isInsertRevision());

        // Paragraphs that we remove after starting revision tracking will register as "Delete" revisions.
        ParagraphCollection paragraphs = body.getParagraphs();

        Assert.assertEquals(4, paragraphs.getCount());

        para = paragraphs.get(2);
        para.remove();

        // Such paragraphs will remain until we either accept or reject the delete revision.
        // Accepting the revision will remove the paragraph for good,
        // and rejecting the revision will leave it in the document as if we never deleted it.
        Assert.assertEquals(4, paragraphs.getCount());
        Assert.assertTrue(para.isDeleteRevision());

        // Accept the revision, and then verify that the paragraph is gone.
        doc.acceptAllRevisions();

        Assert.assertEquals(3, paragraphs.getCount());
        Assert.That(para, Is.Empty);
        Assert.assertEquals(
            "Paragraph 1. \r" +
            "Paragraph 2. \r" +
            "Paragraph 4.", doc.getText().trim());
        //ExEnd
    }

    @Test
    public void breakIsStyleSeparator() throws Exception
    {
        //ExStart
        //ExFor:Paragraph.BreakIsStyleSeparator
        //ExSummary:Shows how to write text to the same line as a TOC heading and have it not show up in the TOC.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertTableOfContents("\\o \\h \\z \\u");
        builder.insertBreak(BreakType.PAGE_BREAK);

        // Insert a paragraph with a style that the TOC will pick up as an entry.
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);

        // Both these strings are in the same paragraph and will therefore show up on the same TOC entry.
        builder.write("Heading 1. ");
        builder.write("Will appear in the TOC. ");

        // If we insert a style separator, we can write more text in the same paragraph
        // and use a different style without showing up in the TOC.
        // If we use a heading type style after the separator, we can draw multiple TOC entries from one document text line.
        builder.insertStyleSeparator();
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.QUOTE);
        builder.write("Won't appear in the TOC. ");

        Assert.assertTrue(doc.getFirstSection().getBody().getFirstParagraph().getBreakIsStyleSeparator());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Paragraph.BreakIsStyleSeparator.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Paragraph.BreakIsStyleSeparator.docx");

        TestUtil.verifyField(FieldType.FIELD_TOC, "TOC \\o \\h \\z \\u", 
            "\u0013 HYPERLINK \\l \"_Toc256000000\" \u0014Heading 1. Will appear in the TOC.\t\u0013 PAGEREF _Toc256000000 \\h \u00142\u0015\u0015\r", doc.getRange().getFields().get(0));
        Assert.assertFalse(doc.getFirstSection().getBody().getFirstParagraph().getBreakIsStyleSeparator());
    }

    @Test
    public void tabStops() throws Exception
    {
        //ExStart
        //ExFor:Paragraph.GetEffectiveTabStops
        //ExSummary:Shows how to set custom tab stops for a paragraph.
        Document doc = new Document();
        Paragraph para = doc.getFirstSection().getBody().getFirstParagraph();

        // If we are in a paragraph with no tab stops in this collection,
        // the cursor will jump 36 points each time we press the Tab key in Microsoft Word.
        Assert.assertEquals(0, doc.getFirstSection().getBody().getFirstParagraph().getEffectiveTabStops().length);

        // We can add custom tab stops in Microsoft Word if we enable the ruler via the "View" tab.
        // Each unit on this ruler is two default tab stops, which is 72 points.
        // We can add custom tab stops programmatically like this.
        TabStopCollection tabStops = doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getTabStops();
        tabStops.add(72.0, TabAlignment.LEFT, TabLeader.DOTS);
        tabStops.add(216.0, TabAlignment.CENTER, TabLeader.DASHES);
        tabStops.add(360.0, TabAlignment.RIGHT, TabLeader.LINE);

        // We can see these tab stops in Microsoft Word by enabling the ruler via "View" -> "Show" -> "Ruler".
        Assert.assertEquals(3, para.getEffectiveTabStops().length);

        // Any tab characters we add will make use of the tab stops on the ruler and may,
        // depending on the tab leader's value, leave a line between the tab departure and arrival destinations.
        para.appendChild(new Run(doc, "\tTab 1\tTab 2\tTab 3"));

        doc.save(getArtifactsDir() + "Paragraph.TabStops.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Paragraph.TabStops.docx");
        tabStops = doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getTabStops();

        TestUtil.verifyTabStop(72.0d, TabAlignment.LEFT, TabLeader.DOTS, false, tabStops.get(0));
        TestUtil.verifyTabStop(216.0d, TabAlignment.CENTER, TabLeader.DASHES, false, tabStops.get(1));
        TestUtil.verifyTabStop(360.0d, TabAlignment.RIGHT, TabLeader.LINE, false, tabStops.get(2));
    }

    @Test
    public void joinRuns() throws Exception
    {
        //ExStart
        //ExFor:Paragraph.JoinRunsWithSameFormatting
        //ExSummary:Shows how to simplify paragraphs by merging superfluous runs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert four runs of text into the paragraph.
        builder.write("Run 1. ");
        builder.write("Run 2. ");
        builder.write("Run 3. ");
        builder.write("Run 4. ");

        // If we open this document in Microsoft Word, the paragraph will look like one seamless text body.
        // However, it will consist of four separate runs with the same formatting. Fragmented paragraphs like this
        // may occur when we manually edit parts of one paragraph many times in Microsoft Word.
        Paragraph para = builder.getCurrentParagraph();

        Assert.assertEquals(4, para.getRuns().getCount());

        // Change the style of the last run to set it apart from the first three.
        para.getRuns().get(3).getFont().setStyleIdentifier(StyleIdentifier.EMPHASIS);

        // We can run the "JoinRunsWithSameFormatting" method to optimize the document's contents
        // by merging similar runs into one, reducing their overall count.
        // This method also returns the number of runs that this method merged.
        // These two merges occurred to combine Runs #1, #2, and #3,
        // while leaving out Run #4 because it has an incompatible style.
        Assert.assertEquals(2, para.joinRunsWithSameFormatting());

        // The number of runs left will equal the original count
        // minus the number of run merges that the "JoinRunsWithSameFormatting" method carried out.
        Assert.assertEquals(2, para.getRuns().getCount());
        Assert.assertEquals("Run 1. Run 2. Run 3. ", para.getRuns().get(0).getText());
        Assert.assertEquals("Run 4. ", para.getRuns().get(1).getText());
        //ExEnd
    }
}
