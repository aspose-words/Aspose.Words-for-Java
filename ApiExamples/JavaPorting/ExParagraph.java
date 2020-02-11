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
import com.aspose.words.Body;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.BreakType;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.TabAlignment;
import com.aspose.words.TabLeader;
import com.aspose.words.LineSpacingRule;


@Test
class ExParagraph !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void insertField() throws Exception
    {
        //ExStart
        //ExFor:Paragraph.AppendField(FieldType, Boolean)
        //ExFor:Paragraph.AppendField(String)
        //ExFor:Paragraph.AppendField(String, String)
        //ExFor:Paragraph.InsertField(string, Node, bool)
        //ExFor:Paragraph.InsertField(FieldType, bool, Node, bool)
        //ExFor:Paragraph.InsertField(string, string, Node, bool)
        //ExSummary:Shows how to insert fields in different ways.
        // Create a blank document and get its first paragraph
        Document doc = new Document();
        Paragraph para = doc.getFirstSection().getBody().getFirstParagraph();

        // Choose a field by FieldType, append it to the end of the paragraph and update it
        para.appendField(FieldType.FIELD_DATE, true);

        // Append a field with a field code created by hand 
        para.appendField(" TIME  \\@ \"HH:mm:ss\" ");

        // Append a field that will display a placeholder value until it is updated manually in Microsoft Word
        // or programmatically with Document.UpdateFields() or Field.Update()
        para.appendField(" QUOTE \"Real value\"", "Placeholder value");

        // We can choose a node in the paragraph and insert a field
        // before or after that node instead of appending it to the end of a paragraph
        para = doc.getFirstSection().getBody().appendParagraph("");
        Run run = new Run(doc); { run.setText(" My Run. "); }
        para.appendChild(run);

        // Insert a field into the paragraph and place it before the run we created
        doc.getBuiltInDocumentProperties().get("Author").setValue("John Doe");
        para.insertField(FieldType.FIELD_AUTHOR, true, run, false);

        // Insert another field designated by field code before the run
        para.insertField(" QUOTE \"Real value\" ", run, false);

        // Insert another field with a place holder value and place it after the run
        para.insertField(" QUOTE \"Real value\"", " Placeholder value. ", run, true);

        doc.save(getArtifactsDir() + "Paragraph.InsertField.docx");
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
        Document doc = new Document(getMyDir() + "Format revision.docx");

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
        Document doc = new Document(getMyDir() + "Paragraph frame.docx");

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
            msAssert.areEqual(34.05, paragraphs.get(0).getFrameFormat().getHorizontalPosition());
            msAssert.areEqual(RelativeHorizontalPosition.PAGE, paragraphs.get(0).getFrameFormat().getRelativeHorizontalPosition());
            msAssert.areEqual(9, paragraphs.get(0).getFrameFormat().getHorizontalDistanceFromText());
            msAssert.areEqual(20.5, paragraphs.get(0).getFrameFormat().getVerticalPosition());
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
        para.getParagraphFormat().setLinesToDrop(2);

        // This text will be affected
        para.getRuns().add(new Run(doc, "Hello World!"));

        doc.save(getArtifactsDir() + "Paragraph.DropCapPosition.docx");
        //ExEnd
    }

    @Test
    public void isRevision() throws Exception
    {
        //ExStart
        //ExFor:Paragraph.IsDeleteRevision
        //ExFor:Paragraph.IsInsertRevision
        //ExSummary:Shows how to work with revision paragraphs.
        // Create a blank document, populate the first paragraph with text and add two more
        Document doc = new Document();
        Body body = doc.getFirstSection().getBody();
        Paragraph para = body.getFirstParagraph();
        para.appendChild(new Run(doc, "Paragraph 1. "));
        body.appendParagraph("Paragraph 2. ");
        body.appendParagraph("Paragraph 3. ");

        // We have three paragraphs, none of which registered as any type of revision
        // If we add/remove any content in the document while tracking revisions,
        // they will be displayed as such in the document and can be accepted/rejected
        doc.startTrackRevisionsInternal("John Doe", DateTime.getNow());

        // This paragraph is a revision and will have the according "IsInsertRevision" flag set
        para = body.appendParagraph("Paragraph 4. ");
        Assert.assertTrue(para.isInsertRevision());

        // Get the document's paragraph collection and remove a paragraph
        ParagraphCollection paragraphs = body.getParagraphs();
        msAssert.areEqual(4, paragraphs.getCount());
        para = paragraphs.get(2);
        para.remove();

        // Since we are tracking revisions, the paragraph still exists in the document, will have the "IsDeleteRevision" set
        // and will be displayed as a revision in Microsoft Word, until we accept or reject all revisions
        msAssert.areEqual(4, paragraphs.getCount());
        Assert.assertTrue(para.isDeleteRevision());

        // The delete revision paragraph is removed once we accept changes
        doc.acceptAllRevisions();
        msAssert.areEqual(3, paragraphs.getCount());
        Assert.That(para, Is.Empty);
        //ExEnd
    }

    @Test
    public void breakIsStyleSeparator() throws Exception
    {
        //ExStart
        //ExFor:Paragraph.BreakIsStyleSeparator
        //ExSummary:Shows how to write text to the same line as a TOC heading and have it not show up in the TOC.
        // Create a blank document and insert a table of contents field
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertTableOfContents("\\o \\h \\z \\u");
        builder.insertBreak(BreakType.PAGE_BREAK);

        // Insert a paragraph with a style that will be picked up as an entry in the TOC
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);

        // Both these strings are on the same line and same paragraph and will therefore show up on the same TOC entry
        builder.write("Heading 1. ");
        builder.write("Will appear in the TOC. ");

        // Any text on a new line that does not have a heading style will not register as a TOC entry
        // If we insert a style separator, we can write more text on the same line
        // and use a different style without it showing up in the TOC
        // If we use a heading type style afterwards, we can draw two TOC entries from one line of document text
        builder.insertStyleSeparator();
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.QUOTE);
        builder.write("Won't appear in the TOC. ");

        // This flag is set to true for such paragraphs
        Assert.assertTrue(doc.getFirstSection().getBody().getParagraphs().get(0).getBreakIsStyleSeparator());

        // Update the TOC and save the document
        doc.updateFields();
        doc.save(getArtifactsDir() + "Paragraph.BreakIsStyleSeparator.docx");
        //ExEnd
    }

    @Test
    public void tabStops() throws Exception
    {
        //ExStart
        //ExFor:Paragraph.GetEffectiveTabStops
        //ExSummary:Shows how to set custom tab stops.
        // Create a blank document and get the first paragraph
        Document doc = new Document();
        Paragraph para = doc.getFirstSection().getBody().getFirstParagraph();

        // If there are no tab stops in this collection, while we are in this paragraph
        // the cursor will jump 36 points each time we press the Tab key in Microsoft Word
        msAssert.areEqual(0, para.getEffectiveTabStops().length);

        // We can add custom tab stops in Microsoft Word if we enable the ruler via the view tab
        // Each unit on that ruler is two default tab stops, which is 72 points
        // Those tab stops can be programmatically added to the paragraph like this
        para.getParagraphFormat().getTabStops().add(72.0, TabAlignment.LEFT, TabLeader.DOTS);
        para.getParagraphFormat().getTabStops().add(216.0, TabAlignment.CENTER, TabLeader.DASHES);
        para.getParagraphFormat().getTabStops().add(360.0, TabAlignment.RIGHT, TabLeader.LINE);

        // These tab stops are added to this collection, and can also be seen by enabling the ruler mentioned above
        msAssert.areEqual(3, para.getEffectiveTabStops().length);

        // Add a Run with tab characters that will snap the text to our TabStop positions and save the document
        para.appendChild(new Run(doc, "\tTab 1\tTab 2\tTab 3"));
        doc.save(getArtifactsDir() + "Paragraph.TabStops.docx");
        //ExEnd
    }

    @Test
    public void joinRuns() throws Exception
    {
        //ExStart
        //ExFor:Paragraph.JoinRunsWithSameFormatting
        //ExSummary:Shows how to simplify paragraphs by merging superfluous runs.
        // Create a blank Document and insert a few short Runs into the first Paragraph
        // Having many small runs with the same formatting can happen if, for instance,
        // we edit a document extensively in Microsoft Word
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Run 1. ");
        builder.write("Run 2. ");
        builder.write("Run 3. ");
        builder.write("Run 4. ");

        // The Paragraph may look like it's in once piece in Microsoft Word,
        // but under the surface it is fragmented into several Runs, which leaves room for optimization
        Paragraph para = builder.getCurrentParagraph();
        msAssert.areEqual(4, para.getRuns().getCount());

        // Change the style of the last run to something different from the first three
        para.getRuns().get(3).getFont().setStyleIdentifier(StyleIdentifier.EMPHASIS);

        // We can run the JoinRunsWithSameFormatting() method to merge similar Runs
        // This method also returns the number of joins that occured during the merge
        // Two merges occured to combine Runs 1-3, while Run 4 was left out because it has an incompatible style
        msAssert.areEqual(2, para.joinRunsWithSameFormatting());

        // The paragraph has been simplified to two runs
        msAssert.areEqual(2, para.getRuns().getCount());
        msAssert.areEqual("Run 1. Run 2. Run 3. ", para.getRuns().get(0).getText());
        msAssert.areEqual("Run 4. ", para.getRuns().get(1).getText());
        //ExEnd
    }

    @Test
    public void lineSpacing() throws Exception
    {
        //ExStart
        //ExFor:ParagraphFormat.LineSpacing
        //ExFor:ParagraphFormat.LineSpacingRule
        //ExSummary:Shows how to work with line spacing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph's line spacing to have a minimum value
        // This will give vertical padding to lines of text of any size that's too small to maintain the line height
        builder.getParagraphFormat().setLineSpacingRule(LineSpacingRule.AT_LEAST);
        builder.getParagraphFormat().setLineSpacing(20.0);

        builder.writeln("Minimum line spacing of 20.");
        builder.writeln("Minimum line spacing of 20.");

        // Set the line spacing to always be exactly 5 points
        // If the font size is larger than the spacing, the top of the text will be truncated
        builder.insertParagraph();
        builder.getParagraphFormat().setLineSpacingRule(LineSpacingRule.EXACTLY);
        builder.getParagraphFormat().setLineSpacing(5.0);

        builder.writeln("Line spacing of exactly 5.");
        builder.writeln("Line spacing of exactly 5.");

        // Set the line spacing to a multiple of the default line spacing, which is 12 points by default
        // 18 points will set the spacing to always be 1.5 lines, which will scale with different font sizes
        builder.insertParagraph();
        builder.getParagraphFormat().setLineSpacingRule(LineSpacingRule.MULTIPLE);
        builder.getParagraphFormat().setLineSpacing(18.0);

        builder.writeln("Line spacing of 1.5 default lines.");
        builder.writeln("Line spacing of 1.5 default lines.");

        doc.save(getArtifactsDir() + "Paragraph.LineSpacing.docx");
        //ExEnd
    }

    @Test
    public void paragraphSpacing() throws Exception
    {
        //ExStart
        //ExFor:ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle
        //ExFor:ParagraphFormat.SpaceAfter
        //ExFor:ParagraphFormat.SpaceAfterAuto
        //ExFor:ParagraphFormat.SpaceBefore
        //ExFor:ParagraphFormat.SpaceBeforeAuto
        //ExSummary:Shows how to work with paragraph spacing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the amount of white space before and after each paragraph to 12 points
        builder.getParagraphFormat().setSpaceBefore(12.0f);
        builder.getParagraphFormat().setSpaceAfter(12.0f);

        // We can set these flags to apply default spacing, effectively ignoring the spacing in the attributes we set above
        Assert.assertFalse(builder.getParagraphFormat().getSpaceAfterAuto());
        Assert.assertFalse(builder.getParagraphFormat().getSpaceBeforeAuto());
        Assert.assertFalse(builder.getParagraphFormat().getNoSpaceBetweenParagraphsOfSameStyle());

        // Insert two paragraphs which will have padding above and below them and save the document
        builder.writeln("Paragraph 1.");
        builder.writeln("Paragraph 2.");

        doc.save(getArtifactsDir() + "Paragraph.ParagraphSpacing.docx");
        //ExEnd
    }

    @Test
    public void outlineLevel() throws Exception
    {
        //ExStart
        //ExFor:ParagraphFormat.OutlineLevel
        //ExSummary:Shows how to set paragraph outline levels to create collapsible text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Each paragraph has an OutlineLevel, which could be any number from 1 to 9, or at the default "BodyText" value
        // Setting the attribute to one of the numbered values will enable an arrow in Microsoft Word
        // next to the beginning of the paragraph that, when clicked, will collapse the paragraph
        builder.getParagraphFormat().setOutlineLevel(com.aspose.words.OutlineLevel.LEVEL_1);
        builder.writeln("Paragraph outline level 1.");

        // Level 1 is the topmost level, which practically means that clicking its arrow will also collapse
        // any following paragraph with a lower level, like the paragraphs below
        builder.getParagraphFormat().setOutlineLevel(com.aspose.words.OutlineLevel.LEVEL_2);
        builder.writeln("Paragraph outline level 2.");

        // Two paragraphs of the same level will not collapse each other
        builder.getParagraphFormat().setOutlineLevel(com.aspose.words.OutlineLevel.LEVEL_3);
        builder.writeln("Paragraph outline level 3.");
        builder.writeln("Paragraph outline level 3.");

        // The default "BodyText" value is the lowest
        builder.getParagraphFormat().setOutlineLevel(com.aspose.words.OutlineLevel.BODY_TEXT);
        builder.writeln("Paragraph at main text level.");

        doc.save(getArtifactsDir() + "Paragraph.OutlineLevel.docx");
        //ExEnd
    }

    @Test
    public void pageBreakBefore() throws Exception
    {
        //ExStart
        //ExFor:ParagraphFormat.PageBreakBefore
        //ExSummary:Shows how to force a page break before each paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set this to insert a page break before this paragraph
        builder.getParagraphFormat().setPageBreakBefore(true);

        // The value we set is propagated to all paragraphs that are created afterwards
        builder.writeln("Paragraph 1, page 1.");
        builder.writeln("Paragraph 2, page 2.");

        doc.save(getArtifactsDir() + "Paragraph.PageBreakBefore.docx");
        //ExEnd
    }

    @Test
    public void widowControl() throws Exception
    {
        //ExStart
        //ExFor:ParagraphFormat.WidowControl
        //ExSummary:Shows how to enable widow/orphan control for a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert text that will not fit on one page, with one line spilling into page 2
        builder.getFont().setSize(68.0);
        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                        "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // This line is referred to as an "Orphan",
        // and a line left behind on the end of the previous page is likewise called a "Widow"
        // These are not ideal for readability, and the alternative to changing size/line spacing/page margins
        // in order to accomodate ill fitting text is this flag, for which the corresponding Microsoft Word option is 
        // found in Home > Paragraph > Paragraph Settings (button on the bottom right of the tab) 
        // In our document this will add more text to the orphan by putting two lines of text into the second page
        builder.getParagraphFormat().setWidowControl(true);

        doc.save(getArtifactsDir() + "Paragraph.WidowControl.docx");
        //ExEnd
    }

    @Test
    public void linesToDrop() throws Exception
    {
        //ExStart
        //ExFor:ParagraphFormat.LinesToDrop
        //ExSummary:Shows how to set the size of the drop cap text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Setting this attribute will designate the current paragraph as a drop cap,
        // in this case with a height of 4 lines of text
        builder.getParagraphFormat().setLinesToDrop(4);
        builder.write("H");

        // Any subsequent paragraphs will wrap around the drop cap
        builder.insertParagraph();
        builder.write("ello world.");

        doc.save(getArtifactsDir() + "Paragraph.LinesToDrop.odt");
        //ExEnd
    }
}
