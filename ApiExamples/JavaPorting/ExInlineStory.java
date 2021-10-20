// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.FootnotePosition;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FootnoteType;
import org.testng.Assert;
import com.aspose.words.Footnote;
import com.aspose.words.NodeType;
import com.aspose.words.EndnotePosition;
import com.aspose.words.BreakType;
import com.aspose.words.NumberStyle;
import com.aspose.words.FootnoteNumberingRule;
import com.aspose.words.Comment;
import com.aspose.ms.System.DateTime;
import com.aspose.words.Paragraph;
import java.util.ArrayList;
import com.aspose.words.Table;
import com.aspose.ms.System.Drawing.msColor;
import java.awt.Color;
import com.aspose.words.StoryType;
import java.util.Date;
import com.aspose.words.SaveFormat;
import com.aspose.words.ShapeType;
import org.testng.annotations.DataProvider;


@Test
public class ExInlineStory extends ApiExampleBase
{
    @Test (dataProvider = "positionFootnoteDataProvider")
    public void positionFootnote(/*FootnotePosition*/int footnotePosition) throws Exception
    {
        //ExStart
        //ExFor:Document.FootnoteOptions
        //ExFor:FootnoteOptions
        //ExFor:FootnoteOptions.Position
        //ExFor:FootnotePosition
        //ExSummary:Shows how to select a different place where the document collects and displays its footnotes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // A footnote is a way to attach a reference or a side comment to text
        // that does not interfere with the main body text's flow.  
        // Inserting a footnote adds a small superscript reference symbol
        // at the main body text where we insert the footnote.
        // Each footnote also creates an entry at the bottom of the page, consisting of a symbol
        // that matches the reference symbol in the main body text.
        // The reference text that we pass to the document builder's "InsertFootnote" method.
        builder.write("Hello world!");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote contents.");

        // We can use the "Position" property to determine where the document will place all its footnotes.
        // If we set the value of the "Position" property to "FootnotePosition.BottomOfPage",
        // every footnote will show up at the bottom of the page that contains its reference mark. This is the default value.
        // If we set the value of the "Position" property to "FootnotePosition.BeneathText",
        // every footnote will show up at the end of the page's text that contains its reference mark.
        doc.getFootnoteOptions().setPosition(footnotePosition);

        doc.save(getArtifactsDir() + "InlineStory.PositionFootnote.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "InlineStory.PositionFootnote.docx");

        Assert.assertEquals(footnotePosition, doc.getFootnoteOptions().getPosition());

        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "",
            "Footnote contents.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 0, true));
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "positionFootnoteDataProvider")
	public static Object[][] positionFootnoteDataProvider() throws Exception
	{
		return new Object[][]
		{
			{FootnotePosition.BENEATH_TEXT},
			{FootnotePosition.BOTTOM_OF_PAGE},
		};
	}

    @Test (dataProvider = "positionEndnoteDataProvider")
    public void positionEndnote(/*EndnotePosition*/int endnotePosition) throws Exception
    {
        //ExStart
        //ExFor:Document.EndnoteOptions
        //ExFor:EndnoteOptions
        //ExFor:EndnoteOptions.Position
        //ExFor:EndnotePosition
        //ExSummary:Shows how to select a different place where the document collects and displays its endnotes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // An endnote is a way to attach a reference or a side comment to text
        // that does not interfere with the main body text's flow. 
        // Inserting an endnote adds a small superscript reference symbol
        // at the main body text where we insert the endnote.
        // Each endnote also creates an entry at the end of the document, consisting of a symbol
        // that matches the reference symbol in the main body text.
        // The reference text that we pass to the document builder's "InsertEndnote" method.
        builder.write("Hello world!");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Endnote contents.");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.write("This is the second section.");

        // We can use the "Position" property to determine where the document will place all its endnotes.
        // If we set the value of the "Position" property to "EndnotePosition.EndOfDocument",
        // every footnote will show up in a collection at the end of the document. This is the default value.
        // If we set the value of the "Position" property to "EndnotePosition.EndOfSection",
        // every footnote will show up in a collection at the end of the section whose text contains the endnote's reference mark.
        doc.getEndnoteOptions().setPosition(endnotePosition);

        doc.save(getArtifactsDir() + "InlineStory.PositionEndnote.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "InlineStory.PositionEndnote.docx");

        Assert.assertEquals(endnotePosition, doc.getEndnoteOptions().getPosition());

        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, true, "",
            "Endnote contents.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 0, true));
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "positionEndnoteDataProvider")
	public static Object[][] positionEndnoteDataProvider() throws Exception
	{
		return new Object[][]
		{
			{EndnotePosition.END_OF_DOCUMENT},
			{EndnotePosition.END_OF_SECTION},
		};
	}

    @Test
    public void refMarkNumberStyle() throws Exception
    {
        //ExStart
        //ExFor:Document.EndnoteOptions
        //ExFor:EndnoteOptions
        //ExFor:EndnoteOptions.NumberStyle
        //ExFor:Document.FootnoteOptions
        //ExFor:FootnoteOptions
        //ExFor:FootnoteOptions.NumberStyle
        //ExSummary:Shows how to change the number style of footnote/endnote reference marks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Footnotes and endnotes are a way to attach a reference or a side comment to text
        // that does not interfere with the main body text's flow. 
        // Inserting a footnote/endnote adds a small superscript reference symbol
        // at the main body text where we insert the footnote/endnote.
        // Each footnote/endnote also creates an entry, which consists of a symbol that matches the reference
        // symbol in the main body text. The reference text that we pass to the document builder's "InsertEndnote" method.
        // Footnote entries, by default, show up at the bottom of each page that contains
        // their reference symbols, and endnotes show up at the end of the document.
        builder.write("Text 1. ");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote 1.");
        builder.write("Text 2. ");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote 2.");
        builder.write("Text 3. ");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote 3.", "Custom footnote reference mark");

        builder.insertParagraph();

        builder.write("Text 1. ");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Endnote 1.");
        builder.write("Text 2. ");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Endnote 2.");
        builder.write("Text 3. ");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Endnote 3.", "Custom endnote reference mark");

        // By default, the reference symbol for each footnote and endnote is its index
        // among all the document's footnotes/endnotes. Each document maintains separate counts
        // for footnotes and for endnotes. By default, footnotes display their numbers using Arabic numerals,
        // and endnotes display their numbers in lowercase Roman numerals.
        Assert.assertEquals(NumberStyle.ARABIC, doc.getFootnoteOptions().getNumberStyle());
        Assert.assertEquals(NumberStyle.LOWERCASE_ROMAN, doc.getEndnoteOptions().getNumberStyle());

        // We can use the "NumberStyle" property to apply custom numbering styles to footnotes and endnotes.
        // This will not affect footnotes/endnotes with custom reference marks.
        doc.getFootnoteOptions().setNumberStyle(NumberStyle.UPPERCASE_ROMAN);
        doc.getEndnoteOptions().setNumberStyle(NumberStyle.UPPERCASE_LETTER);

        doc.save(getArtifactsDir() + "InlineStory.RefMarkNumberStyle.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "InlineStory.RefMarkNumberStyle.docx");

        Assert.assertEquals(NumberStyle.UPPERCASE_ROMAN, doc.getFootnoteOptions().getNumberStyle());
        Assert.assertEquals(NumberStyle.UPPERCASE_LETTER, doc.getEndnoteOptions().getNumberStyle());

        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "",
            "Footnote 1.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 0, true));
        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "",
            "Footnote 2.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 1, true));
        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, false, "Custom footnote reference mark",
            "Custom footnote reference mark Footnote 3.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 2, true));
        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, true, "",
            "Endnote 1.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 3, true));
        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, true, "",
            "Endnote 2.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 4, true));
        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, false, "Custom endnote reference mark",
            "Custom endnote reference mark Endnote 3.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 5, true));
    }

    @Test
    public void numberingRule() throws Exception
    {
        //ExStart
        //ExFor:Document.EndnoteOptions
        //ExFor:EndnoteOptions
        //ExFor:EndnoteOptions.RestartRule
        //ExFor:FootnoteNumberingRule
        //ExFor:Document.FootnoteOptions
        //ExFor:FootnoteOptions
        //ExFor:FootnoteOptions.RestartRule
        //ExSummary:Shows how to restart footnote/endnote numbering at certain places in the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Footnotes and endnotes are a way to attach a reference or a side comment to text
        // that does not interfere with the main body text's flow. 
        // Inserting a footnote/endnote adds a small superscript reference symbol
        // at the main body text where we insert the footnote/endnote.
        // Each footnote/endnote also creates an entry, which consists of a symbol that matches the reference
        // symbol in the main body text. The reference text that we pass to the document builder's "InsertEndnote" method.
        // Footnote entries, by default, show up at the bottom of each page that contains
        // their reference symbols, and endnotes show up at the end of the document.
        builder.write("Text 1. ");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote 1.");
        builder.write("Text 2. ");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote 2.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.write("Text 3. ");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote 3.");
        builder.write("Text 4. ");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote 4.");

        builder.insertBreak(BreakType.PAGE_BREAK);

        builder.write("Text 1. ");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Endnote 1.");
        builder.write("Text 2. ");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Endnote 2.");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.write("Text 3. ");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Endnote 3.");
        builder.write("Text 4. ");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Endnote 4.");

        // By default, the reference symbol for each footnote and endnote is its index
        // among all the document's footnotes/endnotes. Each document maintains separate counts
        // for footnotes and endnotes and does not restart these counts at any point.
        Assert.assertEquals(doc.getFootnoteOptions().getRestartRule(), FootnoteNumberingRule.DEFAULT);
        Assert.assertEquals(FootnoteNumberingRule.DEFAULT, FootnoteNumberingRule.CONTINUOUS);

        // We can use the "RestartRule" property to get the document to restart
        // the footnote/endnote counts at a new page or section.
        doc.getFootnoteOptions().setRestartRule(FootnoteNumberingRule.RESTART_PAGE);
        doc.getEndnoteOptions().setRestartRule(FootnoteNumberingRule.RESTART_SECTION);

        doc.save(getArtifactsDir() + "InlineStory.NumberingRule.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "InlineStory.NumberingRule.docx");

        Assert.assertEquals(FootnoteNumberingRule.RESTART_PAGE, doc.getFootnoteOptions().getRestartRule());
        Assert.assertEquals(FootnoteNumberingRule.RESTART_SECTION, doc.getEndnoteOptions().getRestartRule());

        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "",
            "Footnote 1.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 0, true));
        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "",
            "Footnote 2.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 1, true));
        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "",
            "Footnote 3.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 2, true));
        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "",
            "Footnote 4.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 3, true));
        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, true, "",
            "Endnote 1.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 4, true));
        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, true, "",
            "Endnote 2.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 5, true));
        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, true, "",
            "Endnote 3.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 6, true));
        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, true, "",
            "Endnote 4.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 7, true));
    }

    @Test
    public void startNumber() throws Exception
    {
        //ExStart
        //ExFor:Document.EndnoteOptions
        //ExFor:EndnoteOptions
        //ExFor:EndnoteOptions.StartNumber
        //ExFor:Document.FootnoteOptions
        //ExFor:FootnoteOptions
        //ExFor:FootnoteOptions.StartNumber
        //ExSummary:Shows how to set a number at which the document begins the footnote/endnote count.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Footnotes and endnotes are a way to attach a reference or a side comment to text
        // that does not interfere with the main body text's flow. 
        // Inserting a footnote/endnote adds a small superscript reference symbol
        // at the main body text where we insert the footnote/endnote.
        // Each footnote/endnote also creates an entry, which consists of a symbol
        // that matches the reference symbol in the main body text.
        // The reference text that we pass to the document builder's "InsertEndnote" method.
        // Footnote entries, by default, show up at the bottom of each page that contains
        // their reference symbols, and endnotes show up at the end of the document.
        builder.write("Text 1. ");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote 1.");
        builder.write("Text 2. ");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote 2.");
        builder.write("Text 3. ");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote 3.");

        builder.insertParagraph();

        builder.write("Text 1. ");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Endnote 1.");
        builder.write("Text 2. ");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Endnote 2.");
        builder.write("Text 3. ");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Endnote 3.");

        // By default, the reference symbol for each footnote and endnote is its index
        // among all the document's footnotes/endnotes. Each document maintains separate counts
        // for footnotes and for endnotes, which both begin at 1.
        Assert.assertEquals(1, doc.getFootnoteOptions().getStartNumber());
        Assert.assertEquals(1, doc.getEndnoteOptions().getStartNumber());

        // We can use the "StartNumber" property to get the document to
        // begin a footnote or endnote count at a different number.
        doc.getEndnoteOptions().setNumberStyle(NumberStyle.ARABIC);
        doc.getEndnoteOptions().setStartNumber(50);

        doc.save(getArtifactsDir() + "InlineStory.StartNumber.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "InlineStory.StartNumber.docx");

        Assert.assertEquals(1, doc.getFootnoteOptions().getStartNumber());
        Assert.assertEquals(50, doc.getEndnoteOptions().getStartNumber());
        Assert.assertEquals(NumberStyle.ARABIC, doc.getFootnoteOptions().getNumberStyle());
        Assert.assertEquals(NumberStyle.ARABIC, doc.getEndnoteOptions().getNumberStyle());

        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "",
            "Footnote 1.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 0, true));
        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "",
            "Footnote 2.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 1, true));
        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "",
            "Footnote 3.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 2, true));
        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, true, "",
            "Endnote 1.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 3, true));
        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, true, "",
            "Endnote 2.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 4, true));
        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, true, "",
            "Endnote 3.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 5, true));
    }

    @Test
    public void addFootnote() throws Exception
    {
        //ExStart
        //ExFor:Footnote
        //ExFor:Footnote.IsAuto
        //ExFor:Footnote.ReferenceMark
        //ExFor:InlineStory
        //ExFor:InlineStory.Paragraphs
        //ExFor:InlineStory.FirstParagraph
        //ExFor:FootnoteType
        //ExFor:Footnote.#ctor
        //ExSummary:Shows how to insert and customize footnotes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add text, and reference it with a footnote. This footnote will place a small superscript reference
        // mark after the text that it references and create an entry below the main body text at the bottom of the page.
        // This entry will contain the footnote's reference mark and the reference text,
        // which we will pass to the document builder's "InsertFootnote" method.
        builder.write("Main body text.");
        Footnote footnote = builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote text.");

        // If this property is set to "true", then our footnote's reference mark
        // will be its index among all the section's footnotes.
        // This is the first footnote, so the reference mark will be "1".
        Assert.assertTrue(footnote.isAuto());

        // We can move the document builder inside the footnote to edit its reference text. 
        builder.moveTo(footnote.getFirstParagraph());
        builder.write(" More text added by a DocumentBuilder.");
        builder.moveToDocumentEnd();

        Assert.assertEquals("\u0002 Footnote text. More text added by a DocumentBuilder.", footnote.getText().trim());

        builder.write(" More main body text.");
        footnote = builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote text.");

        // We can set a custom reference mark which the footnote will use instead of its index number.
        footnote.setReferenceMark("RefMark");

        Assert.assertFalse(footnote.isAuto());

        // A bookmark with the "IsAuto" flag set to true will still show its real index
        // even if previous bookmarks display custom reference marks, so this bookmark's reference mark will be a "3".
        builder.write(" More main body text.");
        footnote = builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote text.");

        Assert.assertTrue(footnote.isAuto());

        doc.save(getArtifactsDir() + "InlineStory.AddFootnote.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "InlineStory.AddFootnote.docx");

        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "", 
            "Footnote text. More text added by a DocumentBuilder.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 0, true));
        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, false, "RefMark", 
            "Footnote text.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 1, true));
        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "", 
            "Footnote text.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 2, true));
    }

    @Test
    public void footnoteEndnote() throws Exception
    {
        //ExStart
        //ExFor:Footnote.FootnoteType
        //ExSummary:Shows the difference between footnotes and endnotes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Below are two ways of attaching numbered references to the text. Both these references will add a
        // small superscript reference mark at the location that we insert them.
        // The reference mark, by default, is the index number of the reference among all the references in the document.
        // Each reference will also create an entry, which will have the same reference mark as in the body text
        // and reference text, which we will pass to the document builder's "InsertFootnote" method.
        // 1 -  A footnote, whose entry will appear on the same page as the text that it references:
        builder.write("Footnote referenced main body text.");
        Footnote footnote = builder.insertFootnote(FootnoteType.FOOTNOTE, 
            "Footnote text, will appear at the bottom of the page that contains the referenced text.");

        // 2 -  An endnote, whose entry will appear at the end of the document:
        builder.write("Endnote referenced main body text.");
        Footnote endnote = builder.insertFootnote(FootnoteType.ENDNOTE, 
            "Endnote text, will appear at the very end of the document.");

        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);

        Assert.assertEquals(FootnoteType.FOOTNOTE, footnote.getFootnoteType());
        Assert.assertEquals(FootnoteType.ENDNOTE, endnote.getFootnoteType());

        doc.save(getArtifactsDir() + "InlineStory.FootnoteEndnote.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "InlineStory.FootnoteEndnote.docx");

        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "",
            "Footnote text, will appear at the bottom of the page that contains the referenced text.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 0, true));
        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, true, "",
            "Endnote text, will appear at the very end of the document.", (Footnote)doc.getChild(NodeType.FOOTNOTE, 1, true));
    }

    @Test
    public void addComment() throws Exception
    {
        //ExStart
        //ExFor:Comment
        //ExFor:InlineStory
        //ExFor:InlineStory.Paragraphs
        //ExFor:InlineStory.FirstParagraph
        //ExFor:Comment.#ctor(DocumentBase, String, String, DateTime)
        //ExSummary:Shows how to add a comment to a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Hello world!");

        Comment comment = new Comment(doc, "John Doe", "JD", DateTime.getToday());
        builder.getCurrentParagraph().appendChild(comment);
        builder.moveTo(comment.appendChild(new Paragraph(doc)));
        builder.write("Comment text.");

        Assert.assertEquals(DateTime.getToday(), comment.getDateTimeInternal());

        // In Microsoft Word, we can right-click this comment in the document body to edit it, or reply to it. 
        doc.save(getArtifactsDir() + "InlineStory.AddComment.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "InlineStory.AddComment.docx");
        comment = (Comment)doc.getChild(NodeType.COMMENT, 0, true);
        
        Assert.assertEquals("Comment text.\r", comment.getText());
        Assert.assertEquals("John Doe", comment.getAuthor());
        Assert.assertEquals("JD", comment.getInitial());
        Assert.assertEquals(DateTime.getToday(), comment.getDateTimeInternal());
    }

    @Test
    public void inlineStoryRevisions() throws Exception
    {
        //ExStart
        //ExFor:InlineStory.IsDeleteRevision
        //ExFor:InlineStory.IsInsertRevision
        //ExFor:InlineStory.IsMoveFromRevision
        //ExFor:InlineStory.IsMoveToRevision
        //ExSummary:Shows how to view revision-related properties of InlineStory nodes.
        Document doc = new Document(getMyDir() + "Revision footnotes.docx");

        // When we edit the document while the "Track Changes" option, found in via Review -> Tracking,
        // is turned on in Microsoft Word, the changes we apply count as revisions.
        // When editing a document using Aspose.Words, we can begin tracking revisions by
        // invoking the document's "StartTrackRevisions" method and stop tracking by using the "StopTrackRevisions" method.
        // We can either accept revisions to assimilate them into the document
        // or reject them to undo and discard the proposed change.
        Assert.assertTrue(doc.hasRevisions());

        ArrayList<Footnote> footnotes = doc.getChildNodes(NodeType.FOOTNOTE, true).<Footnote>Cast().ToList();

        Assert.assertEquals(5, footnotes.size());

        // Below are five types of revisions that can flag an InlineStory node.
        // 1 -  An "insert" revision:
        // This revision occurs when we insert text while tracking changes.
        Assert.assertTrue(footnotes.get(2).isInsertRevision());

        // 2 -  A "move from" revision:
        // When we highlight text in Microsoft Word, and then drag it to a different place in the document
        // while tracking changes, two revisions appear.
        // The "move from" revision is a copy of the text originally before we moved it.
        Assert.assertTrue(footnotes.get(4).isMoveFromRevision());

        // 3 -  A "move to" revision:
        // The "move to" revision is the text that we moved in its new position in the document.
        // "Move from" and "move to" revisions appear in pairs for every move revision we carry out.
        // Accepting a move revision deletes the "move from" revision and its text,
        // and keeps the text from the "move to" revision.
        // Rejecting a move revision conversely keeps the "move from" revision and deletes the "move to" revision.
        Assert.assertTrue(footnotes.get(1).isMoveToRevision());

        // 4 -  A "delete" revision:
        // This revision occurs when we delete text while tracking changes. When we delete text like this,
        // it will stay in the document as a revision until we either accept the revision,
        // which will delete the text for good, or reject the revision, which will keep the text we deleted where it was.
        Assert.assertTrue(footnotes.get(3).isDeleteRevision());
        //ExEnd
    }

    @Test
    public void insertInlineStoryNodes() throws Exception
    {
        //ExStart
        //ExFor:Comment.StoryType
        //ExFor:Footnote.StoryType
        //ExFor:InlineStory.EnsureMinimum
        //ExFor:InlineStory.Font
        //ExFor:InlineStory.LastParagraph
        //ExFor:InlineStory.ParentParagraph
        //ExFor:InlineStory.StoryType
        //ExFor:InlineStory.Tables
        //ExSummary:Shows how to insert InlineStory nodes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Footnote footnote = builder.insertFootnote(FootnoteType.FOOTNOTE, null);

        // Table nodes have an "EnsureMinimum()" method that makes sure the table has at least one cell.
        Table table = new Table(doc);
        table.ensureMinimum();

        // We can place a table inside a footnote, which will make it appear at the referencing page's footer.
        Assert.That(footnote.getTables(), Is.Empty);
        footnote.appendChild(table);
        Assert.assertEquals(1, footnote.getTables().getCount());
        Assert.assertEquals(NodeType.TABLE, footnote.getLastChild().getNodeType());

        // An InlineStory has an "EnsureMinimum()" method as well, but in this case,
        // it makes sure the last child of the node is a paragraph,
        // for us to be able to click and write text easily in Microsoft Word.
        footnote.ensureMinimum();
        Assert.assertEquals(NodeType.PARAGRAPH, footnote.getLastChild().getNodeType());

        // Edit the appearance of the anchor, which is the small superscript number
        // in the main text that points to the footnote.
        footnote.getFont().setName("Arial");
        footnote.getFont().setColor(msColor.getGreen());

        // All inline story nodes have their respective story types.
        Assert.assertEquals(StoryType.FOOTNOTES, footnote.getStoryType());

        // A comment is another type of inline story.
        Comment comment = (Comment)builder.getCurrentParagraph().appendChild(new Comment(doc, "John Doe", "J. D.", new Date()));

        // The parent paragraph of an inline story node will be the one from the main document body.
        Assert.assertEquals(doc.getFirstSection().getBody().getFirstParagraph(), comment.getParentParagraph());

        // However, the last paragraph is the one from the comment text contents,
        // which will be outside the main document body in a speech bubble.
        // A comment will not have any child nodes by default,
        // so we can apply the EnsureMinimum() method to place a paragraph here as well.
        Assert.assertNull(comment.getLastParagraph());
        comment.ensureMinimum();
        Assert.assertEquals(NodeType.PARAGRAPH, comment.getLastChild().getNodeType());

        // Once we have a paragraph, we can move the builder to do it and write our comment.
        builder.moveTo(comment.getLastParagraph());
        builder.write("My comment.");

        Assert.assertEquals(StoryType.COMMENTS, comment.getStoryType());

        doc.save(getArtifactsDir() + "InlineStory.InsertInlineStoryNodes.docx");
        //ExEnd
        
        doc = new Document(getArtifactsDir() + "InlineStory.InsertInlineStoryNodes.docx");

        footnote = (Footnote)doc.getChild(NodeType.FOOTNOTE, 0, true);

        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "", "", 
            (Footnote)doc.getChild(NodeType.FOOTNOTE, 0, true));
        Assert.assertEquals("Arial", footnote.getFont().getName());
        Assert.assertEquals(msColor.getGreen().getRGB(), footnote.getFont().getColor().getRGB());

        comment = (Comment)doc.getChild(NodeType.COMMENT, 0, true);

        Assert.assertEquals("My comment.", comment.toString(SaveFormat.TEXT).trim());
    }

    @Test
    public void deleteShapes() throws Exception
    {
        //ExStart
        //ExFor:Story
        //ExFor:Story.DeleteShapes
        //ExFor:Story.StoryType
        //ExFor:StoryType
        //ExSummary:Shows how to remove all shapes from a node.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a DocumentBuilder to insert a shape. This is an inline shape,
        // which has a parent Paragraph, which is a child node of the first section's Body.
        builder.insertShape(ShapeType.CUBE, 100.0, 100.0);

        Assert.assertEquals(1, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        // We can delete all shapes from the child paragraphs of this Body.
        Assert.assertEquals(StoryType.MAIN_TEXT, doc.getFirstSection().getBody().getStoryType());
        doc.getFirstSection().getBody().deleteShapes();

        Assert.assertEquals(0, doc.getChildNodes(NodeType.SHAPE, true).getCount());
        //ExEnd
    }
}
