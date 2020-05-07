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
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Footnote;
import com.aspose.words.FootnoteType;
import org.testng.Assert;
import com.aspose.ms.System.msString;
import com.aspose.words.SaveFormat;
import com.aspose.words.NodeType;
import com.aspose.words.BreakType;
import com.aspose.words.Comment;
import com.aspose.ms.System.DateTime;
import com.aspose.words.Paragraph;
import com.aspose.words.Run;
import java.util.ArrayList;
import com.aspose.words.Table;
import com.aspose.ms.System.Drawing.msColor;
import java.awt.Color;
import com.aspose.words.StoryType;
import com.aspose.words.ShapeType;


@Test
public class ExInlineStory extends ApiExampleBase
{
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
        //ExSummary:Shows how to add a footnote to a paragraph in the document and set its marker.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add text that will be referenced by a footnote
        builder.write("Main body text.");

        // Add a footnote and give it text, which will appear at the bottom of the page
        Footnote footnote = builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote text.");

        // This attribute means the footnote in the main text will automatically be assigned a number, "1" in this instance
        // The next footnote will get "2"
        Assert.assertTrue(footnote.isAuto());

        // We can edit the footnote's text like this
        // Make sure to move the builder back into the document body afterwards
        builder.moveTo(footnote.getFirstParagraph());
        builder.write(" More text added by a DocumentBuilder.");
        builder.moveToDocumentEnd();

        Assert.assertEquals("Footnote text. More text added by a DocumentBuilder.", msString.trim(footnote.getParagraphs().get(0).toString(SaveFormat.TEXT)));

        builder.write(" More main body text.");
        footnote = builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote text.");

        // Substitute the reference number for our own custom mark by setting this variable, "IsAuto" will also be set to false
        footnote.setReferenceMark("RefMark");
        Assert.assertFalse(footnote.isAuto());

        // This bookmark will get a number "3" even though there was no "2"
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
        //ExSummary:Demonstrates the difference between footnotes and endnotes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write text and insert a footnote to reference it at the bottom of the page
        builder.write("Footnote referenced main body text.");
        Footnote footnote = builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote text, will appear at the bottom of the page that contains the referenced text.");

        // Write text and insert an endnote to reference it at the end of the document
        builder.write("Endnote referenced main body text.");
        Footnote endnote = builder.insertFootnote(FootnoteType.ENDNOTE, "Endnote text, will appear at the very end of the document.");

        // Since endnotes are at the end of the document, breaks like this will push them down while the footnotes stay where they are
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
        //ExSummary:Shows how to add a comment to a paragraph in the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Some text is added.");

        Comment comment = new Comment(doc, "Amy Lee", "AL", DateTime.getToday());
        builder.getCurrentParagraph().appendChild(comment);
        comment.getParagraphs().add(new Paragraph(doc));
        comment.getFirstParagraph().getRuns().add(new Run(doc, "Comment text."));

        doc.save(getArtifactsDir() + "InlineStory.AddComment.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "InlineStory.AddComment.docx");
        comment = (Comment)doc.getChild(NodeType.COMMENT, 0, true);
        
        Assert.assertEquals("Comment text.\r", comment.getText());
        Assert.assertEquals("Amy Lee", comment.getAuthor());
        Assert.assertEquals("AL", comment.getInitial());
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
        // Open a document that has revisions from changes being tracked
        Document doc = new Document(getMyDir() + "Revision footnotes.docx");
        Assert.assertTrue(doc.hasRevisions());

        // Get a collection of all footnotes from the document
        ArrayList<Footnote> footnotes = doc.getChildNodes(NodeType.FOOTNOTE, true).<Footnote>Cast().ToList();
        Assert.assertEquals(5, footnotes.size());

        // If a node was inserted in Microsoft Word while changes were being tracked, this flag will be set to true
        Assert.assertTrue(footnotes.get(2).isInsertRevision());

        // If one node was moved from one place to another while changes were tracked,
        // the node will be placed at the departure location as a "move to revision",
        // and a "move from revision" node will be left behind at the origin, in case we want to reject changes
        // Highlighting text and dragging it to another place with the mouse and cut-and-pasting (but not copy-pasting) both count as "move revisions"
        // The node with the "IsMoveToRevision" flag is the arrival of the move operation, and the node with the "IsMoveFromRevision" flag is the departure point
        Assert.assertTrue(footnotes.get(1).isMoveToRevision());
        Assert.assertTrue(footnotes.get(4).isMoveFromRevision());

        // If a node was deleted while changes were being tracked, it will stay behind as a delete revision until we accept/reject changes
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

        // Table nodes have an "EnsureMinimum()" method that makes sure the table has at least one cell
        Table table = new Table(doc);
        table.ensureMinimum();

        // We can place a table inside a footnote, which will make it appear at the footer of the referencing page
        Assert.That(footnote.getTables(), Is.Empty);
        footnote.appendChild(table);
        Assert.assertEquals(1, footnote.getTables().getCount());
        Assert.assertEquals(NodeType.TABLE, footnote.getLastChild().getNodeType());

        // An InlineStory has an "EnsureMinimum()" method as well, but in this case it makes sure the last child of the node is a paragraph,
        // so we can click and write text easily in Microsoft Word
        footnote.ensureMinimum();
        Assert.assertEquals(NodeType.PARAGRAPH, footnote.getLastChild().getNodeType());

        // Edit the appearance of the anchor, which is the small superscript number in the main text that points to the footnote
        footnote.getFont().setName("Arial");
        footnote.getFont().setColor(msColor.getGreen());

        // All inline story nodes have their own respective story types
        Assert.assertEquals(StoryType.FOOTNOTES, footnote.getStoryType());

        // A comment is another type of inline story
        Comment comment = (Comment)builder.getCurrentParagraph().appendChild(new Comment(doc, "John Doe", "J. D.", DateTime.getNow()));

        // The parent paragraph of an inline story node will be the one from the main document body
        Assert.assertEquals(doc.getFirstSection().getBody().getFirstParagraph(), comment.getParentParagraph());

        // However, the last paragraph is the one from the comment text contents, which will be outside the main document body in a speech bubble
        // A comment won't have any child nodes by default, so we can apply the EnsureMinimum() method to place a paragraph here as well
        Assert.assertNull(comment.getLastParagraph());
        comment.ensureMinimum();
        Assert.assertEquals(NodeType.PARAGRAPH, comment.getLastChild().getNodeType());

        // Once we have a paragraph, we can move the builder do it and write our comment
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

        Assert.assertEquals("My comment.", msString.trim(comment.toString(SaveFormat.TEXT)));
    }

    @Test
    public void deleteShapes() throws Exception
    {
        //ExStart
        //ExFor:Story
        //ExFor:Story.DeleteShapes
        //ExFor:Story.StoryType
        //ExFor:StoryType
        //ExSummary:Shows how to clear a body of inline shapes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a DocumentBuilder to insert a shape
        // This is an inline shape, which has a parent Paragraph, which is in turn a child of the Body
        builder.insertShape(ShapeType.CUBE, 100.0, 100.0);

        Assert.assertEquals(1, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        // We can delete all such shapes from the Body, affecting all child Paragraphs
        Assert.assertEquals(StoryType.MAIN_TEXT, doc.getFirstSection().getBody().getStoryType());
        doc.getFirstSection().getBody().deleteShapes();

        Assert.assertEquals(0, doc.getChildNodes(NodeType.SHAPE, true).getCount());
        //ExEnd
    }
}
