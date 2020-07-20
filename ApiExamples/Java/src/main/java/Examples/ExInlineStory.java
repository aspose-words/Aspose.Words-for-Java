package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.awt.*;
import java.util.Date;

public class ExInlineStory extends ApiExampleBase {
    @Test
    public void addFootnote() throws Exception {
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

        Assert.assertEquals(footnote.getParagraphs().get(0).toString(SaveFormat.TEXT).trim(), "Footnote text. More text added by a DocumentBuilder.");

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
                "Footnote text. More text added by a DocumentBuilder.", (Footnote) doc.getChild(NodeType.FOOTNOTE, 0, true));
        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, false, "RefMark",
                "Footnote text.", (Footnote) doc.getChild(NodeType.FOOTNOTE, 1, true));
        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "",
                "Footnote text.", (Footnote) doc.getChild(NodeType.FOOTNOTE, 2, true));
    }

    @Test
    public void footnoteEndnote() throws Exception {
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

        Assert.assertEquals(footnote.getFootnoteType(), FootnoteType.FOOTNOTE);
        Assert.assertEquals(endnote.getFootnoteType(), FootnoteType.ENDNOTE);

        doc.save(getArtifactsDir() + "InlineStory.FootnoteEndnote.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "InlineStory.FootnoteEndnote.docx");

        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "",
                "Footnote text, will appear at the bottom of the page that contains the referenced text.", (Footnote) doc.getChild(NodeType.FOOTNOTE, 0, true));
        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, true, "",
                "Endnote text, will appear at the very end of the document.", (Footnote) doc.getChild(NodeType.FOOTNOTE, 1, true));
    }

    @Test(enabled = false, description = "WORDSJAVA-2406")
    public void addComment() throws Exception {
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

        Date currentDate = new Date();

        Comment comment = new Comment(doc, "Amy Lee", "AL", currentDate);
        builder.getCurrentParagraph().appendChild(comment);
        comment.getParagraphs().add(new Paragraph(doc));
        comment.getFirstParagraph().getRuns().add(new Run(doc, "Comment text."));

        doc.save(getArtifactsDir() + "InlineStory.AddComment.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "InlineStory.AddComment.docx");
        comment = (Comment) doc.getChild(NodeType.COMMENT, 0, true);

        Assert.assertEquals("Comment text.\r", comment.getText());
        Assert.assertEquals("Amy Lee", comment.getAuthor());
        Assert.assertEquals("AL", comment.getInitial());
        Assert.assertEquals(currentDate, comment.getDateTime());
    }

    @Test
    public void inlineStoryRevisions() throws Exception {
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
        NodeCollection footnotes = doc.getChildNodes(NodeType.FOOTNOTE, true);
        Assert.assertEquals(5, footnotes.getCount());

        // If a node was inserted in Microsoft Word while changes were being tracked, this flag will be set to true
        Footnote footnote = (Footnote) footnotes.get(2);
        Assert.assertTrue(footnote.isInsertRevision());

        // If one node was moved from one place to another while changes were tracked,
        // the node will be placed at the departure location as a "move to revision",
        // and a "move from revision" node will be left behind at the origin, in case we want to reject changes
        // Highlighting text and dragging it to another place with the mouse and cut-and-pasting (but not copy-pasting) both count as "move revisions"
        // The node with the "IsMoveToRevision" flag is the arrival of the move operation, and the node with the "IsMoveFromRevision" flag is the departure point
        footnote = (Footnote) footnotes.get(1);
        Assert.assertTrue(footnote.isMoveToRevision());
        footnote = (Footnote) footnotes.get(4);
        Assert.assertTrue(footnote.isMoveFromRevision());

        // If a node was deleted while changes were being tracked, it will stay behind as a delete revision until we accept/reject changes
        footnote = (Footnote) footnotes.get(3);
        Assert.assertTrue(footnote.isDeleteRevision());
        //ExEnd
    }

    @Test
    public void insertInlineStoryNodes() throws Exception {
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
        Assert.assertEquals(footnote.getTables().getCount(), 0);
        footnote.appendChild(table);
        Assert.assertEquals(footnote.getTables().getCount(), 1);
        Assert.assertEquals(footnote.getLastChild().getNodeType(), NodeType.TABLE);

        // An InlineStory has an "EnsureMinimum()" method as well, but in this case it makes sure the last child of the node is a paragraph,
        // so we can click and write text easily in Microsoft Word
        footnote.ensureMinimum();
        Assert.assertEquals(footnote.getLastChild().getNodeType(), NodeType.PARAGRAPH);

        // Edit the appearance of the anchor, which is the small superscript number in the main text that points to the footnote
        footnote.getFont().setName("Arial");
        footnote.getFont().setColor(Color.GREEN);

        // All inline story nodes have their own respective story types
        Assert.assertEquals(footnote.getStoryType(), StoryType.FOOTNOTES);

        // A comment is another type of inline story
        Comment comment = (Comment) builder.getCurrentParagraph().appendChild(new Comment(doc, "John Doe", "J. D.", new Date()));

        // The parent paragraph of an inline story node will be the one from the main document body
        Assert.assertEquals(comment.getParentParagraph(), doc.getFirstSection().getBody().getFirstParagraph());

        // However, the last paragraph is the one from the comment text contents, which will be outside the main document body in a speech bubble
        // A comment won't have any child nodes by default, so we can apply the EnsureMinimum() method to place a paragraph here as well
        Assert.assertNull(comment.getLastParagraph());
        comment.ensureMinimum();
        Assert.assertEquals(comment.getLastChild().getNodeType(), NodeType.PARAGRAPH);

        // Once we have a paragraph, we can move the builder do it and write our comment
        builder.moveTo(comment.getLastParagraph());
        builder.write("My comment.");

        Assert.assertEquals(comment.getStoryType(), StoryType.COMMENTS);

        doc.save(getArtifactsDir() + "InlineStory.InsertInlineStoryNodes.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "InlineStory.InsertInlineStoryNodes.docx");

        footnote = (Footnote) doc.getChild(NodeType.FOOTNOTE, 0, true);

        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "", "",
                (Footnote) doc.getChild(NodeType.FOOTNOTE, 0, true));
        Assert.assertEquals("Arial", footnote.getFont().getName());
        Assert.assertEquals(Color.GREEN.getRGB(), footnote.getFont().getColor().getRGB());

        comment = (Comment) doc.getChild(NodeType.COMMENT, 0, true);

        Assert.assertEquals("My comment.", comment.toString(SaveFormat.TEXT).trim());
    }

    @Test
    public void deleteShapes() throws Exception {
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

        Assert.assertEquals(doc.getChildNodes(NodeType.SHAPE, true).getCount(), 1);

        // We can delete all such shapes from the Body, affecting all child Paragraphs
        Assert.assertEquals(doc.getFirstSection().getBody().getStoryType(), StoryType.MAIN_TEXT);
        doc.getFirstSection().getBody().deleteShapes();

        Assert.assertEquals(doc.getChildNodes(NodeType.SHAPE, true).getCount(), 0);
        //ExEnd
    }
}
