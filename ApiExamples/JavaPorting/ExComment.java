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
import com.aspose.words.Comment;
import java.util.Date;
import com.aspose.ms.System.DateTime;
import org.testng.Assert;
import com.aspose.words.NodeType;
import com.aspose.words.NodeCollection;
import com.aspose.ms.System.msConsole;
import com.aspose.words.Paragraph;
import com.aspose.words.CommentRangeStart;
import com.aspose.words.Run;
import com.aspose.words.CommentRangeEnd;
import com.aspose.words.DocumentVisitor;
import com.aspose.words.VisitorAction;
import com.aspose.ms.System.Text.msStringBuilder;


@Test
public class ExComment extends ApiExampleBase
{
    @Test
    public void addCommentWithReply() throws Exception
    {
        //ExStart
        //ExFor:Comment
        //ExFor:Comment.SetText(String)
        //ExFor:Comment.AddReply(String, String, DateTime, String)
        //ExSummary:Shows how to add a comment to a document, and then reply to it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Comment comment = new Comment(doc, "John Doe", "J.D.", new Date());
        comment.setText("My comment.");
        
        // Place the comment at a node in the document's body.
        // This comment will show up at the location of its paragraph,
        // outside the right-side margin of the page, and with a dotted line connecting it to its paragraph.
        builder.getCurrentParagraph().appendChild(comment);

        // Add a reply, which will show up under its parent comment.
        comment.addReplyInternal("Joe Bloggs", "J.B.", new Date(), "New reply");

        // Comments and replies are both Comment nodes.
        Assert.assertEquals(2, doc.getChildNodes(NodeType.COMMENT, true).getCount());

        // Comments that do not reply to other comments are "top-level". They have no ancestor comments.
        Assert.assertNull(comment.getAncestor());

        // Replies have an ancestor top-level comment.
        Assert.assertEquals(comment, comment.getReplies().get(0).getAncestor());

        doc.save(getArtifactsDir() + "Comment.AddCommentWithReply.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Comment.AddCommentWithReply.docx");
        Comment docComment = (Comment)doc.getChild(NodeType.COMMENT, 0, true);

        Assert.assertEquals(1, docComment.getCount());
        Assert.assertEquals(1, comment.getReplies().getCount());

        Assert.assertEquals("\u0005My comment.\r", docComment.getText());
        Assert.assertEquals("\u0005New reply\r", docComment.getReplies().get(0).getText());
    }

    @Test
    public void printAllComments() throws Exception
    {
        //ExStart
        //ExFor:Comment.Ancestor
        //ExFor:Comment.Author
        //ExFor:Comment.Replies
        //ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
        //ExSummary:Shows how to print all of a document's comments and their replies.
        Document doc = new Document(getMyDir() + "Comments.docx");

        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
        Assert.assertEquals(12, comments.getCount()); //ExSkip

        // If a comment has no ancestor, it is a "top-level" comment as opposed to a reply-type comment.
        // Print all top-level comments along with any replies they may have.
        for (Comment comment : comments.<Comment>OfType().Where(c => c.Ancestor == null) !!Autoporter error: Undefined expression type )
        {
            System.out.println("Top-level comment:");
            System.out.println("\t\"{comment.GetText().Trim()}\", by {comment.Author}");
            System.out.println("Has {comment.Replies.Count} replies");
            for (Comment commentReply : (Iterable<Comment>) comment.getReplies())
            {
                System.out.println("\t\"{commentReply.GetText().Trim()}\", by {commentReply.Author}");
            }
            msConsole.writeLine();
        }
        //ExEnd
    }

    @Test
    public void removeCommentReplies() throws Exception
    {
        //ExStart
        //ExFor:Comment.RemoveAllReplies
        //ExFor:Comment.RemoveReply(Comment)
        //ExFor:CommentCollection.Item(Int32)
        //ExSummary:Shows how to remove comment replies.
        Document doc = new Document();

        Comment comment = new Comment(doc, "John Doe", "J.D.", new Date());
        comment.setText("My comment.");

        doc.getFirstSection().getBody().getFirstParagraph().appendChild(comment);
        
        comment.addReplyInternal("Joe Bloggs", "J.B.", new Date(), "New reply");
        comment.addReplyInternal("Joe Bloggs", "J.B.", new Date(), "Another reply");

        Assert.AreEqual(2, comment.getReplies().Count()); 

        // Below are two ways of removing replies from a comment.
        // 1 -  Use the "RemoveReply" method to remove replies from a comment individually:
        comment.removeReply(comment.getReplies().get(0));

        Assert.AreEqual(1, comment.getReplies().Count());

        // 2 -  Use the "RemoveAllReplies" method to remove all replies from a comment at once:
        comment.removeAllReplies();

        Assert.AreEqual(0, comment.getReplies().Count()); 
        //ExEnd
    }

    @Test
    public void done() throws Exception
    {
        //ExStart
        //ExFor:Comment.Done
        //ExFor:CommentCollection
        //ExSummary:Shows how to mark a comment as "done".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Helo world!");

        // Insert a comment to point out an error. 
        Comment comment = new Comment(doc, "John Doe", "J.D.", new Date());
        comment.setText("Fix the spelling error!");
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(comment);

        // Comments have a "Done" flag, which is set to "false" by default. 
        // If a comment suggests that we make a change within the document,
        // we can apply the change, and then also set the "Done" flag afterwards to indicate the correction.
        Assert.assertFalse(comment.getDone());

        doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0).setText("Hello world!");
        comment.setDone(true);

        // Comments that are "done" will differentiate themselves
        // from ones that are not "done" with a faded text color.
        comment = new Comment(doc, "John Doe", "J.D.", new Date());
        comment.setText("Add text to this paragraph.");
        builder.getCurrentParagraph().appendChild(comment);

        doc.save(getArtifactsDir() + "Comment.Done.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Comment.Done.docx");
        comment = (Comment)doc.getChildNodes(NodeType.COMMENT, true).get(0);

        Assert.assertTrue(comment.getDone());
        Assert.assertEquals("\u0005Fix the spelling error!", comment.getText().trim());
        Assert.assertEquals("Hello world!", doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0).getText());
    }
    
    //ExStart
    //ExFor:Comment.Done
    //ExFor:Comment.#ctor(DocumentBase)
    //ExFor:Comment.Accept(DocumentVisitor)
    //ExFor:Comment.DateTime
    //ExFor:Comment.Id
    //ExFor:Comment.Initial
    //ExFor:CommentRangeEnd
    //ExFor:CommentRangeEnd.#ctor(DocumentBase,Int32)
    //ExFor:CommentRangeEnd.Accept(DocumentVisitor)
    //ExFor:CommentRangeEnd.Id
    //ExFor:CommentRangeStart
    //ExFor:CommentRangeStart.#ctor(DocumentBase,Int32)
    //ExFor:CommentRangeStart.Accept(DocumentVisitor)
    //ExFor:CommentRangeStart.Id
    //ExSummary:Shows how print the contents of all comments and their comment ranges using a document visitor.
    @Test //ExSkip
    public void createCommentsAndPrintAllInfo() throws Exception
    {
        Document doc = new Document();
        
        Comment newComment = new Comment(doc);
        {
            newComment.setAuthor("VDeryushev");
            newComment.setInitial("VD");
            newComment.setDateTime(new Date());
        }

        newComment.setText("Comment regarding text.");

        // Add text to the document, warp it in a comment range, and then add your comment.
        Paragraph para = doc.getFirstSection().getBody().getFirstParagraph();
        para.appendChild(new CommentRangeStart(doc, newComment.getId()));
        para.appendChild(new Run(doc, "Commented text."));
        para.appendChild(new CommentRangeEnd(doc, newComment.getId()));
        para.appendChild(newComment); 
        
        // Add two replies to the comment.
        newComment.addReplyInternal("John Doe", "JD", new Date(), "New reply.");
        newComment.addReplyInternal("John Doe", "JD", new Date(), "Another reply.");

        printAllCommentInfo(doc.getChildNodes(NodeType.COMMENT, true));
    }
    
    /// <summary>
    /// Iterates over every top-level comment and prints its comment range, contents, and replies.
    /// </summary>
    private static void printAllCommentInfo(NodeCollection comments) throws Exception
    {
        CommentInfoPrinter commentVisitor = new CommentInfoPrinter();

        // Iterate over all top-level comments. Unlike reply-type comments, top-level comments have no ancestor.
        for (Comment comment : comments.Where(c => ((Comment)c).Ancestor == null) !!Autoporter error: Undefined expression type )
        {
            // First, visit the start of the comment range.
            CommentRangeStart commentRangeStart = (CommentRangeStart)comment.getPreviousSibling().getPreviousSibling().getPreviousSibling();
            commentRangeStart.accept(commentVisitor);

            // Then, visit the comment, and any replies that it may have.
            comment.accept(commentVisitor);

            for (Comment reply : (Iterable<Comment>) comment.getReplies())
                reply.accept(commentVisitor);

            // Finally, visit the end of the comment range, and then print the visitor's text contents.
            CommentRangeEnd commentRangeEnd = (CommentRangeEnd)comment.getPreviousSibling();
            commentRangeEnd.accept(commentVisitor);

            System.out.println(commentVisitor.getText());
        }
    }

    /// <summary>
    /// Prints information and contents of all comments and comment ranges encountered in the document.
    /// </summary>
    public static class CommentInfoPrinter extends DocumentVisitor
    {
        public CommentInfoPrinter()
        {
            mBuilder = new StringBuilder();
            mVisitorIsInsideComment = false;
        }

        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        public String getText()
        {
            return mBuilder.toString();
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitRun(Run run)
        {
            if (mVisitorIsInsideComment) indentAndAppendLine("[Run] \"" + run.getText() + "\"");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a CommentRangeStart node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitCommentRangeStart(CommentRangeStart commentRangeStart)
        {
            indentAndAppendLine("[Comment range start] ID: " + commentRangeStart.getId());
            mDocTraversalDepth++;
            mVisitorIsInsideComment = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a CommentRangeEnd node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitCommentRangeEnd(CommentRangeEnd commentRangeEnd)
        {
            mDocTraversalDepth--;
            indentAndAppendLine("[Comment range end] ID: " + commentRangeEnd.getId() + "\n");
            mVisitorIsInsideComment = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Comment node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitCommentStart(Comment comment)
        {
            indentAndAppendLine(
                $"[Comment start] For comment range ID {comment.Id}, By {comment.Author} on {comment.DateTime}");
            mDocTraversalDepth++;
            mVisitorIsInsideComment = true;


            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the visiting of a Comment node is ended in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitCommentEnd(Comment comment)
        {
            mDocTraversalDepth--;
            indentAndAppendLine("[Comment end]");
            mVisitorIsInsideComment = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++)
            {
                msStringBuilder.append(mBuilder, "|  ");
            }

            msStringBuilder.appendLine(mBuilder, text);
        }

        private boolean mVisitorIsInsideComment;
        private int mDocTraversalDepth;
        private /*final*/ StringBuilder mBuilder;
    }
    //ExEnd
}
