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
import com.aspose.words.Comment;
import com.aspose.ms.System.DateTime;
import com.aspose.words.NodeType;
import org.testng.Assert;
import com.aspose.words.NodeCollection;
import com.aspose.ms.System.msConsole;
import com.aspose.words.CommentCollection;
import com.aspose.words.Section;
import com.aspose.words.Body;
import com.aspose.words.Paragraph;
import com.aspose.words.CommentRangeStart;
import com.aspose.words.Run;
import com.aspose.words.CommentRangeEnd;
import java.util.ArrayList;
import com.aspose.ms.System.Collections.msArrayList;
import java.util.Iterator;
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
        //ExSummary:Shows how to add a comment with a reply to a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create new comment
        Comment newComment = new Comment(doc, "John Doe", "J.D.", DateTime.getNow());
        newComment.setText("My comment.");
        
        // Add this comment to a document node
        builder.getCurrentParagraph().appendChild(newComment);

        // Add comment reply
        newComment.addReplyInternal("John Doe", "JD", new DateTime(2017, 9, 25, 12, 15, 0), "New reply");
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        Comment docComment = (Comment)doc.getChild(NodeType.COMMENT, 0, true);

        Assert.assertEquals(1, docComment.getCount());
        Assert.assertEquals(1, newComment.getReplies().getCount());

        Assert.assertEquals("\u0005My comment.\r", docComment.getText());
        Assert.assertEquals("\u0005New reply\r", docComment.getReplies().get(0).getText());
    }

    @Test
    public void getAllCommentsAndReplies() throws Exception
    {
        //ExStart
        //ExFor:Comment.Ancestor
        //ExFor:Comment.Author
        //ExFor:Comment.Replies
        //ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
        //ExSummary:Shows how to get all comments with all replies.
        Document doc = new Document(getMyDir() + "Comments.docx");

        // Get all comment from the document
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
        Assert.assertEquals(12, comments.getCount()); //ExSkip

        // For all comments and replies we identify comment level and info about it
        for (Comment comment : comments.<Comment>OfType() !!Autoporter error: Undefined expression type )
        {
            if (comment.getAncestor() == null)
            {
                System.out.println("\nThis is a top-level comment");
                System.out.println("Comment author: " + comment.getAuthor());
                System.out.println("Comment text: " + comment.getText());

                for (Comment commentReply : comment.getReplies().<Comment>OfType() !!Autoporter error: Undefined expression type )
                {
                    System.out.println("\n\tThis is a comment reply");
                    System.out.println("\tReply author: " + commentReply.getAuthor());
                    System.out.println("\tReply text: " + commentReply.getText());
                }
            }
        }
        //ExEnd
    }

    @Test
    public void removeCommentReplies() throws Exception
    {
        //ExStart
        //ExFor:Comment.RemoveAllReplies
        //ExSummary:Shows how to remove comment replies.
        Document doc = new Document(getMyDir() + "Comments.docx");

        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
        Comment comment = (Comment)comments.get(0);
        Assert.AreEqual(2, comment.getReplies().Count()); //ExSkip

        comment.removeAllReplies();
        Assert.AreEqual(0, comment.getReplies().Count()); //ExSkip
        //ExEnd
    }

    @Test
    public void removeCommentReply() throws Exception
    {
        //ExStart
        //ExFor:Comment.RemoveReply(Comment)
        //ExFor:CommentCollection.Item(Int32)
        //ExSummary:Shows how to remove specific comment reply.
        Document doc = new Document(getMyDir() + "Comments.docx");

        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);

        Comment parentComment = (Comment)comments.get(0);
        CommentCollection repliesCollection = parentComment.getReplies();
        Assert.AreEqual(2, parentComment.getReplies().Count()); //ExSkip

        // Remove the first reply to comment
        parentComment.removeReply(repliesCollection.get(0));
        Assert.AreEqual(1, parentComment.getReplies().Count()); //ExSkip
        //ExEnd
    }

    @Test
    public void markCommentRepliesDone() throws Exception
    {
        //ExStart
        //ExFor:Comment.Done
        //ExFor:CommentCollection
        //ExSummary:Shows how to mark comment as Done.
        Document doc = new Document(getMyDir() + "Comments.docx");

        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);

        Comment comment = (Comment)comments.get(0);
        CommentCollection repliesCollection = comment.getReplies();

        for (Comment childComment : (Iterable<Comment>) repliesCollection)
        {
            if (!childComment.getDone())
            {
                // Update comment reply Done mark
                childComment.setDone(true);
            }
        }
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        comment = (Comment)doc.getChildNodes(NodeType.COMMENT, true).get(0);
        repliesCollection = comment.getReplies();

        for (Comment childComment : (Iterable<Comment>) repliesCollection)
        {
            Assert.assertTrue(childComment.getDone());
        }
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
    //ExSummary:Shows how to create comments with replies and get all interested info.
    @Test //ExSkip
    public void createCommentsAndPrintAllInfo() throws Exception
    {
        Document doc = new Document();
        doc.removeAllChildren();

        Section sect = (Section)doc.appendChild(new Section(doc));
        Body body = (Body)sect.appendChild(new Body(doc));

        // Create a commented text with several comment replies
        for (int i = 0; i <= 10; i++)
        {
            Comment newComment = createComment(doc, "VDeryushev", "VD", DateTime.getNow(), "My test comment " + i);

            Paragraph para = (Paragraph)body.appendChild(new Paragraph(doc));
            para.appendChild(new CommentRangeStart(doc, newComment.getId()));
            para.appendChild(new Run(doc, "Commented text " + i));
            para.appendChild(new CommentRangeEnd(doc, newComment.getId()));
            para.appendChild(newComment);
            
            for (int y = 0; y <= 2; y++)
            {
                newComment.addReplyInternal("John Doe", "JD", DateTime.getNow(), "New reply " + y);
            }
        }

        // Look at information of our comments
        printAllCommentInfo(extractComments(doc));
    }

    /// <summary>
    /// Create a new comment
    /// </summary>
    @Test (enabled = false)
    public static Comment createComment(Document doc, String author, String initials, DateTime dateTime, String text)
    {
        Comment newComment = new Comment(doc);
        {
            newComment.setAuthor(author); newComment.setInitial(initials); newComment.setDateTime(dateTime);
        }
        newComment.setText(text);

        return newComment;
    }

    /// <summary>
    /// Extract comments from the document without replies.
    /// </summary>
    @Test (enabled = false)
    public static ArrayList<Comment> extractComments(Document doc)
    {
        ArrayList<Comment> collectedComments = new ArrayList<Comment>();
        
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);

        for (Comment comment : (Iterable<Comment>) comments)
        {
            // All replies have ancestor, so we will add this check
            if (comment.getAncestor() == null)
            {
                msArrayList.add(collectedComments, comment);
            }
        }

        return collectedComments;
    }

    /// <summary>
    /// Use an iterator and a visitor to print info of every comment from within a document.
    /// </summary>
    private static void printAllCommentInfo(ArrayList<Comment> comments) throws Exception
    {
        // Create an object that inherits from the DocumentVisitor class
        CommentInfoPrinter commentVisitor = new CommentInfoPrinter();

        // Get the enumerator from the document's comment collection and iterate over the comments
        Iterator<Comment> enumerator = comments.iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
            {
                Comment currentComment = enumerator.next();

                // Accept our DocumentVisitor it to print information about our comments
                if (currentComment != null)
                {
                    // Get CommentRangeStart from our current comment and construct its information
                    CommentRangeStart commentRangeStart = (CommentRangeStart)currentComment.getPreviousSibling().getPreviousSibling().getPreviousSibling();
                    commentRangeStart.accept(commentVisitor);

                    // Construct current comment information
                    currentComment.accept(commentVisitor);
                    
                    // Get CommentRangeEnd from our current comment and construct its information
                    CommentRangeEnd commentRangeEnd = (CommentRangeEnd)currentComment.getPreviousSibling();
                    commentRangeEnd.accept(commentVisitor);
                }
            }

            // Output of all information received
            System.out.println(commentVisitor.getText());
        }
        finally { if (enumerator != null) enumerator.close(); }
    }

    /// <summary>
    /// This Visitor implementation prints information about and contents of comments and comment ranges encountered in the document.
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
