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

import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;

public class ExComment extends ApiExampleBase {
    @Test
    public void addCommentWithReply() throws Exception {
        //ExStart
        //ExFor:Comment
        //ExFor:Comment.SetText(String)
        //ExFor:Comment.AddReply(String, String, DateTime, String)
        //ExSummary:Shows how to add a comment with a reply to a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create new comment
        Comment newComment = new Comment(doc, "John Doe", "J.D.", new Date(System.currentTimeMillis()));
        newComment.setText("My comment.");

        // Add this comment to a document node
        builder.getCurrentParagraph().appendChild(newComment);

        Calendar cal = Calendar.getInstance();
        cal.set(2017, Calendar.SEPTEMBER, 25, 12, 15, 0);
        cal.getTime();

        // Add comment reply
        newComment.addReply("John Doe", "JD", cal.getTime(), "New reply");
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        Comment docComment = (Comment) doc.getChild(NodeType.COMMENT, 0, true);

        Assert.assertEquals(1, docComment.getCount());
        Assert.assertEquals(1, newComment.getReplies().getCount());

        Assert.assertEquals("\u0005My comment.\r", docComment.getText());
        Assert.assertEquals("\u0005New reply\r", docComment.getReplies().get(0).getText());
    }

    @Test
    public void getAllCommentsAndReplies() throws Exception {
        //ExStart
        //ExFor:Comment.Ancestor
        //ExFor:Comment.Author
        //ExFor:Comment.Replies
        //ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
        //ExSummary:Shows how to get all comments with all replies.
        Document doc = new Document(getMyDir() + "Comments.docx");

        // Get all comment from the document
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);

        Assert.assertEquals(comments.getCount(), 12); //ExSkip

        // For all comments and replies we identify comment level and info about it
        for (Comment comment : (Iterable<Comment>) comments) {
            if (comment.getAncestor() == null) {
                System.out.println("\nThis is a top-level comment");
                System.out.println("Comment author: " + comment.getAuthor());
                System.out.println("Comment text: " + comment.getText());

                for (Comment commentReply : comment.getReplies()) {
                    System.out.println("\n\tThis is a comment reply");
                    System.out.println("\tReply author: " + commentReply.getAuthor());
                    System.out.println("\tReply text: " + commentReply.getText());
                }
            }
        }
        //ExEnd
    }

    @Test
    public void removeCommentReplies() throws Exception {
        //ExStart
        //ExFor:Comment.RemoveAllReplies
        //ExSummary:Shows how to remove comment replies.
        Document doc = new Document(getMyDir() + "Comments.docx");

        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
        Comment comment = (Comment) comments.get(0);
        Assert.assertEquals(2, comment.getReplies().getCount()); //ExSkip

        comment.removeAllReplies();
        Assert.assertEquals(0, comment.getReplies().getCount()); //ExSkip
        //ExEnd
    }

    @Test
    public void removeCommentReply() throws Exception {
        //ExStart
        //ExFor:Comment.RemoveReply(Comment)
        //ExFor:CommentCollection.Item(Int32)
        //ExSummary:Shows how to remove specific comment reply.
        Document doc = new Document(getMyDir() + "Comments.docx");

        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);

        Comment parentComment = (Comment) comments.get(0);
        CommentCollection repliesCollection = parentComment.getReplies();
        Assert.assertEquals(2, parentComment.getReplies().getCount()); //ExSkip

        // Remove the first reply to comment
        parentComment.removeReply(repliesCollection.get(0));
        Assert.assertEquals(1, parentComment.getReplies().getCount()); //ExSkip
        //ExEnd
    }

    @Test
    public void markCommentRepliesDone() throws Exception {
        //ExStart
        //ExFor:Comment.Done
        //ExFor:CommentCollection
        //ExSummary:Shows how to mark comment as Done.
        Document doc = new Document(getMyDir() + "Comments.docx");

        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);

        Comment comment = (Comment) comments.get(0);
        CommentCollection repliesCollection = comment.getReplies();

        for (Comment childComment : repliesCollection) {
            if (!childComment.getDone()) {
                // Update comment reply Done mark
                childComment.setDone(true);
            }
        }
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        comment = (Comment) doc.getChildNodes(NodeType.COMMENT, true).get(0);
        repliesCollection = comment.getReplies();

        for (Comment childComment : (Iterable<Comment>) repliesCollection) {
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
    public void createCommentsAndPrintAllInfo() throws Exception {
        Document doc = new Document();
        doc.removeAllChildren();

        Section sect = (Section) doc.appendChild(new Section(doc));
        Body body = (Body) sect.appendChild(new Body(doc));

        // Create a commented text with several comment replies
        for (int i = 0; i <= 10; i++) {
            Comment newComment = createComment(doc, "VDeryushev", "VD", new Date(), "My test comment " + i);

            Paragraph para = (Paragraph) body.appendChild(new Paragraph(doc));
            para.appendChild(new CommentRangeStart(doc, newComment.getId()));
            para.appendChild(new Run(doc, "Commented text " + i));
            para.appendChild(new CommentRangeEnd(doc, newComment.getId()));
            para.appendChild(newComment);

            for (int y = 0; y <= 2; y++) {
                newComment.addReply("John Doe", "JD", new Date(), "New reply " + y);
            }
        }

        // Look at information of our comments
        printAllCommentInfo(extractComments(doc));
    }

    /// <summary>
    /// Create a new comment.
    /// </summary>
    public static Comment createComment(Document doc, String author, String initials, Date dateTime, String text) {
        Comment newComment = new Comment(doc);

        newComment.setAuthor(author);
        newComment.setInitial(initials);
        newComment.setDateTime(dateTime);
        newComment.setText(text);

        return newComment;
    }

    /// <summary>
    /// Extract comments from the document without replies.
    /// </summary>
    public static ArrayList<Comment> extractComments(Document doc) {
        ArrayList<Comment> collectedComments = new ArrayList<>();

        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);

        for (Comment comment : (Iterable<Comment>) comments) {
            // All replies have ancestor, so we will add this check
            if (comment.getAncestor() == null) {
                collectedComments.add(comment);
            }
        }

        return collectedComments;
    }

    /// <summary>
    /// Use an iterator and a visitor to print info of every comment from within a document.
    /// </summary>
    private static void printAllCommentInfo(ArrayList<Comment> comments) throws Exception {
        // Create an object that inherits from the DocumentVisitor class
        CommentInfoPrinter commentVisitor = new CommentInfoPrinter();

        // Get the enumerator from the document's comment collection and iterate over the comments
        Iterator<Comment> enumerator = comments.iterator();

        while (enumerator.hasNext()) {
            Comment currentComment = enumerator.next();

            // Accept our DocumentVisitor it to print information about our comments
            if (currentComment != null) {
                // Get CommentRangeStart from our current comment and construct its information
                CommentRangeStart commentRangeStart = (CommentRangeStart) currentComment.getPreviousSibling().getPreviousSibling().getPreviousSibling();
                commentRangeStart.accept(commentVisitor);

                // Construct current comment information
                currentComment.accept(commentVisitor);

                // Get CommentRangeEnd from our current comment and construct its information
                CommentRangeEnd commentRangeEnd = (CommentRangeEnd) currentComment.getPreviousSibling();
                commentRangeEnd.accept(commentVisitor);
            }
        }

        // Output of all information received
        System.out.println(commentVisitor.getText());
    }

    /// <summary>
    /// This Visitor implementation prints information about and contents of comments and comment ranges encountered in the document.
    /// </summary>
    public static class CommentInfoPrinter extends DocumentVisitor {
        public CommentInfoPrinter() {
            mBuilder = new StringBuilder();
            mVisitorIsInsideComment = false;
        }

        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        public String getText() {
            return mBuilder.toString();
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        public int visitRun(Run run) {
            if (mVisitorIsInsideComment) indentAndAppendLine("[Run] \"" + run.getText() + "\"");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a CommentRangeStart node is encountered in the document.
        /// </summary>
        public int visitCommentRangeStart(CommentRangeStart commentRangeStart) {
            indentAndAppendLine("[Comment range start] ID: " + commentRangeStart.getId());
            mDocTraversalDepth++;
            mVisitorIsInsideComment = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a CommentRangeEnd node is encountered in the document.
        /// </summary>
        public int visitCommentRangeEnd(CommentRangeEnd commentRangeEnd) {
            mDocTraversalDepth--;
            indentAndAppendLine("[Comment range end] ID: " + commentRangeEnd.getId() + "\n");
            mVisitorIsInsideComment = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Comment node is encountered in the document.
        /// </summary>
        public int visitCommentStart(Comment comment) {
            indentAndAppendLine(MessageFormat.format("[Comment start] For comment range ID {0}, By {1} on {2}", comment.getId(),
                    comment.getAuthor(), comment.getDateTime()));
            mDocTraversalDepth++;
            mVisitorIsInsideComment = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the visiting of a Comment node is ended in the document.
        /// </summary>
        public int visitCommentEnd(Comment comment) {
            mDocTraversalDepth--;
            indentAndAppendLine("[Comment end]");
            mVisitorIsInsideComment = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(String text) {
            for (int i = 0; i < mDocTraversalDepth; i++) {
                mBuilder.append("|  ");
            }

            mBuilder.append(text + "\r\n");
        }

        private boolean mVisitorIsInsideComment;
        private int mDocTraversalDepth;
        private StringBuilder mBuilder;
    }
    //ExEnd
}
