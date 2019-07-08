package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.ByteArrayOutputStream;
import java.text.MessageFormat;
import java.util.Calendar;
import java.util.Date;

public class ExComment extends ApiExampleBase {
    @Test
    public void addCommentWithReply() throws Exception {
        //ExStart
        //ExFor:Comment
        //ExFor:Comment.SetText(String)
        //ExFor:Comment.Replies
        //ExFor:Comment.AddReply(String, String, DateTime, String)
        //ExSummary:Shows how to add a comment with a reply to a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        //Create new comment
        Comment newComment = new Comment(doc, "John Doe", "J.D.", new Date(System.currentTimeMillis()));
        newComment.setText("My comment.");

        //Add this comment to a document node
        builder.getCurrentParagraph().appendChild(newComment);

        Calendar cal = Calendar.getInstance();
        cal.set(2017, Calendar.SEPTEMBER, 25, 12, 15, 0);
        cal.getTime();

        //Add comment reply
        newComment.addReply("John Doe", "JD", cal.getTime(), "New reply");
        //ExEnd

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        Comment docComment = (Comment) doc.getChild(NodeType.COMMENT, 0, true);

        Assert.assertEquals(docComment.getCount(), 1);
        Assert.assertEquals(newComment.getReplies().getCount(), 1);

        Assert.assertEquals(docComment.getText(), "\u0005My comment.\r");
        Assert.assertEquals(docComment.getReplies().get(0).getText(), "\u0005New reply\r");
    }

    @Test
    public void getAllCommentsAndReplies() throws Exception {
        //ExStart
        //ExFor:Comment.Ancestor
        //ExFor:Comment.Author
        //ExSummary:Shows how to get all comments with all replies.
        Document doc = new Document(getMyDir() + "Comment.Document.docx");

        //Get all comment from the document
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);

        Assert.assertEquals(comments.getCount(), 12); //ExSkip

        //For all comments and replies we identify comment level and info about it
        for (Comment comment : (Iterable<Comment>) comments) {
            if (comment.getAncestor() == null) {
                System.out.println("This is a top-level comment\n");

                System.out.println(MessageFormat.format("Comment author: ", comment.getAuthor()));
                System.out.println("Comment text: " + comment.getText());

                for (Comment commentReply : comment.getReplies()) {
                    System.out.println("This is a comment reply\n");

                    System.out.println(MessageFormat.format("Comment author: ", commentReply.getAuthor()));
                    System.out.println(MessageFormat.format("Comment text: ", commentReply.getText()));
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
        Document doc = new Document(getMyDir() + "Comment.Document.docx");

        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
        Comment comment = (Comment) comments.get(0);

        comment.removeAllReplies();
        //ExEnd
    }

    @Test
    public void removeCommentReply() throws Exception {
        //ExStart
        //ExFor:Comment.RemoveReply(Comment)
        //ExSummary:Shows how to remove specific comment reply.
        Document doc = new Document(getMyDir() + "Comment.Document.docx");

        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);

        Comment parentComment = (Comment) comments.get(0);

        // Remove the first reply to comment
        parentComment.removeReply(parentComment.getReplies().get(0));
        //ExEnd
    }

    @Test
    public void markCommentRepliesDone() throws Exception {
        //ExStart
        //ExFor:Comment.Done
        //ExSummary:Shows how to mark comment as Done.
        Document doc = new Document(getMyDir() + "Comment.Document.docx");

        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);

        Comment comment = (Comment) comments.get(0);

        for (Comment childComment : comment.getReplies()) {
            if (!childComment.getDone()) {
                // Update comment reply Done mark.
                childComment.setDone(true);
            }
        }
        //ExEnd
    }
}
