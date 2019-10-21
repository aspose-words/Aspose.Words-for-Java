package com.aspose.words.examples.programming_documents.comments;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.util.ArrayList;

@SuppressWarnings("unchecked")
public class ProcessComments {
    public static void main(String[] args) throws Exception {

        //ExStart:ProcessComments
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ProcessComments.class);

        // Open the document.
        Document doc = new Document(dataDir + "TestFile.doc");

        for (String comment : (Iterable<String>) extractComments(doc))
            System.out.print(comment);

        // Remove comments by the "pm" author.
        removeComments(doc, "pm");
        System.out.println("Comments from \"pm\" are removed!");

        // Extract the information about the comments of the "ks" author.
        for (String comment : (Iterable<String>) extractComments(doc, "ks"))
            System.out.print(comment);

        //Read the comment's reply and resolve them.
        System.out.println("Read the comment's reply and resolve them.");
        CommentResolvedandReplies(doc);

        // Remove all comments.
        removeComments(doc);
        System.out.println("All comments are removed!");

        // Save the document.
        doc.save(dataDir + "output.doc");
        //ExEnd:ProcessComments

    }

    //ExFor:Comment.Author
    //ExFor:Comment.DateTime
    //ExId:ProcessComments_Extract_All
    //ExSummary:Extracts the author name, date&time and text of all comments in the document.
    static ArrayList extractComments(Document doc) throws Exception {
        //ExStart:extractComments
        ArrayList collectedComments = new ArrayList();
        // Collect all comments in the document
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
        // Look through all comments and gather information about them.
        for (Comment comment : (Iterable<Comment>) comments) {
            collectedComments.add(comment.getAuthor() + " " + comment.getDateTime() + " " + comment.toString(SaveFormat.TEXT));
        }
        return collectedComments;
        //ExEnd:extractComments
    }

    //ExId:ProcessComments_Extract_Author
    //ExSummary:Extracts the author name, date&time and text of the comments by the specified author.
    static ArrayList extractComments(Document doc, String authorName) throws Exception {
        //ExStart:extractComments_Author
        ArrayList collectedComments = new ArrayList();
        // Collect all comments in the document
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
        // Look through all comments and gather information about those written by the authorName author.
        for (Comment comment : (Iterable<Comment>) comments) {
            if (comment.getAuthor().equals(authorName))
                collectedComments.add(comment.getAuthor() + " " + comment.getDateTime() + " " + comment.toString(SaveFormat.TEXT));
        }
        return collectedComments;
        //ExEnd:extractComments_Author
    }

    //ExId:ProcessComments_Remove_All
    //ExSummary:Removes all comments in the document.
    static void removeComments(Document doc) throws Exception {
        //ExStart:removeComments
        // Collect all comments in the document
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
        // Remove all comments.
        comments.clear();
        //ExEnd:removeComments
    }

    //ExId:ProcessComments_Remove_Author
    //ExSummary:Removes comments by the specified author.
    static void removeComments(Document doc, String authorName) throws Exception {
        //ExStart:removeComments_Author
        // Collect all comments in the document
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
        // Look through all comments and remove those written by the authorName author.
        for (int i = comments.getCount() - 1; i >= 0; i--) {
            Comment comment = (Comment) comments.get(i);
            if (comment.getAuthor().equals(authorName))
                comment.remove();
        }

    }
    //ExEnd:removeComments_Author

    // ExStart:CommentResolvedandReplies
    static void CommentResolvedandReplies(Document doc) {
        NodeCollection<Comment> comments = doc.getChildNodes(NodeType.COMMENT, true);
        Comment parentComment = (Comment) comments.get(0);

        for (Comment childComment : parentComment.getReplies()) {
            // Get comment parent and status.
            System.out.println(childComment.getAncestor().getId());
            System.out.println(childComment.getDone());

            // And update comment Done mark.
            childComment.setDone(true);
        }
    }
    // ExEnd:CommentResolvedandReplies
}