package com.aspose.words.examples.programming_documents.comments;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.util.ArrayList;

@SuppressWarnings("unchecked")
public class ExtractCommentsByAuthor {
    public static void main(String[] args) throws Exception {

        //ExStart:ExtractCommentsByAuthor
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ExtractCommentsByAuthor.class);

        String authorName = "ks";
        // Open the document.
        Document doc = new Document(dataDir + "TestFile.doc");
        ArrayList collectedComments = new ArrayList();
        // Collect all comments in the document
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
        // Look through all comments and gather information about those written by the authorName author.
        for (Comment comment : (Iterable<Comment>) comments) {
            if (comment.getAuthor().equals(authorName))
                collectedComments.add(comment.getAuthor() + " " + comment.getDateTime() + " " + comment.toString(SaveFormat.TEXT));
        }
        System.out.print(collectedComments);
        //ExEnd:ExtractCommentsByAuthor

    }
}