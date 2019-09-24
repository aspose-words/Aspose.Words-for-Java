package com.aspose.words.examples.programming_documents.comments;

import com.aspose.words.Comment;
import com.aspose.words.Document;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.examples.Utils;

@SuppressWarnings("unchecked")
public class RemoveCommentsByAuthor {
    public static void main(String[] args) throws Exception {

        //ExStart:RemoveCommentsByAuthor
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RemoveCommentsByAuthor.class);

        // Open the document.
        Document doc = new Document(dataDir + "TestFile.doc");
        String authorName = "pm";
        // Collect all comments in the document
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
        // Look through all comments and remove those written by the authorName author.
        for (int i = comments.getCount() - 1; i >= 0; i--) {
            Comment comment = (Comment) comments.get(i);
            if (comment.getAuthor().equals(authorName))
                comment.remove();
        }
        doc.save(dataDir + "output.doc");
        //ExEnd:RemoveCommentsByAuthor

    }
}