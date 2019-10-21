package com.aspose.words.examples.programming_documents.comments;

import com.aspose.words.Comment;
import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.examples.Utils;

import java.util.Date;


/**
 * Created by Home on 10/13/2017.
 */
public class AddRemoveCommentReply {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AddRemoveCommentReply.class);
        // ExStart:AddRemoveCommentReply
        Document doc = new Document(dataDir + "TestFile.doc");
        Comment comment = (Comment) doc.getChild(NodeType.COMMENT, 0, true);

        //Remove the reply
        comment.removeReply(comment.getReplies().get(0));

        //Add a reply to comment
        comment.addReply("John Doe", "JD", new Date(), "New reply");

        dataDir = dataDir + "TestFile_Out.doc";

        // Save the document to disk.
        doc.save(dataDir);
        // ExEnd:AddRemoveCommentReply
        System.out.println("\nComment's reply is removed successfully.\nFile saved at " + dataDir);
    }
}
