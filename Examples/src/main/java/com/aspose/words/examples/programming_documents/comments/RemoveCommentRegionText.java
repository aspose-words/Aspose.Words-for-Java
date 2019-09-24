package com.aspose.words.examples.programming_documents.comments;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

@SuppressWarnings("unchecked")
public class RemoveCommentRegionText {
    public static void main(String[] args) throws Exception {

        //ExStart:RemoveCommentRegionText
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RemoveCommentRegionText.class);

        // Open the document.
        Document doc = new Document(dataDir + "TestFile.doc");
        CommentRangeStart commentStart = (CommentRangeStart) doc.getChild(NodeType.COMMENT_RANGE_START, 0, true);
        CommentRangeEnd commentEnd = (CommentRangeEnd) doc.getChild(NodeType.COMMENT_RANGE_END, 0, true);

        Node currentNode = commentStart;
        Boolean isRemoving = true;
        while (currentNode != null && isRemoving) {
            if (currentNode.getNodeType() == NodeType.COMMENT_RANGE_END)
                isRemoving = false;

            Node nextNode = currentNode.nextPreOrder(doc);
            currentNode.remove();
            currentNode = nextNode;
        }

        doc.save(dataDir + "output.doc");
        //ExEnd:RemoveCommentRegionText

    }
}