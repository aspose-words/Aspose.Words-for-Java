package com.aspose.words.examples.programming_documents.comments;

import com.aspose.words.Document;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.examples.Utils;

@SuppressWarnings("unchecked")
public class RemoveComments {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RemoveComments.class);

        // Open the document.
        Document doc = new Document(dataDir + "TestFile.doc");
        // Collect all comments in the document
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
        // Remove all comments.
        comments.clear();
        doc.save(dataDir + "output.doc");

    }
}