package com.aspose.words.examples.document_object_model;

import com.aspose.words.CompositeNode;
import com.aspose.words.Document;
import com.aspose.words.Node;
import com.aspose.words.examples.Utils;

public class SiblingNodes {
    //ExStart:
    public static void main(String[] args) throws Exception {
        String dataDir = Utils.getSharedDataDir(ChildNodes.class) + "DocumentObjectModel/";
        recurseAllNodes(dataDir);
    }

    public static void recurseAllNodes(String dataDir) throws Exception {
        // Open a document
        Document doc = new Document(dataDir + "Node.RecurseAllNodes.doc");
        // Invoke the recursive function that will walk the tree.
        traverseAllNodes(doc);
    }

    /**
     * A simple function that will walk through all children of a specified node
     * recursively and print the type of each node to the screen.
     */
    public static void traverseAllNodes(CompositeNode parentNode) throws Exception {
        // This is the most efficient way to loop through immediate children of a node.
        for (Node childNode = parentNode.getFirstChild(); childNode != null; childNode = childNode.getNextSibling()) {
            // Do some useful work.
            System.out.println(Node.nodeTypeToString(childNode.getNodeType()));

            // Recurse into the node if it is a composite node.
            if (childNode.isComposite())
                traverseAllNodes((CompositeNode) childNode);
        }
    }
    //ExEnd:
}
