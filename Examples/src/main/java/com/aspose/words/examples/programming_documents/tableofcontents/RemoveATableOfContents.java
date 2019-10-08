package com.aspose.words.examples.programming_documents.tableofcontents;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.util.ArrayList;

public class RemoveATableOfContents {

    private static final String dataDir = Utils.getSharedDataDir(RemoveATableOfContents.class) + "TableOfContents/";

    public static void main(String[] args) throws Exception {

        //ExStart:RemoveATableOfContents
        // Open a document which contains a TOC.
        Document doc = new Document(dataDir + "Document.TableOfContents.doc");

        // Remove the first table of contents from the document.
        removeTableOfContents(doc, 0);

        // Save the output.
        doc.save(dataDir + "Document.TableOfContentsRemoveToc_Out.doc");
        //ExEnd:RemoveATableOfContents

    }

    //ExStart:removeTableOfContents

    /**
     * Removes the specified table of contents field from the document.
     *
     * @param doc   The document to remove the field from.
     * @param index The zero-based index of the TOC to remove.
     */
    public static void removeTableOfContents(Document doc, int index) throws Exception {
        // Store the FieldStart nodes of TOC fields in the document for quick access.
        ArrayList<FieldStart> fieldStarts = new ArrayList<FieldStart>();
        // This is a list to store the nodes found inside the specified TOC. They will be removed
        // at the end of this method.
        ArrayList<Node> nodeList = new ArrayList<Node>();

        for (FieldStart start : (Iterable<FieldStart>) doc.getChildNodes(NodeType.FIELD_START, true)) {
            if (start.getFieldType() == FieldType.FIELD_TOC) {
                // Add all FieldStarts which are of type FieldTOC.
                fieldStarts.add(start);
            }
        }

        // Ensure the TOC specified by the passed index exists.
        if (index > fieldStarts.size() - 1)
            throw new ArrayIndexOutOfBoundsException("TOC index is out of range");

        boolean isRemoving = true;
        // Get the FieldStart of the specified TOC.
        Node currentNode = fieldStarts.get(index);

        while (isRemoving) {
            // It is safer to store these nodes and delete them all at once later.
            nodeList.add(currentNode);
            currentNode = currentNode.nextPreOrder(doc);

            // Once we encounter a FieldEnd node of type FieldTOC then we know we are at the end
            // of the current TOC and we can stop here.
            if (currentNode.getNodeType() == NodeType.FIELD_END) {
                FieldEnd fieldEnd = (FieldEnd) currentNode;
                if (fieldEnd.getFieldType() == FieldType.FIELD_TOC)
                    isRemoving = false;
            }
        }

        // Remove all nodes found in the specified TOC.
        for (Node node : nodeList) {
            node.remove();
        }
    }
    //ExEnd:removeTableOfContents
}