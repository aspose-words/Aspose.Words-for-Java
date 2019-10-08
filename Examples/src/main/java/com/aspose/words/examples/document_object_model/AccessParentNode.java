package com.aspose.words.examples.document_object_model;

import com.aspose.words.Document;
import com.aspose.words.Node;

public class AccessParentNode {

    public static void main(String[] args) throws Exception {
        //ExStart:
        // Create a new empty document. It has one section.
        Document doc = new Document();
        // The section is the first child node of the document.
        Node section = doc.getFirstChild();
        // The section's parent node is the document.
        System.out.println("Section parent is the document: " + (doc == section.getParentNode()));
        //ExEnd:
    }
}
