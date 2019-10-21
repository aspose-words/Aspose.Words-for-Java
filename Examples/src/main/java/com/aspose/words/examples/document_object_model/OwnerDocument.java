package com.aspose.words.examples.document_object_model;

import com.aspose.words.Document;
import com.aspose.words.Paragraph;

public class OwnerDocument {

    public static void main(String[] args) throws Exception {
        //ExStart:
        // Open a file from disk.
        Document doc = new Document();

        // Creating a new node of any type requires a document passed into the constructor.
        Paragraph para = new Paragraph(doc);

        // The new paragraph node does not yet have a parent.
        System.out.println("Paragraph has no parent node: " + (para.getParentNode() == null));

        // But the paragraph node knows its document.
        System.out.println("Both nodes' documents are the same: " + (para.getDocument() == doc));

        // The fact that a node always belongs to a document allows us to access and modify
        // properties that reference the document-wide data such as styles or lists.
        para.getParagraphFormat().setStyleName("Heading 1");

        // Now add the paragraph to the main text of the first section.
        doc.getFirstSection().getBody().appendChild(para);

        // The paragraph node is now a child of the Body node.
        System.out.println("Paragraph has a parent node: " + (para.getParentNode() != null));
        //ExEnd:
    }

}
