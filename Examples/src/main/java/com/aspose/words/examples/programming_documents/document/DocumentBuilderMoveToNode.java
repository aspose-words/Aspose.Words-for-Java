package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;


public class DocumentBuilderMoveToNode {
    public static void main(String[] args) throws Exception {

        //ExStart:DocumentBuilderMoveToNode
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderMoveToNode.class);

        // Open the document.
        Document doc = new Document(dataDir + "DocumentBuilder.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveTo(doc.getFirstSection().getBody().getLastParagraph());
        doc.save(dataDir + "output.doc");
        //ExEnd:DocumentBuilderMoveToNode

    }
}