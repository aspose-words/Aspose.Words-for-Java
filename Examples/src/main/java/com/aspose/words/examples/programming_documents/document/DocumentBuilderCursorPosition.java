package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Node;
import com.aspose.words.Paragraph;
import com.aspose.words.examples.Utils;


public class DocumentBuilderCursorPosition {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderCursorPosition.class);

        // Open the document.
        Document doc = new Document(dataDir + "DocumentBuilder.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);
        Node node = builder.getCurrentNode();
        Paragraph curParagraph = builder.getCurrentParagraph();
        doc.save(dataDir + "output.doc");

    }
}