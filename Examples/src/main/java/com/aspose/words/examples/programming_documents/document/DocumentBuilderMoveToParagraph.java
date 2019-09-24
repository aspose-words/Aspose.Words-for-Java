package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;


public class DocumentBuilderMoveToParagraph {
    public static void main(String[] args) throws Exception {

        //ExStart:DocumentBuilderMoveToParagraph
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderMoveToParagraph.class);

        // Open the document.
        Document doc = new Document(dataDir + "DocumentBuilder.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToParagraph(2, 0);
        builder.writeln("This is the 3rd paragraph.");

        doc.save(dataDir + "output.doc");
        //ExEnd:DocumentBuilderMoveToParagraph

    }
}