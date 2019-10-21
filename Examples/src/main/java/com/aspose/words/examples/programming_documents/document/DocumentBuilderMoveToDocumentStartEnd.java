package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;


public class DocumentBuilderMoveToDocumentStartEnd {
    public static void main(String[] args) throws Exception {

        //ExStart:DocumentBuilderMoveToDocumentStartEnd
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderMoveToDocumentStartEnd.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToDocumentEnd();
        builder.write("\n\nThis is the end of the document.");

        builder.insertParagraph();
        builder.moveToDocumentStart();
        builder.write("\nThis is the beginning of the document.");
        doc.save(dataDir + "output.doc");
        //ExEnd:DocumentBuilderMoveToDocumentStartEnd

    }
}