package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;


public class DocumentBuilderMoveToDocumentStartEnd {
    public static void main(String[] args) throws Exception {

        //ExStart:DocumentBuilderMoveToDocumentStartEnd
        String dataDir = Utils.getDataDir(DocumentBuilderMoveToDocumentStartEnd.class);

        Document doc = new Document(dataDir + "Document.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor position to the beginning of your document.
        builder.moveToDocumentStart();
        builder.writeln("This is the beginning of the document.");

        // Move the cursor position to the end of your document.
        builder.moveToDocumentEnd();
        builder.writeln("This is the end of the document.");
        //ExEnd:DocumentBuilderMoveToDocumentStartEnd

    }
}