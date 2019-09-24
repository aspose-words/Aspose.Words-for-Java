package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;


public class DocumentBuilderMoveToMergeField {
    public static void main(String[] args) throws Exception {

        //ExStart:DocumentBuilderMoveToMergeField
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderMoveToMergeField.class);

        // Open the document.
        Document doc = new Document(dataDir + "DocumentBuilder.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToMergeField("NiceMergeField");
        builder.writeln("This is a very nice merge field.");
        // doc.save(dataDir + "output.doc");
        doc.save(dataDir + "output.doc");
        //ExEnd:DocumentBuilderMoveToMergeField

    }
}