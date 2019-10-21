package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;


public class DocumentBuilderInsertInlineImage {
    public static void main(String[] args) throws Exception {

        //ExStart:DocumentBuilderInsertInlineImage
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderInsertInlineImage.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertImage(dataDir + "test.jpg");

        doc.save(dataDir + "output.doc");
        //ExEnd:DocumentBuilderInsertInlineImage

    }
}