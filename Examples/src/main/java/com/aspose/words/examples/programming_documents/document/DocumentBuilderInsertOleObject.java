package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;


public class DocumentBuilderInsertOleObject {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderInsertOleObject.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertOleObject("http://www.aspose.com", "htmlfile", true, true, null);
        doc.save(dataDir + "output.doc");

    }
}