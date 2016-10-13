package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;


public class DocumentBuilderInsertCheckBoxFormField {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderInsertCheckBoxFormField.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertCheckBox("CheckBox", true, true, 0);
        doc.save(dataDir + "output.doc");

    }
}