package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;


public class DocumentBuilderInsertField {
    public static void main(String[] args) throws Exception {

        //ExStart:DocumentBuilderInsertField
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderInsertField.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getFont().setLocaleId(1031);
        builder.insertField("MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"");
        builder.write(" - ");
        builder.insertField("MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"");

        doc.save(dataDir + "output.doc");
        //ExEnd:DocumentBuilderInsertField

    }
}