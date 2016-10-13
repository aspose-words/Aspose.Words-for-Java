package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;

public class InsertField {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(InsertField.class);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField("MERGEFIELD MyFieldName \\* MERGEFORMAT");

        doc.save(dataDir + "output.docx");


    }
}




