package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class GetFieldNames {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(GetFieldNames.class);

        Document doc = new Document(dataDir + "Rendering.doc");
        String[] fieldNames = doc.getMailMerge().getFieldNames();
        System.out.println("\nDocument have " + fieldNames.length + " fields.");
        for (String name : fieldNames) {
            System.out.println(name);
        }


    }
}




