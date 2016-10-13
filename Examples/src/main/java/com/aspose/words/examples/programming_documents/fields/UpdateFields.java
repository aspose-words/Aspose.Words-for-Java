package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class UpdateFields {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(UpdateFields.class);

        Document doc = new Document(dataDir + "in.doc");


        // update fields
        doc.updateFields();
        doc.save(dataDir + "output.docx");


    }
}




