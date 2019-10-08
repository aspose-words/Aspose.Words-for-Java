package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class UpdateDocFields {
    public static void main(String[] args) throws Exception {

        //ExStart:UpdateDocFields
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(UpdateDocFields.class);

        Document doc = new Document(dataDir + "Rendering.doc");

        doc.updateFields();
        doc.save(dataDir + "output.docx");
        //ExEnd:UpdateDocFields

    }
}




