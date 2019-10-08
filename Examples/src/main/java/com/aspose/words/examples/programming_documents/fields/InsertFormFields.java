package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;

public class InsertFormFields {
    public static void main(String[] args) throws Exception {

        //ExStart:InsertFormFields
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(InsertFormFields.class);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        String[] items = {"One", "Two", "Three"};
        builder.insertComboBox("DropDown", items, 0);
        doc.save(dataDir + "output.docx");
        //ExEnd:InsertFormFields


    }
}




