package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;


public class DocumentBuilderInsertComboBoxFormField {
    public static void main(String[] args) throws Exception {

        //ExStart:DocumentBuilderInsertComboBoxFormField
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderInsertComboBoxFormField.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        String[] items = {"One", "Two", "Three"};
        builder.insertComboBox("DropDown", items, 0);
        doc.save(dataDir + "output.doc");
        //ExEnd:DocumentBuilderInsertComboBoxFormField

    }
}