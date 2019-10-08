package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FormField;
import com.aspose.words.FormFieldCollection;
import com.aspose.words.examples.Utils;

public class FormFieldsGetByName {
    public static void main(String[] args) throws Exception {

        //ExStart:FormFieldsGetByName
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(FormFieldsGetByName.class);

        Document doc = new Document(dataDir + "FormFields.doc");

        DocumentBuilder builder = new DocumentBuilder(doc);
        // FormFieldCollection formFields = doc.getRange().getFormFields();
        FormFieldCollection documentFormFields = doc.getRange().getFormFields();

        FormField formField1 = documentFormFields.get(3);
        FormField formField2 = documentFormFields.get("Text2");
        System.out.println("Name: " + formField2.getName());
        doc.save(dataDir + "output.docx");
        //ExEnd:FormFieldsGetByName

    }
}




