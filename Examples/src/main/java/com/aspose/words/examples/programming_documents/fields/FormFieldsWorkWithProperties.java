package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class FormFieldsWorkWithProperties {
    public static void main(String[] args) throws Exception {

        //ExStart:FormFieldsWorkWithProperties
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(FormFieldsWorkWithProperties.class);

        Document doc = new Document(dataDir + "FormFields.doc");

        DocumentBuilder builder = new DocumentBuilder(doc);
        FormFieldCollection documentFormFields = doc.getRange().getFormFields();

        FormField formField = doc.getRange().getFormFields().get(3);
        if (formField.getType() == FieldType.FIELD_FORM_TEXT_INPUT)
            formField.setResult("Field Name :" + formField.getName());

        doc.save(dataDir + "output.docx");
        //ExEnd:FormFieldsWorkWithProperties

    }
}




