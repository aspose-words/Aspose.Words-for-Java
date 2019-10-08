package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.Document;
import com.aspose.words.FormFieldCollection;
import com.aspose.words.examples.Utils;

public class FormFieldsGetFormFieldsCollection {
    public static void main(String[] args) throws Exception {

        //ExStart:FormFieldsGetFormFieldsCollection
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(FormFieldsGetFormFieldsCollection.class);

        Document doc = new Document(dataDir + "FormFields.doc");
        FormFieldCollection formFields = doc.getRange().getFormFields();
        doc.save(dataDir + "output.docx");
        //ExEnd:FormFieldsGetFormFieldsCollection

    }
}




