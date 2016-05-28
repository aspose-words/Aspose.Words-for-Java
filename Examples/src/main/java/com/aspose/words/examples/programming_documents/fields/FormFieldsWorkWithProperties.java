/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class FormFieldsWorkWithProperties
{
    public static void main(String[] args) throws Exception
    {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(FormFieldsWorkWithProperties.class);

        Document doc = new Document(dataDir + "FormFields.doc");

        DocumentBuilder builder = new DocumentBuilder(doc);
        FormFieldCollection documentFormFields = doc.getRange().getFormFields();

        FormField formField = doc.getRange().getFormFields().get(3);
        if(formField.getType() == FieldType.FIELD_FORM_TEXT_INPUT)

            formField.setResult("Field Name :" + formField.getName());

        doc.save(dataDir + "Output.docx");
        //ExEnd:1
   }
}




