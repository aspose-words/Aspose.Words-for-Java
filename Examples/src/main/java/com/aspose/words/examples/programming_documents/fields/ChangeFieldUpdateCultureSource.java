/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FieldUpdateCultureSource;
import com.aspose.words.examples.Utils;

public class ChangeFieldUpdateCultureSource {
    public static void main(String[] args) throws Exception {
        //TODO

        //ExStart:ChangeFieldUpdateCultureSource
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ChangeFieldUpdateCultureSource.class);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);


        builder.getFont().setLocaleId(1031);
        builder.insertField("MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"");
        builder.write(" - ");
        builder.insertField("MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"");
        // Shows how to specify where the culture used for date formatting during field update and mail merge is chosen from.
        // Set the culture used during field update to the culture used by the field.
        doc.getFieldOptions().setFieldUpdateCultureSource(FieldUpdateCultureSource.FIELD_CODE);
//DateTime object issue
        //  doc.getMailMerge().ex
        // doc.getMailMerge().execute(new String[] { "Date2" }, new Object[] { new (2011, 1, 01) });
        // doc.MailMerge.Execute(new string[] { "Date2" }, new object[] { new DateTime(2011, 1, 01) });

        doc.save(dataDir + "InsertNestedFields Out.docx");
        //ExEnd:ChangeFieldUpdateCultureSource


        System.out.println("Nested fields inserted into the document successfully.");
    }
}




