// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.FormFieldCollection;
import com.aspose.words.FormField;
import com.aspose.words.FieldType;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.TextFormFieldType;
import org.testng.Assert;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.words.SaveFormat;
import com.aspose.words.BookmarkCollection;
import com.aspose.ms.NUnit.Framework.msAssert;


@Test
public class ExFormFields extends ApiExampleBase
{
    @Test
    public void formFieldsGetFormFieldsCollection() throws Exception
    {
        //ExStart
        //ExFor:Range.FormFields
        //ExFor:FormFieldCollection
        //ExId:FormFieldsGetFormFieldsCollection
        //ExSummary:Shows how to get a collection of form fields.
        Document doc = new Document(getMyDir() + "FormFields.doc");
        FormFieldCollection formFields = doc.getRange().getFormFields();
        //ExEnd
    }

    @Test
    public void formFieldsGetByName() throws Exception
    {
        //ExStart
        //ExFor:FormField
        //ExId:FormFieldsGetByName
        //ExSummary:Shows how to access form fields.
        Document doc = new Document(getMyDir() + "FormFields.doc");
        FormFieldCollection documentFormFields = doc.getRange().getFormFields();

        FormField formField1 = documentFormFields.get(3);
        FormField formField2 = documentFormFields.get("CustomerName");
        //ExEnd
    }

    @Test
    public void formFieldsWorkWithProperties() throws Exception
    {
        //ExStart
        //ExFor:FormField
        //ExFor:FormField.Result
        //ExFor:FormField.Type
        //ExFor:FormField.Name
        //ExId:FormFieldsWorkWithProperties
        //ExSummary:Shows how to work with form field name, type, and result.
        Document doc = new Document(getMyDir() + "FormFields.doc");

        FormField formField = doc.getRange().getFormFields().get(3);

        if (((formField.getType()) == (FieldType.FIELD_FORM_TEXT_INPUT)))
            formField.setResult("My name is " + formField.getName());
        //ExEnd
    }

    @Test
    public void insertAndRetrieveFormFields() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertTextInput
        //ExId:FormFieldsInsertAndRetrieve
        //ExSummary:Shows how to insert form fields, set options and gather them back in for use 
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input field. The unique name of this field is "TextInput1", the other parameters define
        // what type of FormField it is, the format of the text, the field result and the maximum text length (0 = no limit)
        builder.insertTextInput("TextInput1", TextFormFieldType.REGULAR, "", "", 0);
        //ExEnd
    }

    @Test
    public void deleteFormField() throws Exception
    {
        //ExStart
        //ExFor:FormField.RemoveField
        //ExSummary:Shows how to delete complete form field
        Document doc = new Document(getMyDir() + "FormFields.doc");

        FormField formField = doc.getRange().getFormFields().get(3);
        formField.removeField();
        //ExEnd

        FormField formFieldAfter = doc.getRange().getFormFields().get(3);

        Assert.assertNull(formFieldAfter);
    }

    @Test
    public void deleteFormFieldAssociatedWithBookmark() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("MyBookmark");
        builder.insertTextInput("TextInput1", TextFormFieldType.REGULAR, "TestFormField", "SomeText", 0);
        builder.endBookmark("MyBookmark");

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        BookmarkCollection bookmarkBeforeDeleteFormField = doc.getRange().getBookmarks();
        msAssert.areEqual("MyBookmark", bookmarkBeforeDeleteFormField.get(0).getName());

        FormField formField = doc.getRange().getFormFields().get(0);
        formField.removeField();

        BookmarkCollection bookmarkAfterDeleteFormField = doc.getRange().getBookmarks();
        msAssert.areEqual("MyBookmark", bookmarkAfterDeleteFormField.get(0).getName());
    }
}
