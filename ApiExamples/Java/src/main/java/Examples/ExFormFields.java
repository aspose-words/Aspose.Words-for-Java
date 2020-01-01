package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////


import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.ByteArrayOutputStream;

public class ExFormFields extends ApiExampleBase {
    @Test
    public void formFieldsGetFormFieldsCollection() throws Exception {
        //ExStart
        //ExFor:Range.FormFields
        //ExFor:FormFieldCollection
        //ExSummary:Shows how to get a collection of form fields.
        Document doc = new Document(getMyDir() + "FormFields.doc");
        FormFieldCollection formFields = doc.getRange().getFormFields();
        //ExEnd
    }

    @Test
    public void formFieldsGetByName() throws Exception {
        //ExStart
        //ExFor:FormField
        //ExSummary:Shows how to access form fields.
        Document doc = new Document(getMyDir() + "FormFields.doc");
        FormFieldCollection documentFormFields = doc.getRange().getFormFields();

        FormField formField1 = documentFormFields.get(3);
        FormField formField2 = documentFormFields.get("CustomerName");
        //ExEnd
    }

    @Test
    public void formFieldsWorkWithProperties() throws Exception {
        //ExStart
        //ExFor:FormField
        //ExFor:FormField.Result
        //ExFor:FormField.Type
        //ExFor:FormField.Name
        //ExSummary:Shows how to work with form field name, type, and result.
        Document doc = new Document(getMyDir() + "FormFields.doc");

        FormField formField = doc.getRange().getFormFields().get(3);

        if (formField.getType() == FieldType.FIELD_FORM_TEXT_INPUT)
            formField.setResult("My name is " + formField.getName());
        //ExEnd
    }

    @Test
    public void insertAndRetrieveFormFields() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertTextInput
        //ExSummary:Shows how to insert form fields, set options and gather them back in for use
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input field. The unique name of this field is "TextInput1", the other parameters define
        // what type of FormField it is, the format of the text, the field result and the maximum text length (0 = no limit)
        builder.insertTextInput("TextInput1", TextFormFieldType.REGULAR, "", "", 0);
        //ExEnd
    }

    @Test
    public void deleteFormField() throws Exception {
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
    public void deleteFormFieldAssociatedWithBookmark() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("MyBookmark");
        builder.insertTextInput("TextInput1", TextFormFieldType.REGULAR, "TestFormField", "SomeText", 0);
        builder.endBookmark("MyBookmark");

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        BookmarkCollection bookmarkBeforeDeleteFormField = doc.getRange().getBookmarks();
        Assert.assertEquals(bookmarkBeforeDeleteFormField.get(0).getName(), "MyBookmark");

        FormField formField = doc.getRange().getFormFields().get(0);
        formField.removeField();

        BookmarkCollection bookmarkAfterDeleteFormField = doc.getRange().getBookmarks();
        Assert.assertEquals(bookmarkAfterDeleteFormField.get(0).getName(), "MyBookmark");
    }
}
