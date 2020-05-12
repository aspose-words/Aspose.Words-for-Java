// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FormField;
import org.testng.Assert;
import com.aspose.words.FieldType;
import com.aspose.words.TextFormFieldType;
import com.aspose.words.BookmarkCollection;
import com.aspose.words.FormFieldCollection;
import java.util.Iterator;
import com.aspose.ms.System.msConsole;
import com.aspose.words.DocumentVisitor;
import com.aspose.words.VisitorAction;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.words.FieldCollection;


@Test
public class ExFormFields extends ApiExampleBase
{
    @Test
    public void formFieldsWorkWithProperties() throws Exception
    {
        //ExStart
        //ExFor:FormField
        //ExFor:FormField.Result
        //ExFor:FormField.Type
        //ExFor:FormField.Name
        //ExSummary:Shows how to work with form field name, type, and result.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a DocumentBuilder to insert a combo box form field
        FormField comboBox = builder.insertComboBox("MyComboBox", new String[] { "One", "Two", "Three" }, 0);

        // Verify some of our form field's attributes
        Assert.assertEquals("MyComboBox", comboBox.getName());
        Assert.assertEquals(FieldType.FIELD_FORM_DROP_DOWN, comboBox.getType());
        Assert.assertEquals("One", comboBox.getResult());
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        comboBox = doc.getRange().getFormFields().get(0);

        Assert.assertEquals("MyComboBox", comboBox.getName());
        Assert.assertEquals(FieldType.FIELD_FORM_DROP_DOWN, comboBox.getType());
        Assert.assertEquals("One", comboBox.getResult());
    }

    @Test
    public void insertAndRetrieveFormFields() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertTextInput
        //ExSummary:Shows how to insert form fields, set options and gather them back in for use.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input field. The unique name of this field is "TextInput1", the other parameters define
        // what type of FormField it is, the format of the text, the field result and the maximum text length (0 = no limit)
        builder.insertTextInput("TextInput1", TextFormFieldType.REGULAR, "", "", 0);
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        FormField textInput = doc.getRange().getFormFields().get(0);

        Assert.assertEquals("TextInput1", textInput.getName());
        Assert.assertEquals(TextFormFieldType.REGULAR, textInput.getTextInputType());
        Assert.assertEquals("", textInput.getTextInputFormat());
        Assert.assertEquals("", textInput.getResult());
        Assert.assertEquals(0, textInput.getMaxLength());
    }

    @Test
    public void deleteFormField() throws Exception
    {
        //ExStart
        //ExFor:FormField.RemoveField
        //ExSummary:Shows how to delete complete form field.
        Document doc = new Document(getMyDir() + "Form fields.docx");

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

        doc = DocumentHelper.saveOpen(doc);

        BookmarkCollection bookmarkBeforeDeleteFormField = doc.getRange().getBookmarks();
        Assert.assertEquals("MyBookmark", bookmarkBeforeDeleteFormField.get(0).getName());

        FormField formField = doc.getRange().getFormFields().get(0);
        formField.removeField();

        BookmarkCollection bookmarkAfterDeleteFormField = doc.getRange().getBookmarks();
        Assert.assertEquals("MyBookmark", bookmarkAfterDeleteFormField.get(0).getName());
    }

    //ExStart
    //ExFor:FormField.Accept(DocumentVisitor)
    //ExFor:FormField.CalculateOnExit
    //ExFor:FormField.CheckBoxSize
    //ExFor:FormField.Checked
    //ExFor:FormField.Default
    //ExFor:FormField.DropDownItems
    //ExFor:FormField.DropDownSelectedIndex
    //ExFor:FormField.Enabled
    //ExFor:FormField.EntryMacro
    //ExFor:FormField.ExitMacro
    //ExFor:FormField.HelpText
    //ExFor:FormField.IsCheckBoxExactSize
    //ExFor:FormField.MaxLength
    //ExFor:FormField.OwnHelp
    //ExFor:FormField.OwnStatus
    //ExFor:FormField.SetTextInputValue(Object)
    //ExFor:FormField.StatusText
    //ExFor:FormField.TextInputDefault
    //ExFor:FormField.TextInputFormat
    //ExFor:FormField.TextInputType
    //ExFor:FormFieldCollection
    //ExFor:FormFieldCollection.Clear
    //ExFor:FormFieldCollection.Count
    //ExFor:FormFieldCollection.GetEnumerator
    //ExFor:FormFieldCollection.Item(Int32)
    //ExFor:FormFieldCollection.Item(String)
    //ExFor:FormFieldCollection.Remove(String)
    //ExFor:FormFieldCollection.RemoveAt(Int32)
    //ExFor:Range.FormFields
    //ExSummary:Shows how insert different kinds of form fields into a document and process them with a visitor implementation.
    @Test //ExSkip
    public void formField() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert a combo box
        FormField comboBox = builder.insertComboBox("MyComboBox", new String[] { "One", "Two", "Three" }, 0);
        comboBox.setCalculateOnExit(true);
        Assert.assertEquals(3, comboBox.getDropDownItems().getCount());
        Assert.assertEquals(0, comboBox.getDropDownSelectedIndex());
        Assert.assertTrue(comboBox.getEnabled());

        // Use a document builder to insert a check box
        FormField checkBox = builder.insertCheckBox("MyCheckBox", false, 50);
        checkBox.isCheckBoxExactSize(true);
        checkBox.setHelpText("Right click to check this box");
        checkBox.setOwnHelp(true);
        checkBox.setStatusText("Checkbox status text");
        checkBox.setOwnStatus(true);
        Assert.assertEquals(50.0d, checkBox.getCheckBoxSize());
        Assert.assertFalse(checkBox.getChecked());
        Assert.assertFalse(checkBox.getDefault());

        builder.writeln();

        // Use a document builder to insert text input form field
        FormField textInput = builder.insertTextInput("MyTextInput", TextFormFieldType.REGULAR, "", "Your text goes here", 50);
        textInput.setEntryMacro("EntryMacro");
        textInput.setExitMacro("ExitMacro");
        textInput.setTextInputDefault("Regular");
        textInput.setTextInputFormat("FIRST CAPITAL");
        textInput.setTextInputValue("This value overrides the one we set during initialization");
        Assert.assertEquals(TextFormFieldType.REGULAR, textInput.getTextInputType());
        Assert.assertEquals(50, textInput.getMaxLength());

        // Get the collection of form fields that has accumulated in our document
        FormFieldCollection formFields = doc.getRange().getFormFields();
        Assert.assertEquals(3, formFields.getCount());

        // Our form fields are represented as fields, with field codes FORMDROPDOWN, FORMCHECKBOX and FORMTEXT respectively,
        // made visible by pressing Alt + F9 in Microsoft Word
        // These fields have no switches and the content of their form fields is fully governed by members of the FormField object
        Assert.assertEquals(3, doc.getRange().getFields().getCount());

        // Iterate over the collection with an enumerator, accepting a visitor with each form field
        FormFieldVisitor formFieldVisitor = new FormFieldVisitor();

        Iterator<FormField> fieldEnumerator = formFields.iterator();
        try /*JAVA: was using*/
    	{
            while (fieldEnumerator.hasNext())
                fieldEnumerator.next().accept(formFieldVisitor);
    	}
        finally { if (fieldEnumerator != null) fieldEnumerator.close(); }

        System.out.println(formFieldVisitor.getText());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.FormField.docx");
        testFormField(doc); //ExSkip
    }

    /// <summary>
    /// Visitor implementation that prints information about visited form fields. 
    /// </summary>
    public static class FormFieldVisitor extends DocumentVisitor
    {
        public FormFieldVisitor()
        {
            mBuilder = new StringBuilder();
        }

        /// <summary>
        /// Called when a FormField node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitFormField(FormField formField)
        {
            appendLine(formField.getType() + ": \"" + formField.getName() + "\"");
            appendLine("\tStatus: " + (formField.getEnabled() ? "Enabled" : "Disabled"));
            appendLine("\tHelp Text:  " + formField.getHelpText());
            appendLine("\tEntry macro name: " + formField.getEntryMacro());
            appendLine("\tExit macro name: " + formField.getExitMacro());

            switch (formField.getType())
            {
                case FieldType.FIELD_FORM_DROP_DOWN:
                    appendLine("\tDrop down items count: " + formField.getDropDownItems().getCount() + ", default selected item index: " + formField.getDropDownSelectedIndex());
                    AppendLine("\tDrop down items: " + String.Join(", ", formField.getDropDownItems().ToArray()));
                    break;
                case FieldType.FIELD_FORM_CHECK_BOX:
                    appendLine("\tCheckbox size: " + formField.getCheckBoxSize());
                    appendLine("\t" + "Checkbox is currently: " + (formField.getChecked() ? "checked, " : "unchecked, ") + "by default: " + (formField.getDefault() ? "checked" : "unchecked"));
                    break;
                case FieldType.FIELD_FORM_TEXT_INPUT:
                    appendLine("\tInput format: " + formField.getTextInputFormat());
                    appendLine("\tCurrent contents: " + formField.getResult());
                    break;
            }

            // Let the visitor continue visiting other nodes.
            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Adds newline char-terminated text to the current output.
        /// </summary>
        private void appendLine(String text)
        {
            msStringBuilder.append(mBuilder, text + '\n');
        }

        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        public String getText()
        {
            return mBuilder.toString();
        }

        private /*final*/ StringBuilder mBuilder;
    }
    //ExEnd

    private void testFormField(Document doc) throws Exception
    {
        doc = DocumentHelper.saveOpen(doc);
        FieldCollection fields = doc.getRange().getFields();
        Assert.assertEquals(3, fields.getCount());

        TestUtil.verifyField(FieldType.FIELD_FORM_DROP_DOWN, " FORMDROPDOWN \u0001", "", doc.getRange().getFields().get(0));
        TestUtil.verifyField(FieldType.FIELD_FORM_CHECK_BOX, " FORMCHECKBOX \u0001", "", doc.getRange().getFields().get(1));
        TestUtil.verifyField(FieldType.FIELD_FORM_TEXT_INPUT, " FORMTEXT \u0001", "This value overrides the one we set during initialization", doc.getRange().getFields().get(2));

        FormFieldCollection formFields = doc.getRange().getFormFields();
        Assert.assertEquals(3, formFields.getCount());

        Assert.assertEquals(FieldType.FIELD_FORM_DROP_DOWN, formFields.get(0).getType());
        Assert.assertEquals(new String[] { "One", "Two", "Three" }, formFields.get(0).getDropDownItems());
        Assert.assertTrue(formFields.get(0).getCalculateOnExit());
        Assert.assertEquals(0, formFields.get(0).getDropDownSelectedIndex());
        Assert.assertTrue(formFields.get(0).getEnabled());
        Assert.assertEquals("One", formFields.get(0).getResult());

        Assert.assertEquals(FieldType.FIELD_FORM_CHECK_BOX, formFields.get(1).getType());
        Assert.assertTrue(formFields.get(1).isCheckBoxExactSize());
        Assert.assertEquals("Right click to check this box", formFields.get(1).getHelpText());
        Assert.assertTrue(formFields.get(1).getOwnHelp());
        Assert.assertEquals("Checkbox status text", formFields.get(1).getStatusText());
        Assert.assertTrue(formFields.get(1).getOwnStatus());
        Assert.assertEquals(50.0d, formFields.get(1).getCheckBoxSize());
        Assert.assertFalse(formFields.get(1).getChecked());
        Assert.assertFalse(formFields.get(1).getDefault());
        Assert.assertEquals("0", formFields.get(1).getResult());

        Assert.assertEquals(FieldType.FIELD_FORM_TEXT_INPUT, formFields.get(2).getType());
        Assert.assertEquals("EntryMacro", formFields.get(2).getEntryMacro());
        Assert.assertEquals("ExitMacro", formFields.get(2).getExitMacro());
        Assert.assertEquals("Regular", formFields.get(2).getTextInputDefault());
        Assert.assertEquals("FIRST CAPITAL", formFields.get(2).getTextInputFormat());
        Assert.assertEquals(TextFormFieldType.REGULAR, formFields.get(2).getTextInputType());
        Assert.assertEquals(50, formFields.get(2).getMaxLength());
        Assert.assertEquals("This value overrides the one we set during initialization", formFields.get(2).getResult());
    }
}
