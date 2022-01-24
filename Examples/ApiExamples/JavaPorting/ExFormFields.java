// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
import java.awt.Color;
import com.aspose.words.Run;
import com.aspose.words.BreakType;
import com.aspose.words.FormFieldCollection;
import java.util.Iterator;
import com.aspose.ms.System.msConsole;
import com.aspose.words.DocumentVisitor;
import com.aspose.words.VisitorAction;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.words.FieldCollection;
import com.aspose.words.DropDownItemCollection;


@Test
public class ExFormFields extends ApiExampleBase
{
    @Test
    public void create() throws Exception
    {
        //ExStart
        //ExFor:FormField
        //ExFor:FormField.Result
        //ExFor:FormField.Type
        //ExFor:FormField.Name
        //ExSummary:Shows how to insert a combo box.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Please select a fruit: ");

        // Insert a combo box which will allow a user to choose an option from a collection of strings.
        FormField comboBox = builder.insertComboBox("MyComboBox", new String[] { "Apple", "Banana", "Cherry" }, 0);

        Assert.assertEquals("MyComboBox", comboBox.getName());
        Assert.assertEquals(FieldType.FIELD_FORM_DROP_DOWN, comboBox.getType());
        Assert.assertEquals("Apple", comboBox.getResult());

        // The form field will appear in the form of a "select" html tag.
        doc.save(getArtifactsDir() + "FormFields.Create.html");
        //ExEnd

        doc = new Document(getArtifactsDir() + "FormFields.Create.html");
        comboBox = doc.getRange().getFormFields().get(0);

        Assert.assertEquals("MyComboBox", comboBox.getName());
        Assert.assertEquals(FieldType.FIELD_FORM_DROP_DOWN, comboBox.getType());
        Assert.assertEquals("Apple", comboBox.getResult());
    }

    @Test
    public void textInput() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertTextInput
        //ExSummary:Shows how to insert a text input form field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Please enter text here: ");

        // Insert a text input field, which will allow the user to click it and enter text.
        // Assign some placeholder text that the user may overwrite and pass
        // a maximum text length of 0 to apply no limit on the form field's contents.
        builder.insertTextInput("TextInput1", TextFormFieldType.REGULAR, "", "Placeholder text", 0);

        // The form field will appear in the form of an "input" html tag, with a type of "text".
        doc.save(getArtifactsDir() + "FormFields.TextInput.html");
        //ExEnd

        doc = new Document(getArtifactsDir() + "FormFields.TextInput.html");

        FormField textInput = doc.getRange().getFormFields().get(0);

        Assert.assertEquals("TextInput1", textInput.getName());
        Assert.assertEquals(TextFormFieldType.REGULAR, textInput.getTextInputType());
        Assert.assertEquals("", textInput.getTextInputFormat());
        Assert.assertEquals("Placeholder text", textInput.getResult());
        Assert.assertEquals(0, textInput.getMaxLength());
    }

    @Test
    public void deleteFormField() throws Exception
    {
        //ExStart
        //ExFor:FormField.RemoveField
        //ExSummary:Shows how to delete a form field.
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

    @Test
    public void formFieldFontFormatting() throws Exception
    {
        //ExStart
        //ExFor:FormField
        //ExSummary:Shows how to formatting the entire FormField, including the field value.
        Document doc = new Document(getMyDir() + "Form fields.docx");

        FormField formField = doc.getRange().getFormFields().get(0);
        formField.getFont().setBold(true);
        formField.getFont().setSize(24.0);
        formField.getFont().setColor(Color.RED);

        formField.setResult("Aspose.FormField");

        doc = DocumentHelper.saveOpen(doc);
        
        Run formFieldRun = doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(1);

        Assert.assertEquals("Aspose.FormField", formFieldRun.getText());
        Assert.assertEquals(true, formFieldRun.getFont().getBold());
        Assert.assertEquals(24, formFieldRun.getFont().getSize());
        Assert.assertEquals(Color.RED.getRGB(), formFieldRun.getFont().getColor().getRGB());
        //ExEnd
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
    //ExSummary:Shows how insert different kinds of form fields into a document, and process them with using a document visitor implementation.
    @Test //ExSkip
    public void visitor() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert a combo box.
        builder.write("Choose a value from this combo box: ");
        FormField comboBox = builder.insertComboBox("MyComboBox", new String[] { "One", "Two", "Three" }, 0);
        comboBox.setCalculateOnExit(true);
        Assert.assertEquals(3, comboBox.getDropDownItems().getCount());
        Assert.assertEquals(0, comboBox.getDropDownSelectedIndex());
        Assert.assertTrue(comboBox.getEnabled());

        builder.insertBreak(BreakType.PARAGRAPH_BREAK);

        // Use a document builder to insert a check box.
        builder.write("Click this check box to tick/untick it: ");
        FormField checkBox = builder.insertCheckBox("MyCheckBox", false, 50);
        checkBox.isCheckBoxExactSize(true);
        checkBox.setHelpText("Right click to check this box");
        checkBox.setOwnHelp(true);
        checkBox.setStatusText("Checkbox status text");
        checkBox.setOwnStatus(true);
        Assert.assertEquals(50.0d, checkBox.getCheckBoxSize());
        Assert.assertFalse(checkBox.getChecked());
        Assert.assertFalse(checkBox.getDefault());

        builder.insertBreak(BreakType.PARAGRAPH_BREAK);

        // Use a document builder to insert text input form field.
        builder.write("Enter text here: ");
        FormField textInput = builder.insertTextInput("MyTextInput", TextFormFieldType.REGULAR, "", "Placeholder text", 50);
        textInput.setEntryMacro("EntryMacro");
        textInput.setExitMacro("ExitMacro");
        textInput.setTextInputDefault("Regular");
        textInput.setTextInputFormat("FIRST CAPITAL");
        textInput.setTextInputValue("New placeholder text");
        Assert.assertEquals(TextFormFieldType.REGULAR, textInput.getTextInputType());
        Assert.assertEquals(50, textInput.getMaxLength());

        // This collection contains all our form fields.
        FormFieldCollection formFields = doc.getRange().getFormFields();
        Assert.assertEquals(3, formFields.getCount());

        // Fields display our form fields. We can see their field codes by opening this document
        // in Microsoft and pressing Alt + F9. These fields have no switches,
        // and members of the FormField object fully govern their form fields' content.
        Assert.assertEquals(3, doc.getRange().getFields().getCount());
        Assert.assertEquals(" FORMDROPDOWN \u0001", doc.getRange().getFields().get(0).getFieldCode());
        Assert.assertEquals(" FORMCHECKBOX \u0001", doc.getRange().getFields().get(1).getFieldCode());
        Assert.assertEquals(" FORMTEXT \u0001", doc.getRange().getFields().get(2).getFieldCode());

        // Allow each form field to accept a document visitor.
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
        doc.save(getArtifactsDir() + "FormFields.Visitor.html");
        testFormField(doc); //ExSkip
    }

    /// <summary>
    /// Visitor implementation that prints details of form fields that it visits. 
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
                    appendLine("\tDrop-down items count: " + formField.getDropDownItems().getCount() + ", default selected item index: " + formField.getDropDownSelectedIndex());
                    AppendLine("\tDrop-down items: " + String.Join(", ", formField.getDropDownItems().ToArray()));
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
        TestUtil.verifyField(FieldType.FIELD_FORM_TEXT_INPUT, " FORMTEXT \u0001", "New placeholder text", doc.getRange().getFields().get(2));

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
        Assert.assertEquals("New placeholder text", formFields.get(2).getResult());
    }

    @Test
    public void dropDownItemCollection() throws Exception
    {
        //ExStart
        //ExFor:Fields.DropDownItemCollection
        //ExFor:Fields.DropDownItemCollection.Add(String)
        //ExFor:Fields.DropDownItemCollection.Clear
        //ExFor:Fields.DropDownItemCollection.Contains(String)
        //ExFor:Fields.DropDownItemCollection.Count
        //ExFor:Fields.DropDownItemCollection.GetEnumerator
        //ExFor:Fields.DropDownItemCollection.IndexOf(String)
        //ExFor:Fields.DropDownItemCollection.Insert(Int32, String)
        //ExFor:Fields.DropDownItemCollection.Item(Int32)
        //ExFor:Fields.DropDownItemCollection.Remove(String)
        //ExFor:Fields.DropDownItemCollection.RemoveAt(Int32)
        //ExSummary:Shows how to insert a combo box field, and edit the elements in its item collection.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a combo box, and then verify its collection of drop-down items.
        // In Microsoft Word, the user will click the combo box,
        // and then choose one of the items of text in the collection to display.
        String[] items = { "One", "Two", "Three" };
        FormField comboBoxField = builder.insertComboBox("DropDown", items, 0);
        DropDownItemCollection dropDownItems = comboBoxField.getDropDownItems();

        Assert.assertEquals(3, dropDownItems.getCount());
        Assert.assertEquals("One", dropDownItems.get(0));
        Assert.assertEquals(1, dropDownItems.indexOf("Two"));
        Assert.assertTrue(dropDownItems.contains("Three"));

        // There are two ways of adding a new item to an existing collection of drop-down box items.
        // 1 -  Append an item to the end of the collection:
        dropDownItems.add("Four");

        // 2 -  Insert an item before another item at a specified index:
        dropDownItems.insert(3, "Three and a half");

        Assert.assertEquals(5, dropDownItems.getCount());

        // Iterate over the collection and print every element.
        Iterator<String> dropDownCollectionEnumerator = dropDownItems.iterator();
        try /*JAVA: was using*/
    	{
            while (dropDownCollectionEnumerator.hasNext())
                System.out.println(dropDownCollectionEnumerator.next());
    	}
        finally { if (dropDownCollectionEnumerator != null) dropDownCollectionEnumerator.close(); }

        // There are two ways of removing elements from a collection of drop-down items.
        // 1 -  Remove an item with contents equal to the passed string:
        dropDownItems.remove("Four");

        // 2 -  Remove an item at an index:
        dropDownItems.removeAt(3);

        Assert.assertEquals(3, dropDownItems.getCount());
        Assert.assertFalse(dropDownItems.contains("Three and a half"));
        Assert.assertFalse(dropDownItems.contains("Four"));

        doc.save(getArtifactsDir() + "FormFields.DropDownItemCollection.html");

        // Empty the whole collection of drop-down items.
        dropDownItems.clear();
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        dropDownItems = doc.getRange().getFormFields().get(0).getDropDownItems();

        Assert.assertEquals(0, dropDownItems.getCount());

        doc = new Document(getArtifactsDir() + "FormFields.DropDownItemCollection.html");
        dropDownItems = doc.getRange().getFormFields().get(0).getDropDownItems();

        Assert.assertEquals(3, dropDownItems.getCount());
        Assert.assertEquals("One", dropDownItems.get(0));
        Assert.assertEquals("Two", dropDownItems.get(1));
        Assert.assertEquals("Three", dropDownItems.get(2));
    }
}
