package DocsExamples.Programming_with_documents;

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FormField;
import com.aspose.words.FieldType;
import com.aspose.words.FormFieldCollection;
import java.awt.Color;

@Test
public class WorkingWithFormFields extends DocsExamplesBase
{
    @Test
    public void insertFormFields() throws Exception
    {
        //ExStart:InsertFormFields
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        String[] items = { "One", "Two", "Three" };
        builder.insertComboBox("DropDown", items, 0);
        //ExEnd:InsertFormFields
    }

    @Test
    public void formFieldsWorkWithProperties() throws Exception
    {
        //ExStart:FormFieldsWorkWithProperties
        Document doc = new Document(getMyDir() + "Form fields.docx");
        FormField formField = doc.getRange().getFormFields().get(3);

        if (formField.getType() == FieldType.FIELD_FORM_TEXT_INPUT)
            formField.setResult("My name is " + formField.getName());
        //ExEnd:FormFieldsWorkWithProperties            
    }

    @Test
    public void formFieldsGetFormFieldsCollection() throws Exception
    {
        //ExStart:FormFieldsGetFormFieldsCollection
        Document doc = new Document(getMyDir() + "Form fields.docx");
        
        FormFieldCollection formFields = doc.getRange().getFormFields();
        //ExEnd:FormFieldsGetFormFieldsCollection
    }

    @Test
    public void formFieldsGetByName() throws Exception
    {
        //ExStart:FormFieldsFontFormatting
        //ExStart:FormFieldsGetByName
        Document doc = new Document(getMyDir() + "Form fields.docx");

        FormFieldCollection documentFormFields = doc.getRange().getFormFields();

        FormField formField1 = documentFormFields.get(3);
        FormField formField2 = documentFormFields.get("Text2");
        //ExEnd:FormFieldsGetByName

        formField1.getFont().setSize(20.0);
        formField2.getFont().setColor(Color.RED);
        //ExEnd:FormFieldsFontFormatting
    }
}
