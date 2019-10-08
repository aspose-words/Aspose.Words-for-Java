package com.aspose.words.examples.mail_merge;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

//ExStart:

/**
 * This sample shows how to insert check boxes and text input form fields during
 * mail merge into a document.
 */
public class InsertCheckBoxesOrHTMLDuringMailMerge {

    private static final String dataDir = Utils.getSharedDataDir(InsertCheckBoxesOrHTMLDuringMailMerge.class) + "MailMerge/";

    public static void main(String[] args) throws Exception {
        execute();
    }

    private static void execute() throws Exception {

        // Load the template document.
        Document doc = new Document(dataDir + "Template.doc");

        // Setup mail merge event handler to do the custom work.
        doc.getMailMerge().setFieldMergingCallback(new HandleMergeField());

        // This is the data for mail merge.
        String[] fieldNames = new String[]{"RecipientName", "SenderName", "FaxNumber", "PhoneNumber", "Subject", "Body", "Urgent", "ForReview", "PleaseComment"};
        Object[] fieldValues = new Object[]{"Josh", "Jenny", "123456789", "", "Hello", "<b>HTML Body Test message 1</b>", true, false, true};

        // Execute the mail merge.
        doc.getMailMerge().execute(fieldNames, fieldValues);

        // Save the finished document.
        doc.save(dataDir + "Template Out.doc");
    }
}

class HandleMergeField implements IFieldMergingCallback {
    /**
     * This handler is called for every mail merge field found in the document,
     * for every record found in the data source.
     */
    public void fieldMerging(FieldMergingArgs e) throws Exception {
        if (mBuilder == null)
            mBuilder = new DocumentBuilder(e.getDocument());

        // We decided that we want all boolean values to be output as check box form fields.
        if (e.getFieldValue() instanceof Boolean) {
            // Move the "cursor" to the current merge field.
            mBuilder.moveToMergeField(e.getFieldName());

            // It is nice to give names to check boxes. Lets generate a name such as MyField21 or so.
            String checkBoxName = java.text.MessageFormat.format("{0}{1}", e.getFieldName(), e.getRecordIndex());

            // Insert a check box.
            mBuilder.insertCheckBox(checkBoxName, (Boolean) e.getFieldValue(), 0);

            // Nothing else to do for this field.
            return;
        }

        // We want to insert html during mail merge.
        if ("Body".equals(e.getFieldName())) {
            mBuilder.moveToMergeField(e.getFieldName());
            mBuilder.insertHtml((String) e.getFieldValue());
        }

        // Another example, we want the Subject field to come out as text input form field.
        if ("Subject".equals(e.getFieldName())) {
            mBuilder.moveToMergeField(e.getFieldName());
            String textInputName = java.text.MessageFormat.format("{0}{1}", e.getFieldName(), e.getRecordIndex());
            mBuilder.insertTextInput(textInputName, TextFormFieldType.REGULAR, "", (String) e.getFieldValue(), 0);
        }
    }

    public void imageFieldMerging(ImageFieldMergingArgs args) throws Exception {
        // Do nothing.
    }

    private DocumentBuilder mBuilder;
}
//ExEnd: