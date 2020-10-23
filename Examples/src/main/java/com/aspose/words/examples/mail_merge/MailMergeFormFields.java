package com.aspose.words.examples.mail_merge;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

/**
 * This sample shows how to insert check boxes and text input form fields during mail merge into a document.
 */
public class MailMergeFormFields {
    /**
     * The main entry point for the application.
     */
    public static void main(String[] args) throws Exception {
    	// ExStart:MailMergeFormFields
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(MailMergeFormFields.class);

        // Load the template document.
        Document doc = new Document(dataDir + "Template.doc");

        // Setup mail merge event handler to do the custom work.
        doc.getMailMerge().setFieldMergingCallback(new HandleMergeField());

        // This is the data for mail merge.
        String[] fieldNames = new String[]{"RecipientName", "SenderName", "FaxNumber", "PhoneNumber",
                "Subject", "Body", "Urgent", "ForReview", "PleaseComment"};
        Object[] fieldValues = new Object[]{"Josh", "Jenny", "123456789", "", "Hello",
                "Test message 1", true, false, true};

        // Execute the mail merge.
        doc.getMailMerge().execute(fieldNames, fieldValues);

        // Save the finished document.
        doc.save(dataDir + "Template Out.doc");
        // ExEnd:MailMergeFormFields
        System.out.println("Mail merge performed successfully.");
    }

    // ExStart:HandleMergeField
    private static class HandleMergeField implements IFieldMergingCallback {
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

            // Another example, we want the Subject field to come out as text input form field.
            if ("Subject".equals(e.getFieldName())) {
                mBuilder.moveToMergeField(e.getFieldName());
                String textInputName = java.text.MessageFormat.format("{0}{1}", e.getFieldName(), e.getRecordIndex());
                mBuilder.insertTextInput(textInputName, TextFormFieldType.REGULAR, "", (String) e.getFieldValue(), 0);
            }
        }

        //ExStart:ImageFieldMerging
        public void imageFieldMerging(ImageFieldMergingArgs args) throws Exception {
        	args.setImageFileName("Image.png");
            args.setImageWidth(new MergeFieldImageDimension(200, MergeFieldImageDimensionUnit.POINT));
            args.setImageHeight(new MergeFieldImageDimension(200, MergeFieldImageDimensionUnit.PERCENT));
        }
        //ExEnd:ImageFieldMerging
        
        private DocumentBuilder mBuilder;
    }
    // ExEnd:HandleMergeField
}