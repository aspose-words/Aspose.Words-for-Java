package com.aspose.words.examples.mail_merge;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

/**
 * Created by awaishafeez on 12/19/2017.
 */
public class HandleMailMergeSwitches {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getSharedDataDir(ExecuteSimpleMailMerge.class) + "MailMerge/";

        // Open an existing document.
        Document doc = new Document(dataDir + "MailMergeSwitches.docx");

        doc.getMailMerge().setFieldMergingCallback(new MailMergeSwitches());

        // Fill the fields in the document with user data.
        doc.getMailMerge().execute(
                new String[]{"HTML_Name"},
                new Object[]{"James Bond"});

        dataDir = dataDir + "MergeSwitches_out.doc";
        doc.save(dataDir);

        System.out.println("\nSimple Mail merge performed with array data successfully.\nFile saved at " + dataDir);
    }

    // ExStart:HandleMailMergeSwitches
    static class MailMergeSwitches implements IFieldMergingCallback {
        public void fieldMerging(FieldMergingArgs e) throws Exception {

            if (e.getFieldName().startsWith("HTML")) {
                if (e.getField().getFieldCode().contains("\\b")) {
                    FieldMergeField field = e.getField();

                    DocumentBuilder builder = new DocumentBuilder(e.getDocument());
                    builder.moveToMergeField(e.getDocumentFieldName(), true, false);
                    builder.write(field.getTextBefore());
                    builder.insertHtml(e.getFieldValue().toString());

                    e.setText("");
                }
            }
        }

        public void imageFieldMerging(ImageFieldMergingArgs args) throws Exception {
            // Do Nothing
        }
    }
    // ExEnd:HandleMailMergeSwitches
}
