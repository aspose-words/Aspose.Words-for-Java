package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

public class MailMergeAndConditionalField {
    public static void main(String[] args) throws Exception {
        // ExStart:MailMergeAndConditionalField
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(MailMergeAndConditionalField.class);
        // Open an existing document.
        Document doc = new Document(dataDir + "UnconditionalMergeFieldsAndRegions.docx");

        //Merge fields and merge regions are merged regardless of the parent IF field's condition.
        doc.getMailMerge().setUnconditionalMergeFieldsAndRegions(true);

        // Fill the fields in the document with user data.
        doc.getMailMerge().execute(
                new String[]{"FullName"},
                new Object[]{"James Bond"});

        doc.save(dataDir + "UnconditionalMergeFieldsAndRegions_out.docx", SaveFormat.DOCX);
        // ExEnd:MailMergeAndConditionalField
        System.out.println("\nMail merge with conditional field has performed successfully.\nFile saved at " + dataDir);
    }
}
