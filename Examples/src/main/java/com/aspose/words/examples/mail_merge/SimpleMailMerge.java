package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;


public class SimpleMailMerge {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SimpleMailMerge.class);
        // Open the document.
        Document doc = new Document(dataDir + "MailMerge.ExecuteArray.doc");
        doc.getMailMerge().setUseNonMergeFields(true);
        doc.getMailMerge().execute(
                new String[]{"FullName", "Company", "Address", "Address2", "City"},
                new Object[]{"James Bond", "MI5 Headquarters", "Milbank", "", "London"});
        // Save the output document to disk.
        doc.save(dataDir + "Output.doc");
        //ExEnd:1
    }
}