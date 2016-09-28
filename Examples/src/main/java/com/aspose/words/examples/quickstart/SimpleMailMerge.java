

package com.aspose.words.examples.quickstart;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class SimpleMailMerge
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SimpleMailMerge.class);

        Document doc = new Document(dataDir + "Template.doc");
        // Fill the fields in the document with user data.
        doc.getMailMerge().execute(
                new String[] { "FullName", "Company", "Address", "Address2", "City" },
                new Object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });
        // Saves the document to disk.
        doc.save(dataDir + "MailMerge Result Out.docx");

        System.out.println("Mail merge performed successfully.");
    }
}