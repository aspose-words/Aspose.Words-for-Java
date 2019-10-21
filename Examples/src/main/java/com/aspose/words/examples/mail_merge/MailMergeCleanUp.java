package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;
import com.aspose.words.MailMergeCleanupOptions;
import com.aspose.words.examples.Utils;

public class MailMergeCleanUp {
    public static void main(String[] args) throws Exception {
        String dataDir = Utils.getSharedDataDir(MailMergeCleanUp.class) + "MailMerge/";

        cleanupParagraphsWithPunctuationMarks(dataDir);
    }

    public static void cleanupParagraphsWithPunctuationMarks(String dataDir) throws Exception {
        // ExStart:CleanupParagraphsWithPunctuationMarks
        // Open the document
        Document doc = new Document(dataDir + "MailMerge.CleanupPunctuationMarks.docx");

        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS);
        doc.getMailMerge().setCleanupParagraphsWithPunctuationMarks(false);

        doc.getMailMerge().execute(new String[]{"field1", "field2"}, new Object[]{"", ""});

        dataDir = dataDir + "MailMerge.CleanupPunctuationMarks_out.docx";
        // Save the output document to disk.
        doc.save(dataDir);
        // ExEnd:CleanupParagraphsWithPunctuationMarks

        System.out.println("\nMail merge performed with cleanup paragraphs having punctuation marks successfully.\nFile saved at " + dataDir);
    }
}
