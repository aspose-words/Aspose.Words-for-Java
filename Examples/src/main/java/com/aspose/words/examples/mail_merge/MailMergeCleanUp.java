package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;
import com.aspose.words.MailMergeCleanupOptions;
import com.aspose.words.examples.Utils;
import com.aspose.words.net.System.Data.DataSet;

public class MailMergeCleanUp {
    public static void main(String[] args) throws Exception {
        String dataDir = Utils.getSharedDataDir(MailMergeCleanUp.class) + "MailMerge/";

        RemoveUnmergedRegions(dataDir);
        RemoveEmptyParagraphs(dataDir);
        cleanupParagraphsWithPunctuationMarks(dataDir);
    }

    public static void RemoveEmptyTableRows(String dataDir) throws Exception {
    	//ExStart: RemoveEmptyTableRows
    	// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET
    	Document doc = new Document(dataDir + "RemoveRowfromTable.docx");

    	doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_EMPTY_TABLE_ROWS);

    	doc.getMailMerge().execute(new String[] { "FullName", "Company", "Address", "Address2", "City" }, 
    	    new Object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });
    	            
    	doc.save(dataDir + "MailMerge.ExecuteArray_out.doc");
    	//ExEnd: RemoveEmptyTableRows
    }
    public static void RemoveContainingFields(String dataDir) throws Exception {
    	//ExStart: RemoveContainingFields
    	Document doc = new Document(dataDir + "RemoveRowfromTable.docx");

    	doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_CONTAINING_FIELDS);

    	doc.getMailMerge().execute(new String[] { "FullName", "Company", "Address", "Address2", "City" }, 
    	    new Object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });
    	            
    	doc.save(dataDir + "MailMerge.ExecuteArray_out.doc");
    	//ExEnd: RemoveContainingFields
    }
    public static void RemoveUnusedFields(String dataDir) throws Exception {
        //ExStart:RemoveUnusedFields
    	Document doc = new Document(dataDir + "RemoveRowfromTable.docx");

    	doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_UNUSED_FIELDS);

    	doc.getMailMerge().execute(new String[] { "FullName", "Company", "Address", "Address2", "City" }, 
    	    new Object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });
    	            
    	doc.save(dataDir + "MailMerge.ExecuteArray_out.doc");
        //ExEnd:RemoveUnusedFields
    }
    
    public static void RemoveUnmergedRegions(String dataDir) throws Exception {
        //ExStart:RemoveUnmergedRegions
    	Document doc = new Document(dataDir + "TestFile Empty.doc");

    	// Create an empty data source in the form of a DataSet containing no DataTable objects.
    	DataSet data = new DataSet();

    	// Enable the MailMergeCleanupOptions.RemoveUnusedRegions option.
    	doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_UNUSED_REGIONS);

    	// Merge the data with the document by executing mail merge which will have no effect as there is no data.
    	// However the regions found in the document will be removed automatically as they are unused.
    	doc.getMailMerge().executeWithRegions(data);

        // Save the output document to disk.
        doc.save(dataDir + "TestFile.RemoveEmptyRegions Out.doc");
        //ExEnd:RemoveUnmergedRegions
    }
    
    public static void RemoveEmptyParagraphs(String dataDir) throws Exception {
    	//ExStart: RemoveEmptyParagraphs
    	Document doc = new Document(dataDir + "RemoveRowfromTable.docx");

    	doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS);

    	doc.getMailMerge().execute(new String[] { "FullName", "Company", "Address", "Address2", "City" }, 
    	    new Object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });
    	            
    	doc.save(dataDir + "MailMerge.ExecuteArray_out.doc");
    	//ExEnd: RemoveEmptyParagraphs
    }
    
    public static void cleanupParagraphsWithPunctuationMarks(String dataDir) throws Exception {
        // Open the document
        Document doc = new Document(dataDir + "MailMerge.CleanupPunctuationMarks.docx");

        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS);
        doc.getMailMerge().setCleanupParagraphsWithPunctuationMarks(false);

        doc.getMailMerge().execute(new String[]{"field1", "field2"}, new Object[]{"", ""});

        dataDir = dataDir + "MailMerge.CleanupPunctuationMarks_out.docx";
        // Save the output document to disk.
        doc.save(dataDir);

        System.out.println("\nMail merge performed with cleanup paragraphs having punctuation marks successfully.\nFile saved at " + dataDir);
    }
}
