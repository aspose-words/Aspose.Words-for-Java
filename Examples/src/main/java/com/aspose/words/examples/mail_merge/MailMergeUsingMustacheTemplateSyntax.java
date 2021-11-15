package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;
import com.aspose.words.net.System.Data.DataRow;
import com.aspose.words.net.System.Data.DataTable;

import javax.xml.parsers.DocumentBuilderFactory;

public class MailMergeUsingMustacheTemplateSyntax {

    public static void main(String[] args) throws Exception {
    	String dataDir = Utils.getSharedDataDir(MailMergeUsingMustacheTemplateSyntax.class) + "MailMerge/";
    	
        // Performs a simple insertion of data into merge fields and sends the document to the browser inline.
        simpleInsertionOfDataIntoMergeFields(dataDir);
        MustacheSyntaxUsingDataTable(dataDir);
        useMailMergeUsingMustacheSyntax(dataDir);
        UseOfIfElseMustacheSyntax(dataDir);
    }

    private static void MustacheSyntaxUsingDataTable(String dataDir) throws Exception {
		//ExStart: MustacheSyntaxUsingDataTable
		// Load a document
		Document doc = new Document(dataDir + "Test.docx");

		// Loop through each row and fill it with data
		DataTable dataTable = new DataTable("list");
		dataTable.getColumns().add("Number");
		for (int i = 0; i < 10; i++)
		{
		    DataRow datarow = dataTable.newRow();
		    dataTable.getRows().add(datarow);
		    datarow.set("Number " + i, i);
		}

		// Activate performing a mail merge operation into additional field types 
		doc.getMailMerge().setUseNonMergeFields(true);
		doc.getMailMerge().executeWithRegions(dataTable);
		doc.save(dataDir + "MailMerge.Mustache.docx");
		//ExEnd:MustacheSyntaxUsingDataTable
	}

    public static void simpleInsertionOfDataIntoMergeFields(String dataDir) throws Exception {
        // Open an existing document.
        Document doc = new Document(dataDir + "MailMerge.ExecuteArray.doc");

        doc.getMailMerge().setUseNonMergeFields(true);

        // Fill the fields in the document with user data.
        doc.getMailMerge().execute(new String[]{"FullName", "Company", "Address", "Address2", "City"}, new Object[]{"James Bond", "MI5 Headquarters", "Milbank", "", "London"});

        doc.save(dataDir + "MailMerge.ExecuteArray_Out.doc");
    }

    public static void useMailMergeUsingMustacheSyntax(String dataDir) throws Exception {
        // Use DocumentBuilder from the javax.xml.parsers package and Document class from the org.w3c.dom package to read
        // the XML data file and store it in memory.
        javax.xml.parsers.DocumentBuilder db = DocumentBuilderFactory.newInstance().newDocumentBuilder();

        // Parse the XML data.
        org.w3c.dom.Document xmlData = db.parse(dataDir + "Vendors.xml");

        // Open a template document.
        Document doc = new Document(dataDir + "VendorTemplate.doc");

        doc.getMailMerge().setUseNonMergeFields(true);
        // Note that this class also works with a single repeatable region (and any nested regions).
        // To merge multiple regions at the same time from a single XML data source, use the XmlMailMergeDataSet class.
        // e.g doc.getMailMerge().executeWithRegions(new XmlMailMergeDataSet(xmlData));
        doc.getMailMerge().executeWithRegions(new XmlMailMergeDataSet(xmlData));

        // Save the output document.
        doc.save(dataDir + "MailMergeUsingMustacheSyntax_Out.docx");
    }

    public static void UseOfIfElseMustacheSyntax(String dataDir) throws Exception {
        // ExStart:UseOfIfElseMustacheSyntax
        // Open a template document.
        Document doc = new Document(dataDir + "UseOfifelseMustacheSyntax.docx");

        doc.getMailMerge().setUseNonMergeFields(true);
        
        doc.getMailMerge().execute(new String[]{"GENDER"}, new Object[]{"MALE"});

        // Save the output document.
        doc.save(dataDir + "MailMergeUsingMustacheSyntaxifelse_out.docx");
        // ExEnd:UseOfIfElseMustacheSyntax
        System.out.println("\nMail merge performed with mustache if else syntax successfully.\nFile saved at " + dataDir);
    }
}