package com.aspose.words.examples.mail_merge;

import javax.xml.parsers.DocumentBuilderFactory;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class MailMergeUsingMustacheTemplateSyntax {

	private static final String dataDir = Utils.getSharedDataDir(MailMergeUsingMustacheTemplateSyntax.class) + "MailMerge/";

	public static void main(String[] args) throws Exception {
		// Performs a simple insertion of data into merge fields and sends the document to the browser inline.
		simpleInsertionOfDataIntoMergeFields();

		useMailMergeUsingMustacheSyntax();
	}

	public static void simpleInsertionOfDataIntoMergeFields() throws Exception {
		// Open an existing document.
		Document doc = new Document(dataDir + "MailMerge.ExecuteArray.doc");

		doc.getMailMerge().setUseNonMergeFields(true);

		// Fill the fields in the document with user data.
		doc.getMailMerge().execute(new String[] { "FullName", "Company", "Address", "Address2", "City" }, new Object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });

		doc.save(dataDir + "MailMerge.ExecuteArray_Out.doc");
	}

	public static void useMailMergeUsingMustacheSyntax() throws Exception {
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
}