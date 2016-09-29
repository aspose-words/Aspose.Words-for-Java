package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;
import javax.xml.parsers.DocumentBuilderFactory;

public class MustacheTemplateSyntax {

	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(MustacheTemplateSyntax.class);

		// Use DocumentBuilder from the javax.xml.parsers package and Document class from the org.w3c.dom package to read
		// the XML data file and store it in memory.
		javax.xml.parsers.DocumentBuilder db = DocumentBuilderFactory.newInstance().newDocumentBuilder();

		// Parse the XML data.
		org.w3c.dom.Document xmlData = db.parse(dataDir + "Orders.xml");

		// Open a template document.
		Document doc = new Document(dataDir + "ExecuteTemplate.doc");

		doc.getMailMerge().setUseNonMergeFields(true);

		// Note that this class also works with a single repeatable region (and any nested regions).
		// To merge multiple regions at the same time from a single XML data source, use the XmlMailMergeDataSet class.
		// e.g doc.getMailMerge().executeWithRegions(new XmlMailMergeDataSet(xmlData));
		doc.getMailMerge().executeWithRegions(new XmlMailMergeDataSet(xmlData));

		// Save the output document.
		doc.save(dataDir + "Output.docx");

		System.out.println("Mail merge performed successfully.");
	}

}