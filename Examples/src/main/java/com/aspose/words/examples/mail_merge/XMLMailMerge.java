package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;
import com.aspose.words.net.System.Data.DataSet;
//ExStart:

/**
 * This sample demonstrates how to execute mail merge with data from an XML data
 * source. The XML file is read into memory, stored in a DOM and passed to a
 * custom data source implementing IMailMergeDataSource. This returns each value
 * from XML when called by the mail merge engine.
 */
public class XMLMailMerge {

    public static void main(String[] args) throws Exception {
    	//ExStart: XMLMailMerge
    	// The path to the documents directory.
    	String dataDir = Utils.getSharedDataDir(XMLMailMerge.class) + "MailMerge/";

    	// Create the Dataset and read the XML.
    	DataSet customersDs = new DataSet();
    	customersDs.readXml(dataDir + "Customers.xml");

    	String fileName = "TestFile XML.doc";
    	// Open a template document.
    	Document doc = new Document(dataDir + fileName);

    	// Execute mail merge to fill the template with data from XML using DataTable.
    	// Note that this class also works with a single repeatable region (and any nested regions).
    	// To merge multiple regions at the same time from a single XML data source, use the XmlMailMergeDataSet class.
    	// e.g doc.getMailMerge().executeWithRegions(new XmlMailMergeDataSet(xmlData));
    	doc.getMailMerge().execute(customersDs.getTables().get("Customer"));

    	// Save the output document.
    	doc.save(dataDir + fileName);
    	//ExEnd: XMLMailMerge
        System.out.println("Mail merge performed successfully.");
    }
}
//ExEnd: