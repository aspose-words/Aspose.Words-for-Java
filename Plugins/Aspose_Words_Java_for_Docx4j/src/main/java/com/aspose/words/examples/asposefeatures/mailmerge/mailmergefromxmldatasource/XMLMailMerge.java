package com.aspose.words.examples.asposefeatures.mailmerge.mailmergefromxmldatasource;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

/**
 * This sample demonstrates how to execute mail merge with data from an XML data source. The XML file is read into memory,
 * stored in a DOM and passed to a custom data source implementing IMailMergeDataSource. This returns each value from XML when
 * called by the mail merge engine.
 */
public class XMLMailMerge
{
    public static void main(String[] args) throws Exception
    {
    	// The path to the documents directory.
        String dataDir = Utils.getDataDir(XMLMailMerge.class);
    	
        // Use DocumentBuilder from the javax.xml.parsers package and Document class from the org.w3c.dom package to read
        // the XML data file and store it in memory.
        DocumentBuilder db = DocumentBuilderFactory.newInstance().newDocumentBuilder();
        
        // Parse the XML data.
        org.w3c.dom.Document xmlData = db.parse(dataDir + "Customers.xml");

        // Open a template document.
        Document doc = new Document(dataDir + "mergeDoc.doc");

        // Note that this class also works with a single repeatable region (and any nested regions).
        // To merge multiple regions at the same time from a single XML data source, use the XmlMailMergeDataSet class.
        // e.g doc.getMailMerge().executeWithRegions(new XmlMailMergeDataSet(xmlData));
        doc.getMailMerge().execute(new XmlMailMergeDataTable(xmlData, "customer"));

        // Save the output document.
        doc.save(dataDir + "AsposeMailMerge.doc");
    }
}
//ExEnd