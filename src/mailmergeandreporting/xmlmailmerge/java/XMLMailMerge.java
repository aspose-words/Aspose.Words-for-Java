/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package mailmergeandreporting.xmlmailmerge.java;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.net.URI;

import com.aspose.words.*;

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
        String dataDir = "src/mailmergeandreporting/xmlmailmerge/data/";

        // Use DocumentBuilder from the javax.xml.parsers package and Document class from the org.w3c.dom package to read
        // the XML data file and store it in memory.
        DocumentBuilder db = DocumentBuilderFactory.newInstance().newDocumentBuilder();
        // Parse the XML data.
        org.w3c.dom.Document xmlData = db.parse(dataDir + "Customers.xml");

        // Open a template document.
        Document doc = new Document(dataDir + "TestFile.doc");

        // Note that this class also works with a single repeatable region (and any nested regions).
        // To merge multiple regions at the same time from a single XML data source, use the XmlMailMergeDataSet class.
        // e.g doc.getMailMerge().executeWithRegions(new XmlMailMergeDataSet(xmlData));
        doc.getMailMerge().execute(new XmlMailMergeDataTable(xmlData, "customer"));

        // Save the output document.
        doc.save(dataDir + "TestFile Out.doc");
    }
}
//ExEnd