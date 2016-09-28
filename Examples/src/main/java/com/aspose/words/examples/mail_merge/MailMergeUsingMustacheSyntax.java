package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;
import com.aspose.words.net.System.Data.DataSet;


public class MailMergeUsingMustacheSyntax {
    public static void main(String[] args) throws Exception {

        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(MailMergeUsingMustacheSyntax.class);
        // Open the document.

        DataSet ds = new DataSet();
        ds.readXml(dataDir + "Vendors.xml");

        // Open a template document.
        Document doc = new Document(dataDir + "VendorTemplate.doc");

        doc.getMailMerge().setUseNonMergeFields(true);
        // Execute mail merge to fill the template with data from XML using DataSet.
        doc.getMailMerge().executeWithRegions(ds);

        // Save the output document to disk.
        doc.save(dataDir + "Output.doc");
        //ExEnd:1
    }
}