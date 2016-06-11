/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
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