/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class Doc2PDF
{
    public static void main(String[] args) throws Exception
    {
        // ExStart:Doc2PDF
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(Doc2PDF.class);

        // Load the document from disk.
        Document doc = new Document(dataDir + "Template.doc");

        // Save the document in PDF format.
        dataDir = dataDir + "Template_out_.pdf";
        doc.save(dataDir);
        // ExEnd:Doc2PDF
        System.out.println("\nDocument converted to PDF successfully.\nFile saved at " + dataDir);
    }
}
