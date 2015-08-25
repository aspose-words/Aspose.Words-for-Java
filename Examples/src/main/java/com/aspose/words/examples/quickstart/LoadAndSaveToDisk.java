/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
  
package com.aspose.words.examples.quickstart;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class LoadAndSaveToDisk
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(LoadAndSaveToDisk.class);

        // Load the document from the absolute path on disk.
        Document doc = new Document(dataDir + "Document.doc");
        // Save the dDocument.dococument as DOCX document.");
        doc.save(dataDir + "Document Out.docx");

        System.out.println("Document loaded from disk and saved again successfully.");
    }
}