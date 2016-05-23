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
import com.aspose.words.*;
import java.nio.charset.Charset;

public class ConvertDocumentToEPUB
{
    public static void main(String[] args) throws Exception
    {
        // ExStart:ConvertDocumentToEPUB
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ConvertDocumentToEPUB.class);

        // Open an existing document from disk.
        Document doc = new Document(dataDir + "Document.EpubConversion.doc");

        // Save the document in EPUB format.
        doc.save(dataDir + "Document.EpubConversion_out_.epub");
        // ExEnd:ConvertDocumentToEPUB
        System.out.println("Document converted to EPUB successfully.");



    }
}
