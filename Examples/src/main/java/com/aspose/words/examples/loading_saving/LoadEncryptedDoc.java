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

public class LoadEncryptedDoc
{
    public static void main(String[] args) throws Exception
    {
        // ExStart:LoadEncryptedDocument
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(LoadEncryptedDoc.class);

        // Load the encrypted document from the absolute path on disk.
        Document doc = new Document(dataDir + "LoadEncrypted.docx", new LoadOptions("aspose"));
        // ExEnd:LoadEncryptedDocument

        System.out.println("Encrypted document loaded successfully.");
    }
}
