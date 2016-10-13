package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;
import com.aspose.words.*;

public class LoadEncryptedDoc
{
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(LoadEncryptedDoc.class);

        // Load the encrypted document from the absolute path on disk.
        Document doc = new Document(dataDir + "LoadEncrypted.docx", new LoadOptions("aspose"));

        System.out.println("Encrypted document loaded successfully.");
    }
}
