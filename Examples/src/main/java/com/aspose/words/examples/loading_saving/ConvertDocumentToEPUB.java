package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;
import com.aspose.words.*;
import java.nio.charset.Charset;

public class ConvertDocumentToEPUB
{
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ConvertDocumentToEPUB.class);

        // Open an existing document from disk.
        Document doc = new Document(dataDir + "Document.EpubConversion.doc");

        // Save the document in EPUB format.
        doc.save(dataDir + "Document.EpubConversion_out_.epub");

        System.out.println("Document converted to EPUB successfully.");
    }
}
