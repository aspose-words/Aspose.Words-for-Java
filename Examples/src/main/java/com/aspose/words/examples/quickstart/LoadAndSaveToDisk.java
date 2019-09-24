package com.aspose.words.examples.quickstart;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class LoadAndSaveToDisk {
    public static void main(String[] args) throws Exception {
        //ExStart:LoadAndSaveToDisk
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(LoadAndSaveToDisk.class);
        //ExStart:OpenFromFile
        // Load the document from the absolute path on disk.
        Document doc = new Document(dataDir + "Document.doc");
        //ExEnd:OpenFromFile
        // Save the dDocument.dococument as DOCX document.");
        doc.save(dataDir + "Document Out.docx");
        //ExEnd:LoadAndSaveToDisk
        System.out.println("Document loaded from disk and saved again successfully.");
    }
}