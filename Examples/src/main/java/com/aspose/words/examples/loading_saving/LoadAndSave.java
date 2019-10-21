package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

import java.sql.Connection;

public class LoadAndSave {
    private static Connection mConnection;

    public static void main(String[] args) throws Exception {
        //ExStart:LoadAndSave
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(LoadAndSave.class);
        String fileName = "Test File (doc).doc";
        // Load the document from disk.
        Document doc = new Document(dataDir + fileName);

        // Save the finished document to disk.
        doc.save(dataDir + "output.doc");
        //ExEnd:LoadAndSave
        System.out.println("Document loaded and saved successfully.");
    }
}

