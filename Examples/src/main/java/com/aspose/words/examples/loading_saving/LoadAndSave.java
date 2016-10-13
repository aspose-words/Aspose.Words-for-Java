package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.sql.*;
import java.text.MessageFormat;

public class LoadAndSave {
    private static Connection mConnection;

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(LoadAndSave.class);
        String fileName = "Test File (doc).doc";
        // Load the document from disk.
        Document doc = new Document(dataDir + fileName);

        // Save the finished document to disk.
        doc.save(dataDir + "output.doc");
        System.out.println("Document loaded and saved successfully.");
    }
}

