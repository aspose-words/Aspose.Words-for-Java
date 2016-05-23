/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
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
        // ExStart:LoadAndSave
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(LoadAndSave.class);
        String fileName = "Test File (doc).doc";
        // Load the document from disk.
        Document doc = new Document(dataDir + fileName);

        dataDir = dataDir + Utils.GetOutputFilePath(fileName);

        // Save the finished document to disk.
        doc.save(dataDir);
        System.out.println("Document loaded and saved successfully.");


        // ExEnd:LoadAndSave
    }
}

