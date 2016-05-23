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

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.sql.Connection;
import java.util.stream.Stream;

public class LoadAndSaveToStream {
    private static Connection mConnection;

    public static void main(String[] args) throws Exception {
        // ExStart:LoadAndSave
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(LoadAndSaveToStream.class);
        String inputFile = "Test File (doc).doc";
        String outputFile = "output.png";

        InputStream in = new FileInputStream(dataDir + inputFile);
        OutputStream out = new FileOutputStream(dataDir + outputFile);

        Document doc = new Document(in);

        // Save the finished document to disk.
        doc.save(out, SaveFormat.PNG);
        System.out.println("Document loaded and saved successfully.");


        // ExEnd:LoadAndSave
    }
}

