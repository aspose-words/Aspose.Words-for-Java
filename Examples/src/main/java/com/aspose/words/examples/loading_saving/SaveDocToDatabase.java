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

import java.io.ByteArrayOutputStream;

public class SaveDocToDatabase {
    public static void main(String[] args) throws Exception {
        // ExStart:1
        // Create a new empty document
        Document doc = new Document();
        // Create an output stream which uses byte array to save data
        ByteArrayOutputStream aout = new ByteArrayOutputStream();
        // Save the document to byte array
        doc.save(aout, SaveFormat.DOCX);
        // Get the byte array from output steam
        // the byte array now contains the document
        byte[] buffer = aout.toByteArray();
        // Save the document to database blob
        // ExEnd:1
    }
}
