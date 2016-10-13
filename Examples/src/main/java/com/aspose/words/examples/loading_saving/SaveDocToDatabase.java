package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;

import java.io.ByteArrayOutputStream;

public class SaveDocToDatabase {
    public static void main(String[] args) throws Exception {
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
    }
}
