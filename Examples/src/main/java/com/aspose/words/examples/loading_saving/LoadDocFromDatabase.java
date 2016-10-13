package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.LoadFormat;
import com.aspose.words.LoadOptions;
import com.aspose.words.SaveFormat;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

public class LoadDocFromDatabase {
    public static void main(String[] args) throws Exception {

        // Retrieve the blob from database
        byte[] buffer = new byte[100];
        // Now we have the document in a byte array buffer

        // Create an input steam which uses byte array to read data
        ByteArrayInputStream bin = new ByteArrayInputStream(buffer);
        // Open the doucment from input stream
        //Document doc = new Document(bin);
    }
}
