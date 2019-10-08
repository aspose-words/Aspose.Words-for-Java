package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

public class ConvertDocumentToByte {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ConvertDocumentToByte.class);

        //ExStart:ConvertDocumentToByte
        // Load the document.
        Document doc = new Document(dataDir + "Test File (doc).doc");

        // Create a new memory stream.
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        // Save the document to stream.
        doc.save(outStream, SaveFormat.DOCX);

        // Convert the document to byte form.
        byte[] docBytes = outStream.toByteArray();

        // The bytes are now ready to be stored/transmitted.

        // Now reverse the steps to load the bytes back into a document object.
        ByteArrayInputStream inStream = new ByteArrayInputStream(docBytes);

        // Load the stream into a new document object.
        Document loadDoc = new Document(inStream);
        //ExEnd:ConvertDocumentToByte
        System.out.println("Document converted to byte array successfully.");
    }
}
