

package com.aspose.words.examples.quickstart;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

public class LoadAndSaveToStream {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(LoadAndSaveToStream.class);
        // Open the stream. Read only access is enough for Aspose.Words to load a document.
        InputStream stream = new FileInputStream(dataDir + "Document.doc");
        // Load the entire document into memory.
        Document doc = new Document(stream);
        // You can close the stream now, it is no longer needed because the document is in memory.
        stream.close();

        // ... do something with the document
        // Convert the document to a different format and save to stream.
        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.RTF);
        FileOutputStream output = new FileOutputStream(dataDir + "Document Out.rtf");
        output.write(dstStream.toByteArray());
        output.close();

        System.out.println("Document loaded from stream and then saved successfully.");
    }
}