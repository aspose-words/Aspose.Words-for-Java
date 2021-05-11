package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

public class ConvertDocumentToRtf {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ConvertDocumentToRtf.class);

        //ExStart:ConvertDocumentToRtf
        // Load a big document.
        Document doc = new Document(dataDir + "Word2003RTFSpec.doc");

        doc.save(dataDir + "Word2003RTFSpec_out.rtf");
        //ExEnd:ConvertDocumentToRtf
        System.out.println("Document converted to byte array successfully.");
    }
}
