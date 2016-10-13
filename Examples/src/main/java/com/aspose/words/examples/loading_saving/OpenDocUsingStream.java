package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

import java.io.FileInputStream;
import java.io.InputStream;

public class OpenDocUsingStream {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(OpenDocUsingStream.class);
        String filename = "Test.docx";

        InputStream in = new FileInputStream(dataDir + filename);

        Document doc = new Document(in);
        System.out.println("Document opened. Total pages are " + doc.getPageCount());
        in.close();
    }
}
