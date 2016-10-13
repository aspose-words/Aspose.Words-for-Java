package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class Doc2PDF
{
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(Doc2PDF.class);

        // Load the document from disk.
        Document doc = new Document(dataDir + "Template.doc");

        // Save the document in PDF format.
        dataDir = dataDir + "output.pdf";
        doc.save(dataDir);

        System.out.println("\nDocument converted to PDF successfully.\nFile saved at " + dataDir);
    }
}
