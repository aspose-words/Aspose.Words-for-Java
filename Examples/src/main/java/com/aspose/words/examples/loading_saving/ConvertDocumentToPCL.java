package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.PclSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 8/10/2017.
 */
public class ConvertDocumentToPCL {

    public static void main(String[] args) throws Exception {

        //ExStart:ConvertDocumentToPCL
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ConvertDocumentToPCL.class);

        // Load the document from disk.
        Document doc = new Document(dataDir + "Document.doc");

        PclSaveOptions saveOptions = new PclSaveOptions();

        saveOptions.setSaveFormat(SaveFormat.PCL);
        saveOptions.setRasterizeTransformedElements(false);

        // Export the document as an PCL file.
        doc.save(dataDir + "Document.PclConversion_out.pcl", saveOptions);
        //ExEnd:ConvertDocumentToPCL

        System.out.println("Document converted to PCL successfully.");
    }
}
