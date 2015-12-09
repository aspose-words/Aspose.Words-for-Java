package com.aspose.words.examples.featurescomparison.documents.converttopdf;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

public class AsposeConvertToFormats
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeConvertToFormats.class);

        // Load the document from disk.
        Document doc = new Document(dataDir + "document.doc");

        doc.save(dataDir + "Aspose_DocToPDF.pdf",SaveFormat.PDF); //Save the document in PDF format.
        doc.save(dataDir + "html/Aspose_DocToHTML.html",SaveFormat.HTML); //Save the document in HTML format.
        doc.save(dataDir + "Aspose_DocToTxt.txt",SaveFormat.TEXT); //Save the document in TXT format.
        doc.save(dataDir + "Aspose_DocToJPG.jpg",SaveFormat.JPEG); //Save the document in JPEG format.

        System.out.println("Aspose - Doc file converted in specified formats");
    }
}
