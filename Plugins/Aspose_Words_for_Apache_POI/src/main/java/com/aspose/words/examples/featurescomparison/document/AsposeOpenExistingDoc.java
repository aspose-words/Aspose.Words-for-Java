package com.aspose.words.examples.featurescomparison.document;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

public class AsposeOpenExistingDoc
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeOpenExistingDoc.class);

        Document doc = new Document(dataDir + "document.doc");

        // Save the document in DOCX format.
        // Aspose.Words supports saving any document in many more formats.
        doc.save(dataDir + "Aspose_SaveDoc.docx",SaveFormat.DOCX);
		
        System.out.println("Process Completed Successfully");
    }
}
