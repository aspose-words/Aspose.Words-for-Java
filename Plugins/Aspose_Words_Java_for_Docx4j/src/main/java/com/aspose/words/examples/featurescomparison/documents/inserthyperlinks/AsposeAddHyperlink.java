package com.aspose.words.examples.featurescomparison.documents.inserthyperlinks;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;

public class AsposeAddHyperlink
{
    public static void main(String[] args) throws Exception 
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeAddHyperlink.class);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Please make sure to visit ");
        // Insert the link.
        builder.insertHyperlink("Aspose Website", "http://www.aspose.com", false);

        doc.save(dataDir + "AsposeAddHyperlinks.doc");
        System.out.println("Done.");
    }
}
