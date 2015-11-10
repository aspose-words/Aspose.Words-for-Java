package com.aspose.words.examples.featurescomparison.ranges;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class AsposeDeleteRange
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeDeleteRange.class);

        Document doc = new Document(dataDir + "document.doc");
        doc.getSections().get(0).getRange().delete();

        String text = doc.getRange().getText();

        System.out.println("Range: " + text);
    }
}
