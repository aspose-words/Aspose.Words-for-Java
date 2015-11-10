package com.aspose.words.examples.featurescomparison.ranges;

import com.aspose.words.Document;
import com.aspose.words.Range;
import com.aspose.words.examples.Utils;

public class AsposeRanges
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeRanges.class);

        Document doc = new Document(dataDir + "document.doc");
        Range range = doc.getRange();

        String text = range.getText();
        System.out.println("Range: " + text);
    }
}
