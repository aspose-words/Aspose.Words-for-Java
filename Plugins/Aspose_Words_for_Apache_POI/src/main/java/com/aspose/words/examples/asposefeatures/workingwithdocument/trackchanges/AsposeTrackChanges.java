package com.aspose.words.examples.asposefeatures.workingwithdocument.trackchanges;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

public class AsposeTrackChanges
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeTrackChanges.class);

        Document doc = new Document(dataDir +"trackDoc.doc");
        doc.acceptAllRevisions();
        doc.save(dataDir + "AsposeAcceptChanges.doc", SaveFormat.DOC);

        System.out.println("Done.");
    }
}
