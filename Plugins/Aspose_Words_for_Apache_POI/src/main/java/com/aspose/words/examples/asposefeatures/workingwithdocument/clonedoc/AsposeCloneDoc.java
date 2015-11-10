package com.aspose.words.examples.asposefeatures.workingwithdocument.clonedoc;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

public class AsposeCloneDoc
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeCloneDoc.class);

        Document doc = new Document(dataDir + "document.doc");
        Document clone = doc.deepClone();
        clone.save(dataDir + "AsposeClone.doc", SaveFormat.DOC);

        System.out.println("Done.");
    }
}