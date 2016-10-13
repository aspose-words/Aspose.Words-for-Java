
package com.aspose.words.examples.programming_documents.Ranges;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;


public class RangesDeleteText
{
    private static String gDataDir;

    public static void main(String[] args) throws Exception
    {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RangesDeleteText.class);

        Document doc = new Document(dataDir + "Document.doc");
        doc.getSections().get(0).getRange().delete();

        doc.save(dataDir + "output.doc");

    }
}