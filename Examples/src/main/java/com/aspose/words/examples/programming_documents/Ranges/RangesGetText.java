
package com.aspose.words.examples.programming_documents.Ranges;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.examples.Utils;


public class RangesGetText
{
    private static String gDataDir;

    public static void main(String[] args) throws Exception
    {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RangesGetText.class);

        Document doc = new Document(dataDir + "Document.doc");
        String text = doc.getText();
        System.out.println(text);

    }
}