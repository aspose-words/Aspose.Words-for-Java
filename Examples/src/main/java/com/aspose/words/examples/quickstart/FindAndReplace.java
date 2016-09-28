

package com.aspose.words.examples.quickstart;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class FindAndReplace
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(FindAndReplace.class);

        // Open the document.
        Document doc = new Document(dataDir + "ReplaceSimple.doc");
        // Check the text of the document
        System.out.println("Original document text: " + doc.getRange().getText());
        // Replace the text in the document.
        doc.getRange().replace("_CustomerName_", "James Bond", false, false);
        // Check the replacement was made.
        System.out.println("Document text after replace: " + doc.getRange().getText());
        // Save the modified document.
        doc.save(dataDir + "ReplaceSimple Out.doc");

        System.out.println("Text found and replaced successfully.");
    }
}