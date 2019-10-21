package com.aspose.words.examples.quickstart;

import com.aspose.words.Document;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.examples.Utils;

import java.util.regex.Pattern;

public class FindAndReplace {
    private static final String dataDir = Utils.getSharedDataDir(FindAndReplace.class) + "FindAndReplace/";

    public static void main(String[] args) throws Exception {
        //ExStart:
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(FindAndReplace.class);

        // Open the document.
        Document doc = new Document(dataDir + "ReplaceSimple.doc");
        // Check the text of the document
        System.out.println("Original document text: " + doc.getRange().getText());
        Pattern regex = Pattern.compile("_CustomerName_", Pattern.CASE_INSENSITIVE);
        // Replace the text in the document.
        doc.getRange().replace(regex, "James Bond", new FindReplaceOptions());
        // Check the replacement was made.
        System.out.println("Document text after replace: " + doc.getRange().getText());
        // Save the modified document.
        doc.save(dataDir + "ReplaceSimple Out.doc");
        //ExEnd:

        System.out.println("Text found and replaced successfully.");
    }
}