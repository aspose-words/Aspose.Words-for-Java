package com.aspose.words.examples.programming_documents.styles;

import com.aspose.words.CleanupOptions;
import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 7/12/2017.
 */
public class CleansUnusedStylesandLists {
    private static final String dataDir = Utils.getSharedDataDir(CleansUnusedStylesandLists.class) + "Styles/";

    public static void main(String[] args) throws Exception {

        //ExStart:CleansUnusedStylesandLists
        // Open the document.
        Document doc = new Document(dataDir + "TestFile.doc");
        CleanupOptions cleanupoptions = new CleanupOptions();

        cleanupoptions.setUnusedLists(false);
        cleanupoptions.setUnusedStyles(true);

        // Cleans unused styles and lists from the document depending on given CleanupOptions.
        doc.cleanup(cleanupoptions);
        doc.save(dataDir + "Document.Cleanup_out.docx");
        //ExEnd:CleansUnusedStylesandLists

        System.out.println("Document unused Styles cleaned successfully.");
    }

}
