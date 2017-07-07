package com.aspose.words.examples.programming_documents.document;

import java.util.Date;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;



/**
 * Created by Home on 5/29/2017.
 */
public class CompareTwoWordDocumentswithCompareOptions {


    public static void main(String[] args) throws Exception {

        //ExStart:CompareTwoWordDocumentswithCompareOptions
		String dataDir = Utils.getDataDir(CompareTwoWordDocumentswithCompareOptions.class);

        com.aspose.words.Document docA = new com.aspose.words.Document(dataDir + "DocumentA.doc");
        com.aspose.words.Document docB = new com.aspose.words.Document(dataDir + "DocumentB.doc");

        com.aspose.words.CompareOptions options = new com.aspose.words.CompareOptions();
        options.setIgnoreFormatting(true);
        options.setIgnoreHeadersAndFooters(true);
        // DocA now contains changes as revisions.
        docA.compare(docB, "user", new Date(), options);
        if (docA.getRevisions().getCount() == 0)
            System.out.println("Documents are equal");
        else
            System.out.println("Documents are not equal");
		//ExEnd:CompareTwoWordDocumentswithCompareOptions
    }
}
