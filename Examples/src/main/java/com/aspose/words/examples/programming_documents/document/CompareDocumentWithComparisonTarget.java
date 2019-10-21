package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.CompareOptions;
import com.aspose.words.ComparisonTargetType;
import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

import java.util.Date;

/**
 * Created by Home on 10/13/2017.
 */
public class CompareDocumentWithComparisonTarget {

    public static void main(String[] args) throws Exception {
        String dataDir = Utils.getDataDir(CompareDocumentWithComparisonTarget.class);
        // ExStart:CompareDocumentWithComparisonTarget
        Document docA = new Document(dataDir + "TestFile.doc");
        Document docB = new Document(dataDir + "TestFile - Copy.doc");

        CompareOptions options = new CompareOptions();
        options.setIgnoreFormatting(true);
        // Relates to Microsoft Word "Show changes in" option in "Compare Documents" dialog box.
        options.setTarget(ComparisonTargetType.NEW);

        docA.compare(docB, "user", new Date(), options);
        // ExEnd:CompareDocumentWithComparisonTarget

        System.out.println("\nDocuments have compared successfully.\nFile saved at " + dataDir);
    }
}
