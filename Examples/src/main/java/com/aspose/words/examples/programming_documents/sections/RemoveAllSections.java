
package com.aspose.words.examples.programming_documents.sections;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class RemoveAllSections
{
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SectionsAccessByIndex.class);
        Document doc = new Document(dataDir + "Document.doc");
        doc.getSections().clear();

        System.out.println("All sections removed successfully form the document.");
    }
}