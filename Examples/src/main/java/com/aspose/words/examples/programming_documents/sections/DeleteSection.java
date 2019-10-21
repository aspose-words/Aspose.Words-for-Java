package com.aspose.words.examples.programming_documents.sections;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class DeleteSection {
    public static void main(String[] args) throws Exception {
        //ExStart:DeleteSection
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SectionsAccessByIndex.class);
        Document doc = new Document(dataDir + "Document.doc");
        doc.getSections().removeAt(0);
        //ExEnd:DeleteSection

        System.out.println("Section deleted successfully at 0 index.");
    }
}