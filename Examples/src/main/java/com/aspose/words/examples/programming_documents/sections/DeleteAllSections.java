package com.aspose.words.examples.programming_documents.sections;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class DeleteAllSections {
    public static void main(String[] args) throws Exception {
        //ExStart:DeleteAllSections
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DeleteAllSections.class);
        Document doc = new Document(dataDir + "Document.doc");
        doc.getSections().clear();
        doc.save(dataDir + "output.doc");
        //ExEnd:DeleteAllSections
        System.out.println("All sections deleted successfully form the document.");
    }
}