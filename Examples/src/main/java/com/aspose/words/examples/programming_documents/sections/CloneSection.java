package com.aspose.words.examples.programming_documents.sections;

import com.aspose.words.Document;
import com.aspose.words.Section;
import com.aspose.words.examples.Utils;

public class CloneSection {
    public static void main(String[] args) throws Exception {

        //ExStart:CloneSection
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SectionsAccessByIndex.class);
        Document doc = new Document(dataDir + "Document.doc");
        Section cloneSection = doc.getSections().get(0).deepClone();
        //ExEnd:CloneSection
        System.out.println("0 index section clone successfully.");
    }
}