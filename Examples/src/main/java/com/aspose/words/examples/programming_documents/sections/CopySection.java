package com.aspose.words.examples.programming_documents.sections;

import com.aspose.words.Document;
import com.aspose.words.Section;
import com.aspose.words.examples.Utils;

public class CopySection  {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SectionsAccessByIndex.class);
        Document srcDoc = new Document(dataDir + "Document.doc");
        Document dstDoc = new Document();

        Section sourceSection = srcDoc.getSections().get(0);
        Section newSection = (Section)dstDoc.importNode(sourceSection, true);
        dstDoc.getSections().add(newSection);

        dstDoc.save(dataDir+ "output.doc");
        System.out.println("\nSection copied successfully.\nFile saved at " + dataDir);
    }
}