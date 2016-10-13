
package com.aspose.words.examples.programming_documents.sections;

import com.aspose.words.Document;
import com.aspose.words.Section;
import com.aspose.words.examples.Utils;

public class AddSection
{
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SectionsAccessByIndex.class);
        Document doc = new Document(dataDir + "Document.doc");
        Section sectionToAdd = new Section(doc);
        doc.getSections().add(sectionToAdd);

        System.out.println("Section added successfully to the end of the document.");
    }
}