package com.aspose.words.examples.programming_documents.sections;

import com.aspose.words.Document;
import com.aspose.words.Section;
import com.aspose.words.examples.Utils;

public class AppendSectionContent {
    public static void main(String[] args) throws Exception {
        //ExStart:AppendSectionContent
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SectionsAccessByIndex.class);
        Document doc = new Document(dataDir + "Section.AppendContent.doc");
        // This is the section that we will append and prepend to.
        Section section = doc.getSections().get(2);

        // This copies content of the 1st section and inserts it at the beginning of the specified section.
        Section sectionToPrepend = doc.getSections().get(0);
        section.prependContent(sectionToPrepend);

        // This copies content of the 2nd section and inserts it at the end of the specified section.
        Section sectionToAppend = doc.getSections().get(1);
        section.appendContent(sectionToAppend);
        //ExEnd:AppendSectionContent
        System.out.println("Section content appended successfully.");
    }
}