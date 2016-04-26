/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.sections;

import com.aspose.words.Document;
import com.aspose.words.Section;
import com.aspose.words.examples.Utils;

public class AddSection
{
    public static void main(String[] args) throws Exception
    {
        // ExStart:AddSection
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SectionsAccessByIndex.class);
        Document doc = new Document(dataDir + "Document.doc");
        Section sectionToAdd = new Section(doc);
        doc.getSections().add(sectionToAdd);
        // ExEnd:AddSection
        System.out.println("Section added successfully to the end of the document.");

    }
}