/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.sections;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class DeleteSectionContent
{
    public static void main(String[] args) throws Exception
    {
        // ExStart:DeleteSectionContent
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SectionsAccessByIndex.class);
        Document doc = new Document(dataDir + "Document.doc");
        Section section = doc.getSections().get(0);
        section.clearContent();
        // ExEnd:DeleteSectionContent
        System.out.println("Section content at 0 index deleted successfully.");
    }
}